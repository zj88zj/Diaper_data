proc datasets library=work kill nolist; 
quit; 
%symdel; title; footnote; 
options nofmterr ls = 180 ps=80 mprint source  nodate pageno=1 ;
dm 'log;clear;output;clear;odsresults;select all;clear;'; 

%let path =Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2;
libname der 'Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Derived Data';



proc import out = bm1 datafile= "&path\BM Study October 2019 Raw RGB Data.xlsx" dbms = excel replace; range = "Raw Data$A1:IT486"; run;
proc import out = bm2 datafile= "&path\BM Study October 2019 Raw RGB Data.xlsx" dbms = excel replace; range = "Raw Data$IU1:IY486"; run;

data bm;
merge bm1 bm2;
run;

data use;
 set bm;
 array rval(80) r1-r80;
 array gval(80) g1-g80;
 array bval(80) b1-b80;
  do spot = 1 to 80;
    r = rval(spot);
	g = gval(spot);
	b = bval(spot);
	output;
  end;
 drop r1-r80 g1-g80 b1-b80;
run;

/*data use;*/
/*set use;*/
/*where r < 100000 and g < 100000 and b < 100000;*/
/*    age = input(baby_age, $12.);*/
/*    size = input(diaper_size, $12.);*/
/*    device = input(device__1_2_3_, $12.);*/
/*    array = input(array__, $12.);*/
/*    bm_urine = lowcase(bm_urine);*/
/*    panelist = input(panelist__, $12.);*/
/*if panelist__ = . and bm_urine = 'blank' then panelist = '000000';*/
/*if panelist__ = . and bm_urine = 'clean diaper' then panelist = '111111';*/
/*if bm_urine = 'blank' then tube__ = 'blank';*/
/*    array change(8) BM_Consistency BM_Size baby_gender diet age size device array;*/
/*       do i = 1 to 8;*/
/*          if bm_urine in ('blank', 'clean diaper') then change(i) = 'blank';*/
/*       end;*/
/*drop baby_age diaper_size device__1_2_3_  array__ i panelist__;*/
/*run;*/

/*data der.trans_bm; set use; run;*/

data use;
set use;
where r < 100000 and g < 100000 and b < 100000 and bm_urine ^= 'blank';
    bm_urine = lowcase(bm_urine);
    panelist = input(panelist__, $12.);
if panelist__ = . and bm_urine = 'clean diaper' then panelist = '111111';
drop panelist__;
run;



/****************************************************************/
proc sort data= use; by spot;run;
proc mixed data = use order=data;
  by spot;
  class panelist bm_urine;
  model r = bm_urine/ ddfm=kr;
  random intercept / subject=panelist;
  lsmeans bm_urine / pdiff;
  ods output lsmeans=ls_r diffs=diffs_r tests3=tests_r;
run;
proc mixed data = use order=data;
  by spot;
  class panelist bm_urine;
  model g = bm_urine/ ddfm=kr;
  random intercept / subject=panelist;
  lsmeans bm_urine / pdiff;
  ods output lsmeans=ls_g diffs=diffs_g tests3=tests_g;
run;
proc mixed data = use order=data;
  by spot;
  class panelist bm_urine;
  model b = bm_urine/ ddfm=kr;
  random intercept / subject=panelist;
  lsmeans bm_urine / pdiff;
  ods output lsmeans=ls_b diffs=diffs_b tests3=tests_b;
run;

proc sort data= ls_r; by spot bm_urine;run;
proc sort data= ls_g; by spot bm_urine;run;
proc sort data= ls_b; by spot bm_urine;run;
ods _all_ close;
ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\lsm estimates plots" ;
ods graphics on/ reset imagename = "r estimates series";
proc sgpanel data=ls_r;
 panelby spot / columns=5 rows=2;
 series x=bm_urine y=estimate;
 run; quit;

ods graphics on/ reset imagename = "g estimates series";
proc sgpanel data=ls_g;
 panelby spot / columns=5 rows=2;
 series x=bm_urine y=estimate;
 run; quit;

ods graphics on/ reset imagename = "b estimates series";
proc sgpanel data=ls_b;
 panelby spot / columns=5 rows=2;
 series x=bm_urine y=estimate;
 run; quit;
ods graphics off;


data means;
 merge ls_r (rename=(estimate=r)) ls_g (rename=(estimate=g)) ls_b (rename=(estimate=b));
 by spot bm_urine;
 keep spot bm_urine r g b;
 run;

ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\means plots.rtf" style=statistical;

proc sgpanel data=means;
 panelby spot / columns=4 rows=2;
 series x=bm_urine y=b;
 series x=bm_urine y=r;
 series x=bm_urine y=g;
 run; quit;
ods rtf close;



/****************************************************************/
*** Table of R results ***;
data diffs_r;
 set diffs_r;
 where BM_Urine = 'bm';
 run;

proc sort data = diffs_r; by spot _bm_urine; run;
data rtable;
  merge ls_r(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine)) 
        diffs_r(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  upper=estimate+stderr;
  lower=estimate-stderr;
  if probt<0.05 and probt^=. then label=upper*1.15;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\R Tables.rtf"; 
%rtftable(%str( 

proc rtftable data=rtable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "R Summary by Spot";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean R" stderr="<w=1>Standard Error"
        diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; 
)); 
ods rtf close;

*** Table of G results ***;
data diffs_g;
 set diffs_g;
 where BM_Urine = 'bm';
 run;

 proc sort data = diffs_g; by spot _bm_urine; run;
data gtable;
  merge ls_g(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine)) 
        diffs_g(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  upper=estimate+stderr;
  lower=estimate-stderr;
  if probt<0.05 and probt^=. then label=upper*1.15;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\G Tables.rtf"; 
%rtftable(%str( 
proc rtftable data=gtable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "G Summary by Spot";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean G" stderr="<w=1>Standard Error"
        diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; 
)); 
ods rtf close;

*** Table of B results ***;
data diffs_b;
 set diffs_b;
 where BM_Urine = 'bm';
 run;

 proc sort data = diffs_b; by spot _bm_urine; run;
data btable;
  merge ls_b(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine))  
        diffs_b(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  upper=estimate+stderr;
  lower=estimate-stderr;
  if probt<0.05 and probt^=. then label=upper*1.15;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\B Tables.rtf"; 
%rtftable(%str( 
proc rtftable data=btable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "B Summary by Spot";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean B" stderr="<w=1>Standard Error"
        diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; 
)); 
ods rtf close;




/**********************************************************************/

** heatmaps - significance of bm_urine_clean;

data tests_r;
 set tests_r;
 if probf < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data tests_g;
 set tests_g;
 if probf < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data tests_b;
 set tests_b;
 if probf < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

data tests;
set tests_r(in = a) tests_g(in = b) tests_b(in = c);
if a then color = 'r';
if b then color = 'g';
if c then color = 'b';
run;


ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results" dpi =300;
ods graphics on/ reset imagename = 'R heatmap' border = off noscale width = 2in height = 15in;
proc sgplot data=tests_r;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

ods graphics on/ reset imagename = 'G heatmap' border = off noscale width = 2in height = 15in;
proc sgplot data=tests_g;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run; 

 ods graphics on/ reset imagename = 'B heatmap' border = off noscale width = 2in height = 15in;
proc sgplot data=tests_b;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

 ods graphics on/ reset imagename = 'combine_heatmap' border = off noscale width = 5in height = 15in;
 proc sgpanel data = tests;
 panelby color/ novarname columns=3;
 heatmap x = effect y = spot/ colorresponse= pvalue ybinsize = 1 outline colormodel=(blue white);
 rowaxis values = (1 to 81 by 1);
 run;
ods graphics off;


** Significant spots: 1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78;
** from previous study significant spots
                       32: r
                       55: b, g
                       66: none
                       67: b
;



** heatmaps - pairedwise significance - compare bm_clean;

data diffs_r_1;
 set diffs_r (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_g_1;
 set diffs_g (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_b_1;
 set diffs_b (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

data diffs;
set diffs_r_1(in = a) diffs_g_1(in = b) diffs_b_1(in = c);
if a then color = 'r';
if b then color = 'g';
if c then color = 'b';
run;

ods _all_ close;
ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results" dpi =300;
ods graphics on/ reset imagename = 'R heatmap - BM-Clean' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_r_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

ods graphics on/ reset imagename = 'G heatmap - BM-Clean' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_g_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run; 

 ods graphics on/ reset imagename = 'B heatmap - BM-Clean' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_b_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

 ods graphics on/ reset imagename = 'combine_heatmap - BM-Clean' border = off noscale width = 5in height = 15in;
 proc sgpanel data = diffs;
 panelby color/ novarname columns=3;
 heatmap x = effect y = spot/ colorresponse= pvalue ybinsize = 1 outline colormodel=(blue white);
 rowaxis values = (1 to 81 by 1);
 run;
ods graphics off;


** All Significant spots: none;
** from previous study significant spots
                       32: none
                       55: g
                       66: r
                       67: none
;
data diffs_r_1;
 set diffs_r (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05 and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_g_1;
 set diffs_g (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05  and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_b_1;
 set diffs_b (where = (_bm_urine = 'clean diaper'));
 effect = 'BM-Clean';
 if probt < 0.05  and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

data diffs;
set diffs_r_1(in = a) diffs_g_1(in = b) diffs_b_1(in = c);
if a then color = 'r';
if b then color = 'g';
if c then color = 'b';
run;

ods _all_ close;
ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results" dpi =300;
ods graphics on/ reset imagename = 'R heatmap - BM-Clean - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_r_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

ods graphics on/ reset imagename = 'G heatmap - BM-Clean - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_g_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run; 

 ods graphics on/ reset imagename = 'B heatmap - BM-Clean - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_b_1;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

 ods graphics on/ reset imagename = 'combine_heatmap - BM-Clean - bm' border = off noscale width = 5in height = 15in;
 proc sgpanel data = diffs;
 panelby color/ novarname columns=3;
 heatmap x = effect y = spot/ colorresponse= pvalue ybinsize = 1 outline colormodel=(blue white);
 rowaxis values = (1 to 81 by 1);
 run;
ods graphics off;


** All Significant spots: none;
** from previous study significant spots
                       32: none
                       55: none
                       66: none
                       67: none
;






** heatmaps - pairedwise significance - compare bm_urine;

data diffs_r_2;
 set diffs_r (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_g_2;
 set diffs_g (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_b_2;
 set diffs_b (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05 then pvalue=0.05;
  else pvalue=0.8;
  run;

data diffs2;
set diffs_r_2(in = a) diffs_g_2(in = b) diffs_b_2(in = c);
if a then color = 'r';
if b then color = 'g';
if c then color = 'b';
run;

ods _all_ close;
ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results" dpi =300;
ods graphics on/ reset imagename = 'R heatmap - BM-Urine' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_r_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

ods graphics on/ reset imagename = 'G heatmap - BM-Urine' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_g_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run; 

 ods graphics on/ reset imagename = 'B heatmap - BM-Urine' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_b_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

 ods graphics on/ reset imagename = 'combine_heatmap - BM-Urine' border = off noscale width = 5in height = 15in;
 proc sgpanel data = diffs2;
 panelby color/ novarname columns=3;
 heatmap x = effect y = spot/ colorresponse= pvalue ybinsize = 1 outline colormodel=(blue white);
 rowaxis values = (1 to 81 by 1);
 run;
ods graphics off;


** All Significant spots: 1, 12, 22, 39, 41, 46, 50, 60, 62, 63, 69, 78;
** from previous study significant spots
                       32: b, r
                       55: b
                       66: none
                       67: b
;
data diffs_r_2;
 set diffs_r (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05  and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_g_2;
 set diffs_g (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05 and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

  data diffs_b_2;
 set diffs_b (where = (_bm_urine = 'urine'));
 effect = 'BM-Urine';
 if probt < 0.05 and spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) then pvalue=0.05;
  else pvalue=0.8;
  run;

data diffs2;
set diffs_r_2(in = a) diffs_g_2(in = b) diffs_b_2(in = c);
if a then color = 'r';
if b then color = 'g';
if c then color = 'b';
run;

ods _all_ close;
ods listing gpath = "Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results" dpi =300;
ods graphics on/ reset imagename = 'R heatmap - BM-Urine - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_r_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

ods graphics on/ reset imagename = 'G heatmap - BM-Urine - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_g_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run; 

 ods graphics on/ reset imagename = 'B heatmap - BM-Urine - bm' border = off noscale width = 2in height = 15in;
proc sgplot data=diffs_b_2;
 heatmap x=effect y=spot / colorresponse=pvalue outline nybins=80 colormodel=(blue white);
 yaxis values = (1 to 81 by 1);
 run;

 ods graphics on/ reset imagename = 'combine_heatmap - BM-Urine - bm' border = off noscale width = 5in height = 15in;
 proc sgpanel data = diffs2;
 panelby color/ novarname columns=3;
 heatmap x = effect y = spot/ colorresponse= pvalue ybinsize = 1 outline colormodel=(blue white);
 rowaxis values = (1 to 81 by 1);
 run;
ods graphics off;


** All Significant spots: 1, 12, 22, 39, 41, 46, 50, 60, 62, 69, 78;
** from previous study significant spots
                       32: none
                       55: none
                       66: none
                       67: none
;





/************************** Reduced table - Deltas >= 50 & significant spots **********************************/

/*R*/
proc sql noprint;
select spot into :r_spot_sig separated by ' ' from tests_r where probf <0.05;
quit;

proc sql noprint;
select spot into :r_spot_delta separated by ' ' 
from diffs_r group by spot having count(case when abs(estimate) >= 50 then estimate end) >= 1 ;  ** Delta >= 50 either BM-Urine or BM-Clean;
quit;

proc sort data = diffs_r; by spot _bm_urine; run;
data rtable;
  merge ls_r(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine))  
        diffs_r(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  where spot in (&r_spot_sig) and spot in (&r_spot_delta);
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\R Tables Reduced.rtf"; 
%rtftable(%str( 
proc rtftable data=rtable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "R Summary by Spot - Reduced (Delta >= 50 & Significant)";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean R" stderr="<w=1>Standard Error" diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; )); 
ods rtf close;

/*G*/
proc sql noprint;
select spot into :g_spot_sig separated by ' ' from tests_g where probf <0.05;
quit;

proc sql noprint;
select spot into :g_spot_delta separated by ' ' 
from diffs_g group by spot having count(case when abs(estimate) >= 50 then estimate end) >= 1 ;  ** Delta >= 50 either BM-Urine or BM-Clean;
quit;

proc sort data = diffs_g; by spot _bm_urine; run;
data gtable;
  merge ls_g(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine))  
        diffs_g(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  where spot in (&g_spot_sig) and spot in (&g_spot_delta);
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\G Tables Reduced.rtf"; 
%rtftable(%str( 
proc rtftable data=gtable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "G Summary by Spot - Reduced (Delta >= 50 & Significant)";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean G" stderr="<w=1>Standard Error" diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; )); 
ods rtf close;

/*B*/
proc sql noprint;
select spot into :b_spot_sig separated by ' ' from tests_b where probf <0.05;
quit;

proc sql noprint;
select spot into :b_spot_delta separated by ' ' 
from diffs_b group by spot having count(case when abs(estimate) >= 50 then estimate end) >= 1 ;  ** Delta >= 50 either BM-Urine or BM-Clean;
quit;

proc sort data = diffs_b; by spot _bm_urine; run;
data btable;
  merge ls_b(keep=spot bm_urine estimate stderr rename = (bm_urine = _bm_urine))  
        diffs_b(keep=spot _bm_urine estimate probt rename=(estimate=diff));
  by spot _bm_urine;
  where spot in (&b_spot_sig) and spot in (&b_spot_delta);
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\B Tables Reduced.rtf"; 
%rtftable(%str( 
proc rtftable data=btable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "B Summary by Spot - Reduced (Delta >= 50 & Significant)";
  break spot; 
  id _bm_urine; 
  var estimate stderr ("Comparison to BM" diff probt);
  label trt='<w=1.5>Condition' estimate="<w=1>Mean B" stderr="<w=1>Standard Error" diff="<w=1>Mean Difference" probt="<w=1>P-value";
run; )); 
ods rtf close;


/*Across RGB*/
data rgbtable;
merge rtable(keep = spot _bm_urine estimate diff probt rename = (estimate = r_mean diff = r_diff probt = r_pvalue) in = a)
      gtable(keep = spot _bm_urine estimate diff probt rename = (estimate = g_mean diff = g_diff probt = g_pvalue) in = b)
	  btable(keep = spot _bm_urine estimate diff probt rename = (estimate = b_mean diff = b_diff probt = b_pvalue) in = c);
by spot _bm_urine;
if a and b and c;
where spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78);
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\RGB Tables Reduced.rtf"; 
%rtftable(%str( 
proc rtftable data=rgbtable; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "RGB Summary by Spot - Reduced (Delta >= 50 & Significant across RGB)";
  break spot; 
  id _bm_urine; 
  var ("Means" r_mean g_mean b_mean) ("Comparison to BM" r_diff g_diff b_diff r_pvalue g_pvalue b_pvalue);
  label r_mean="<w=0.6>R" g_mean="<w=0.6>G" b_mean ="<w=0.6>B" r_diff="<w=0.8>R |Difference" g_diff="<w=0.8>G |Difference" b_diff="<w=0.8>B |Difference" r_pvalue="<w=0.7>R |P-value" g_pvalue="<w=0.7>G |P-value" b_pvalue="<w=0.7>B |P-value";
  format r_mean g_mean b_mean r_diff g_diff b_diff 5.1  r_pvalue g_pvalue b_pvalue pvalue7.3;
run; )); 
ods rtf close;






/************************** subset analysis - BM **********************************/
data subset;
set use;
where spot in (1, 12, 15, 22, 39, 41, 43, 46, 50, 54, 60, 62, 69, 74, 78) and bm_urine = 'bm';
run;


/*R*/
** not significant effects:   baby_age  diet ;
** Significant effects:   BM_Consistency BM_Size Baby_Gender Diaper_Size jar_nojar ;

proc sort data = subset; by spot;
proc mixed data = subset order=data;
  by spot;
  class panelist bm_consistency bm_size jar_nojar baby_gender  diaper_size;
  model r = bm_consistency bm_size jar_nojar baby_gender diaper_size/ ddfm=kr; 
  random intercept / subject=panelist; 
/*  lsmeans bm_consistency/ pdiff om bylevel;*/
  ods output tests3=tests_r;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\R Subset.rtf"; 
%rtftable(%str( 
proc rtftable data=tests_r; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "R Subset Summary";
  break spot; 
  id effect; 
  var probf;
  label effect='<w=1.5>Significant Effects' probf="<w=1.5>P-value";
  where probf < 0.05;
run; )); 
ods rtf close;




/*G*/
** not significant effects:   baby_age  diet baby_gender;
** Significant effects:   BM_Consistency BM_Size  Diaper_Size jar_nojar ;

proc sort data = subset; by spot;
proc mixed data = subset order=data;
  by spot;
  class panelist bm_consistency bm_size jar_nojar diaper_size;
  model g = bm_consistency bm_size jar_nojar diaper_size/ ddfm=kr; 
  random intercept / subject=panelist; 
/*  lsmeans bm_consistency/ pdiff om bylevel;*/
  ods output tests3=tests_g;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\G Subset.rtf"; 
%rtftable(%str( 
proc rtftable data=tests_g; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "G Subset Summary";
  break spot; 
  id effect; 
  var probf;
  label effect='<w=1.5>Significant Effects' probf="<w=1.5>P-value";
  where probf < 0.05;
run; )); 
ods rtf close;





/*B*/
** not significant effects:   baby_age  diet Baby_Gender ;
** Significant effects:   BM_Consistency BM_Size Diaper_Size jar_nojar ;

proc sort data = subset; by spot;
proc mixed data = subset order=data;
  by spot;
  class panelist bm_consistency bm_size jar_nojar  diaper_size ;
  model b = bm_consistency bm_size jar_nojar diaper_size / ddfm=kr;
  random intercept / subject=panelist; 
/*  lsmeans bm_consistency/ pdiff om bylevel;*/
  ods output tests3=tests_b;
run;

options nomprint;
%include "\\tsclient\Q\bffc\clinical\Gibb\Lanlan Yao\rtfTable94.sas"; 
ods rtf file="Q:\bffc\clinical\Gibb\Jing Zhang\Color_Diaper2\Results\B Subset.rtf"; 
%rtftable(%str( 
proc rtftable data=tests_b; 
  options order=data orientation=landscape fontSize=11 footnoteFS=9 tocBreak=5; 
  title1 "B Subset Summary";
  break spot; 
  id effect; 
  var probf;
  label effect='<w=1.5>Significant Effects' probf="<w=1.5>P-value";
  where probf < 0.05;
run; )); 
ods rtf close;
