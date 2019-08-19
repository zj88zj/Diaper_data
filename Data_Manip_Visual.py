import pandas as pd
import numpy as np 
import sas7bdat
from sas7bdat import *
import seaborn as sb 
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
import six
import docx
# import pygal
from glob import glob

def format_dict(var_name):
    if var_name == "Trt":
        dict = {'D':'Pampers', 'AW':'First Washout', 'GW':'Second Washout', 'G':'No Preference', 'K':'Huggies',  'A':'Honest/WW', 'X':'Honest/WW'}
        return dict
    elif var_name == "visit":
        dict = {'VISIT 1':'Enrollment', 'VISIT 2':'Day 4', 'VISIT 3':'Day 7', 'VISIT 4':'Day 10', 'VISIT 5':'Day 13'}
        return dict
    elif var_name == "diet":
        dict = {'BM':'Breast Milk Only', 'MBM':'Mostly Breast Milk (' + '\u2265' + '50%)', 'F':'Formula Only or Most (' + '\u2265'+ '50%)'}
        return dict
    elif var_name == 'samp':
        dict = {'PREBM':'Pre-BM', 'POSTBM':'Post-BM', 'BM': 'BM'}
        return dict
    else: print("format wrong!")
#  then use map for formatting

#######################################################
########                                    ###########
########         Data Analysis              ###########
########                                    ###########
#######################################################


###################  data imput  ######################

def data_input(dataname):
    data_path = "Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Data\\Final"
    dataname = SAS7BDAT(data_path + '\\' + dataname + '.sas7bdat').to_data_frame()
    return dataname

aesum = data_input("aesum")
anth = data_input("anth")
ba = data_input("ba")
bm = data_input("bm")
codelists = data_input("codelists")
comments = data_input("comments")
comply = data_input("comply")
criter = data_input("criter")
decodetrtbyidentity = data_input("decodetrtbyidentity")
demog = data_input("demog")
derm = data_input("derm")
diaper = data_input("diaper")
eventdt = data_input("eventdt")
fit = data_input("fit")
meds = data_input("meds")
parent_end = data_input("parent_end")
parent_qnaire = data_input("parent_qnaire")
parent_visit = data_input("parent_visit")
ph = data_input("ph")
samples = data_input("samples")
studynotes = data_input("studynotes")
subjstat = data_input("subjstat")
trt = data_input("trt")
trtdsc = data_input("trtdsc")

# filenames = glob('*.sas7bdat')
# dfs = [pd.read_csv(data_path + '\\' + f) for f in filenames] # get a list of dataframes

# extra data from Analytical & Micro file
micro = SAS7BDAT('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Data\\Analytical & Micro\\micro.sas7bdat').to_data_frame()
analytical_bm = SAS7BDAT('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Data\\Analytical & Micro\\analytical_bm.sas7bdat').to_data_frame()
analytical_swabs = SAS7BDAT('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Data\\Analytical & Micro\\analytical_swabs.sas7bdat').to_data_frame()


###################  Derived Data  ######################

# trt
trt_keep = trt.loc[:,['identity', 'visit', 'Trt']]
trt_keep.visit = trt_keep.visit.replace('VISIT 2', 'VISIT 3').replace('VISIT 4','VISIT 5')
trt_keep.to_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Derived_Data\\ph.xlsx')

# pH
ph_keep = ph.loc[:, ['identity', 'visit', 'page', 'displaytext', 'Value', 'Unit', 'BodyLoc', 'Side', 'Site', 'evalfl', 'Value_Read_datetime']]
ph_keep.to_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Derived_Data\\trt.xlsx')


# BM 
bm_keep = bm.loc[:, ['identity', 'visit', 'bm_pH', 'evalfl']].rename(columns = {'bm_pH': 'value'})
bm_keep['page'] = "BM PH"
bm_keep['displaytext'] = "BM PH"
bm_keep2 =  bm.loc[:, ['identity', 'visit', 'bm_type', 'bm_amount', 'bm_sample', 'bm_vials', 'bm_pH', 'evalfl']]
bm_keep2.to_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Derived_Data\\BMData.xlsx')

# diaper 
diaper_keep = diaper.loc[:, ['identity', 'visit', 'procedure', 'evalfl']]
prepost = diaper_keep.loc[diaper_keep['procedure'] != '']
prepost.to_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Derived_Data\\prepost.xlsx')


###################  Data Check  ######################

#### Compliance Criteria
comply.loc[(comply['CompCd'] == 'QUES01') & (comply['resfl'] == 'N')]  # No babies health changed
comply.loc[(comply['CompCd'] == 'QUES02') & (comply['resfl'] == 'Y')]  # Meds changed: 3001 visit 3, 3007 visit 2, 3010 visit 4, 3025 visit 2, 3033 visit 2 and 3, 3035 visit 2, 3037 visit 5, 3066 visit 3
comply.loc[(comply['CompCd'] == 'QUES03') & (comply['resfl'] == 'N')]  # 3025 visit2 not returned products
#   if CompCd='QUES04' and resfl='N'; ** No one used other diaper
#   if CompCd='QUES05' and resfl='N'; ** No one used other wipe
#   if CompCd='QUES06' and resfl='N'; ** 3003 visit 4 didn't give baby a bath at prior evening, but the morning of the visit
#   if CompCd='QUES07' and resfl='N'; ** Fresh diaper was changed before 30 -60 mins before the visit
#   if CompCd='QUES08' and resfl='N'; ** Use lotion in diaper area: 3005 visit 3, 4, 5
#   if CompCd='QUES09' and resfl='Y'; ** 3003 visit 4 gave bath this morning
#   if CompCd='QUES10' and resfl='Y'; ** No one in another study

#### Inclusion/Exclusion Criteria
criter.loc[(criter['critypcd'] == 'EXCL') & (criter['criansfl'] == 'Y')] 
# 3022 EXCL04 taking oral or topical medication for medical or skin condition, 3029 EXCL07 have underlying skin condition or disease exclude from study, 3072 EXCL06 Baby taking oral or topical medication impactful to skin(steroid, etc.)
#   if critypcd='INCL' and criansfl='N'; *** 3002 INCL08 use lotion near diaper area
#                           3004 INCL08 use lotion near diaper area    3011 INCL04 not giving bath last evening   3021 INCL03 not wear size 0 or 1 diaper    3028 INCL03 not wear size 0 or 1 diaper    3053 INCL03 not wear size 0 or 1 diaper
#                           3061 INCL03 not wear size 0 or 1 diaper    3064 INCL03 not wear size 0 or 1 diaper    3072 INCL03 not wear size 0 or 1 diaper    3076 INCL04 not giving bath last evening;
  
#### Event Date
eventdt.edate.isnull().sum() #check null
eventdt_keep = eventdt.loc[:, ['identity', 'visit', 'edate']]
eventdt_transpose = eventdt_keep.pivot(index = 'identity', columns = 'visit', values = 'edate').fillna(pd.NaT)
eventdt_transpose['d1'] = eventdt_transpose['VISIT 2'] - eventdt_transpose['VISIT 1']
eventdt_transpose['d2'] = eventdt_transpose['VISIT 3'] - eventdt_transpose['VISIT 2']
eventdt_transpose['d3'] = eventdt_transpose['VISIT 4'] - eventdt_transpose['VISIT 3']
eventdt_transpose['d4'] = eventdt_transpose['VISIT 5'] - eventdt_transpose['VISIT 4']
eventdt_transpose[['d1','d2','d3','d4']].nunique() # check if day period are the same

### pH and trt data
ph_check = ph_keep.loc[(ph_keep['visit'] == 'VISIT 2') & ( ~ ph_keep['page'].isin(['CHEST PH', 'PRE-BM PH'])) & (ph_keep['BodyLoc'] == 'GENITAL')]
ph_transpose = ph_check.pivot(index = 'identity', columns = 'page', values = 'Value_Read_datetime').fillna(pd.NaT)
ph_transpose['d1'] = ph_transpose['POST-15 PH'] - ph_transpose['POST-BM PH']
ph_transpose['d2'] = ph_transpose['POST-30 PH'] - ph_transpose['POST-15 PH']
ph_transpose['d3'] = ph_transpose['POST-60 PH'] - ph_transpose['POST-30 PH']

ph_drop = ph_keep.dropna(subset = ['Value'])
ph_merge = pd.merge(trt_keep, ph_drop, on= ['identity', 'visit'], how = 'right')
ph_merge.loc[ph_merge.visit.isin(['VISIT 4', 'VISIT 2']), 'Trt'] = 'W'
ph_merge.loc[ph_merge.visit == 'VISIT 1', 'Trt'] = 'A'

ph_freq_table = pd.crosstab(ph_merge.page, ph_merge.BodyLoc)#.apply(lambda r : r/r.sum(), axis = 1) #for percentage
ph_value_summary = ph_merge.Value.describe()
#num = ph_merge.groupby(['identity', 'visit'])['Value'].count()
num = pd.pivot_table(data = ph_merge, values = 'Value', index = 'identity', columns = 'visit', aggfunc = lambda x : x.count())
num.to_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Num_of_Measure_py.xlsx')


###################  Demographics  ######################

####### Subject Disposition #######
status = subjstat.loc[:, ['identity', 'statypcd', 'wdre']]
trt_sub = trt_keep.identity.unique()
trt_ls = trt_sub.tolist()
# unique, counts = np.unique(trt_sub, return_counts=True)
# counts.sum()
status.loc[status.identity.isin(trt_ls),'trt'] = 1
# pd.merge(pd.DataFrame(trt_sub, columns = ['identity']), status, on = 'identity', how = 'outer')
# status_ls = status.identity.tolist()
# len(status_ls)
# len(trt_ls)
# len([item for item in status_ls if item not in trt_ls])
sta1 = pd.DataFrame(data = {'N' : status.identity.nunique(), 'statypcd': 'Recruited'}, index = [0])
sta2 = pd.DataFrame(data = {'N' : status.loc[status.statypcd == 'SCRFL'].identity.nunique(), 'statypcd': 'Screen Failed'}, index =[0])
sta3 = pd.DataFrame(data = {'N' : status.loc[status.trt == 1].identity.nunique(), 'statypcd': 'Randomized'}, index =[0])
sta4 = pd.DataFrame(data = {'N' : status.loc[status.statypcd == 'DROPOUT'].identity.nunique(), 'statypcd': 'Dropped during Treatment'}, index =[0])
sta5 = pd.DataFrame(data = {'N' : status.loc[status.statypcd == 'STDYCOMP'].identity.nunique(), 'statypcd': 'Completed'}, index =[0])
sta = pd.concat([sta1,sta2,sta3,sta4,sta5])
### rtf table
sta.columns = ['Number of Subject','Subject Status']
def table(data, col_width=3.0, row_height=0.625, font_size=14, header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w', bbox=[0, 0, 1, 1], header_columns=0, ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in  six.iteritems(mpl_table._cells):
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    return ax
table(sta, header_columns=0, row_height = 0.5, col_width=3, font_size = 10)
plt.savefig('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Results\\Demog\\Subject_Disposition.png')
plt.close()

####### Demog #######

# doc = docx.Document()


##### Mixed Model #####

######## t Test #######







#######################################################
########                                    ###########
########         Data Visualizations        ###########
########                                    ###########
#######################################################


################## Histogram Graphs ###################

##### vertical #####

##### 1 pH Analysis -- seaborn
table_ph = pd.read_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Py\\table_ph.xlsx')
table_ph.Trt = table_ph.Trt.map(format_dict('Trt'))
table_ph = table_ph.loc[(table_ph.BodyLoc != 'Chest') & (table_ph.page != 'CHEST PH')].drop(columns = ['Obs', 'point', 'endpoint']) #.rename(columns = {'Estimate':'Mean pH', 'StdErr':'Standard Error', 'N':'Number of Subjects', 'diff':'Mean Difference', 'Probt':'P-Value'})

table_g = table_ph.loc[table_ph.BodyLoc == 'Genital']
table_g_drop = table_g.drop(columns = ['BodyLoc', 'upper', 'lower', 'label'])
table_g_drop = table_g_drop.astype({'diff':str, 'Probt':str})
table_pivot = pd.pivot_table(data = table_g_drop, index = ['page', 'Trt'], values = ['Estimate', 'StdErr', 'N', 'diff', 'Probt'], 
                            aggfunc = {'Probt' : lambda x : ''.join(x), 'diff': lambda x : ''.join(x), 'Estimate': lambda x: x , 'StdErr': lambda x: x, 'N': lambda x: x})
values_order = ['Estimate', 'StdErr', 'N', 'diff', 'Probt']
page_order = ['PRE-BM PH', 'POST-BM PH', 'POST-15 PH', 'POST-30 PH', 'POST-60 PH']
table_show = table_pivot.reindex(values_order, axis =1)
table_show = table_show.reindex(page_order, level = 0)
table_show = pd.DataFrame(table_show).T

sb.set(style="darkgrid")
colors = ["light grey blue", "grey pink"]
myPalette = sb.xkcd_palette(colors)
g = sb.barplot(x = 'page', y = 'Estimate', hue = 'Trt', data = table_g, palette = myPalette, saturation = 0.8)
# g = sb.catplot(x = 'page', y = 'Estimate', hue = 'Trt', data = table_ph, col = 'BodyLoc', kind = 'bar', height = 3, aspect = 2, palette = myPalette, saturation = 0.6) # catplot can combine two locations but without errorbars
g.set(xlabel = '', ylabel = 'Mean pH'+'\u00B1'+'SE')
plt.suptitle('pH Summary at Genital', fontsize = 14, y = 0.95)
plt.ylim(4.5, 6)
for p in g.patches:
    height = p.get_height()
    g.annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.06), ha = 'center')
# ax = g.axes
# for p in ax[0][0].patches:
#     height = p.get_height()
#     ax[0][0].annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.02), ha = 'center')
# for p in ax[0][1].patches:
#     height = p.get_height()
#     ax[0][1].annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.02), ha = 'center')

cell_text = []
for row in range(1, table_show.shape[0]+1):
    for x in table_show[row-1:row].values.astype(str).tolist():
        formatedx = [s[:6] for s in x]
        cell_text.append(formatedx)
collables = ('Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers')
the_table = plt.table(cellText = cell_text, rowLabels = values_order, colLabels = collables , loc = 'bottom',  cellLoc='center', bbox = [0, -0.9, 1, 0.800 ], colColours= ['#9dbcd4', '#c3909b']*5, rowColours= ['lightgrey']*5)
the_table.auto_set_font_size(False)
the_table.set_fontsize(10)
# the_table.scale(1.5,1.5)
plt.subplots_adjust(left =0.2, bottom = 0.5)
plt.tight_layout()
#plt.title('pH Summary')
plt.show()
plt.savefig('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Py\\1_ph_analysis.png', bbox_inches = 'tight', dpi = 300)
plt.close()


##### 6 - Swab and BM Enzyme by Diet -- matplotlib
table_enzyme = pd.read_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Py\\table_enzyme.xlsx')
table_enzyme.samp = table_enzyme.samp.map(format_dict('samp'))
table_enzyme.Trt = table_enzyme.Trt.map(format_dict('Trt'))
table_enzyme['diet'] = table_enzyme['VISIT_3']
table_enzyme = table_enzyme.drop(columns = ['VISIT_3','label', 'order'])

fig  = plt.figure()
gs = gridspec.GridSpec(2,3)
N = table_enzyme.samp.nunique()
width = 0.55
ax1 = fig.add_subplot(gs[0,0])
ax1.bar(np.arange(N), table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'BM')]['Estimate'], width, color = 'lightcoral', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'BM')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightsteelblue'))
ax1.bar(np.arange(N)+width, table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'BM')]['Estimate'], width, color = 'lightsteelblue', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'BM')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightcoral'))
# ax1.scatter() # add lable- significant mark over bars
ax1.set_title('          Breast Milk Only          ', bbox=dict(boxstyle="square", ec=(0.1, 0.2, 0.1), fc=(0.5, 0.6, 0.6), alpha = 0.5), fontweight='bold', fontsize = 7)
ax1.set_xticks(np.arange(N)+ width/2)
ax1.tick_params(axis='y', labelsize=5)
ax1.set_xticklabels(['Pre-BM', 'BM', 'Post-BM'], fontsize = 7)
ax1.set_ylabel('Mean(Log10 ng Trypsin Equivalent)'+'\u00B1'+'SE', fontsize = 7)
ax1.set_ylim(1,4)
ax1.set_facecolor('whitesmoke')
for p in ax1.patches:
    height = p.get_height()
    ax1.annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.15), ha = 'center', fontsize = 5)

ax2 = fig.add_subplot(gs[0,1])
ax2.bar(np.arange(N), table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'MBM')]['Estimate'], width, color = 'lightcoral', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'MBM')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightsteelblue'))
ax2.bar(np.arange(N)+width, table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'MBM')]['Estimate'], width, color = 'lightsteelblue', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'MBM')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightcoral'))
ax2.set_title('  Mostly Breast Milk (' + '\u2265' + '50%)  ', bbox=dict(boxstyle="square", ec=(0.1, 0.2, 0.1), fc=(0.5, 0.6, 0.6), alpha = 0.5), fontweight='bold', fontsize = 7)
ax2.set_xticks(np.arange(N)+ width/2)
ax2.set_xticklabels(['Pre-BM', 'BM', 'Post-BM'], fontsize = 7)
ax2.set_yticks(np.arange(1), (''))
ax2.set_ylim(1,4)
ax2.set_facecolor('whitesmoke')
for p in ax2.patches:
    height = p.get_height()
    ax2.annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.15), ha = 'center', fontsize = 5)

ax3 = fig.add_subplot(gs[0,2])
ax3.bar(np.arange(N), table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'F')]['Estimate'], width, color = 'lightcoral', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Pampers')&(table_enzyme.diet == 'F')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightsteelblue'))
ax3.bar(np.arange(N)+width, table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'F')]['Estimate'], width, color = 'lightsteelblue', yerr = table_enzyme.loc[(table_enzyme.Trt == 'Huggies')&(table_enzyme.diet == 'F')]['StdErr'], error_kw = dict(elinewidth = 2, ecolor = 'lightcoral'))
ax3.set_title('Formula Only or Most (' + '\u2265'+ '50%)', bbox=dict(boxstyle="square", ec=(0.1, 0.2, 0.1), fc=(0.5, 0.6, 0.6), alpha = 0.5), fontweight='bold', fontsize = 7)
ax3.set_xticks(np.arange(N)+ width/2)
ax3.set_xticklabels(['Pre-BM', 'BM', 'Post-BM'], fontsize = 7)
ax3.set_yticks(np.arange(1), (''))
ax3.set_ylim(1,4)
ax3.set_facecolor('whitesmoke')
for p in ax3.patches:
    height = p.get_height()
    ax3.annotate("%.2f" % height, (p.get_x()+ p.get_width()/2., height+0.15), ha = 'center', fontsize = 5)

plt.subplots_adjust(wspace = .005)
red_patch = mpatches.Patch(color='lightcoral', label='Pampers')
blue_patch = mpatches.Patch(color='lightsteelblue', label='Huggies')
plt.legend(handles=[red_patch, blue_patch], loc = 'best', fontsize = 5)
fig.suptitle('Log10 Total Protease Activity (ng Trypsin Equivalent)', fontsize=8, fontweight='bold')
fig.text(0.1, 0.05, 'Clinical Study 2018047', color = 'lightgrey', fontsize = 4)

table_enzyme_pivot= pd.pivot_table(data = table_enzyme, index = ['diet', 'samp', 'Trt'], values = ['Estimate', 'StdErr', 'NumSamp', 'diff', 'Probt'])
values_order = ['Estimate', 'StdErr', 'NumSamp', 'diff', 'Probt']
diet_order = ['BM', 'MBM', 'F']
samp_order = ['Pre-BM', 'BM', 'Post-BM']
table_enzyme_show = table_enzyme_pivot.reindex(values_order, axis =1)
table_enzyme_show = table_enzyme_show.reindex(samp_order, level = 1)
table_enzyme_show = table_enzyme_show.reindex(diet_order, level = 0)
table_enzyme_show = pd.DataFrame(table_enzyme_show).T

ax4 = fig.add_subplot(gs[1,:])
ax4.axis('off')
cell_text = []
for row in range(1, table_enzyme_show.shape[0]+1):
    for x in table_enzyme_show[row-1:row].values.tolist():
        formatedx = [round(s, 2) for s in x]
        cell_text.append(formatedx)
collables = ('Pampers','Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers', 'Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers','Huggies', 'Pampers', 'Huggies')
the_table = plt.table(cellText = cell_text, rowLabels = ['Mean', 'StdErr', 'Num', 'diff', 'P-value'], colLabels = collables , loc = 'upper center',  colColours= ['#F08080', '#B0C4DE']*9, rowColours= ['lightgrey']*5)
the_table.auto_set_font_size(False)
the_table.set_fontsize(5)
plt.show()
plt.savefig('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Py\\6_Swab_BMEnzyme_Diet.png',  bbox_inches = 'tight', dpi = 360)
plt.close()


##### 1 pH Distribution at Genitals
table_dist = pd.read_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\2018047REP\\Py\\table_ph_dist.xlsx')


##### horizonal #####

##### 

############## Scat/Series- Grsphs ###############


####################### Power Plots #########################


