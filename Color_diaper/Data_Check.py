import pandas as pd
import numpy as np
import seaborn as sb
import matplotlib.pyplot as plt

diaper  = pd.read_excel('Q:\\bffc\\clinical\\Gibb\\Jing Zhang\\Color_Diaper\\color.xlsx')

# Null
null_counts = diaper.isnull().sum()
null_counts[null_counts > 0].sort_values(ascending=False) # incubation time  38

# Outliers for all
# Z-score
def detect_outlier(color):
    mean = np.mean(color.tolist())
    std = np.std(color.tolist())
    for i in range(142):
        z_score= (color[i] - mean)/std
        if np.abs(z_score) > 3:
            # v = '('+col+')' + str(y)
            outliers.append(i) # +':'+v
    return outliers
# float(np.where(diaper[col] == y)[0])
outlier_dict = {}
outlier_list = []
for var in diaper.columns.values[6:].tolist():
    color_var = diaper[var]
    outliers=[]
    outlier_datapoints= detect_outlier(color_var)
    outlier_list += outlier_datapoints
    if outlier_datapoints:
       outlier_dict.update({var: outlier_datapoints})
    # print(outliers)

print(outlier_list)
len(outlier_list) # 416 data
print(outlier_dict)

unique_outliers = list(set(outlier_list))
unique_outliers.sort()
print(unique_outliers)
len(unique_outliers) # 60

# check if outliers related to BM/Urine
bm_or_not = []
for i in unique_outliers:
    bm_or_not.append(diaper['BM/Urine'][i])
bm_perc = bm_or_not.count('bm')/len(bm_or_not) # 0.717

# check if outliers related to Information
info = []
for i in unique_outliers:
    info.append(diaper['Information'][i])
info_count = pd.Series(info).value_counts()
sb.distplot(info_count)

# check if outliers related to Inside/Outside
inout = []
for i in unique_outliers:
    inout.append(diaper['Inside/Outside'][i])
inout_count = pd.Series(inout).value_counts()
sb.distplot(inout_count)

# check if outliers related to backsheet
back = []
for i in unique_outliers:
    back.append(diaper['w/wo extra backsheet'][i])
back_count = pd.Series(back).value_counts()
sb.distplot(back_count)

# IQR
# for var in diaper.columns.values[6:].tolist():
#     sorted(diaper[var].tolist())
#     q1, q3= np.percentile(diaper[var].tolist(),[25,75])
#     iqr = q3 - q1
#     lower_bound = q1 -(1.5 * iqr) 
#     upper_bound = q3 +(1.5 * iqr)

# Information 
diaper.loc[diaper['BM/Urine']== 'na'].Information.nunique() # 37 each info has a baseline
diaper.Information.nunique() # 37
diaper['Information'].value_counts().sort_index()



################### BM/Urine ###################

freq = diaper['BM/Urine'].value_counts().sort_index() # na: 37
info_distr = sb.distplot(freq)
bm_data = diaper.loc[diaper['BM/Urine']== 'bm'].reset_index(drop=True)
# check distr and outliers
def detect_outlier(color):
    mean = np.mean(color.tolist())
    std = np.std(color.tolist())
    for i in range(75):
        z_score= (color[i] - mean)/std
        if np.abs(z_score) > 3:
            # v = '('+col+')' + str(y)
            outliers.append(i) # +':'+v
    return outliers
outlier_dict = {}
outlier_list = []
for var in bm_data.columns.values[6:].tolist():
    color_var = bm_data[var]
    outliers=[]
    outlier_datapoints= detect_outlier(color_var)
    outlier_list += outlier_datapoints
len(list(set(outlier_list))) # 34 unique outliers
pd.Series(outlier_list).value_counts().sort_index() # 10+: 7,8,14,15,16,35,44,50,51,52 
bm_data.iloc[[7,8,14,15,16,35,44,50,51,52]]

bm_w = bm_data.loc[bm_data['w/wo extra backsheet'] == 'w'].reset_index(drop=True)
bm_wo = bm_data.loc[bm_data['w/wo extra backsheet'] == 'wo'].reset_index(drop=True)



diaper.loc[diaper['BM/Urine']== 'urine']
diaper.loc[diaper['BM/Urine'].isin(['std', 'std2', 'std3'])]





# backsheet









# plt.savefig()