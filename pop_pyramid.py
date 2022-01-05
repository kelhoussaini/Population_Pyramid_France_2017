from xlrd.formula import cellname


#Run cell
#%%
#click on Run Cell to get all plots with an interactive mode
# packages
import numpy as np
import math  # needed to plot histogram with number of bins, as Sturge rule
import xlrd  # to open excel
import time
from matplotlib import pyplot as plt
#from PIL import Image


###   =========================================   ###
# open xls file

file = xlrd.open_workbook(
    'pop-sexe-age-quinquennal6817.xls', on_demand=True)

name = "COM_2017" #sheet name
sheet = file.sheet_by_name(name)

#print(sheet.__dict__.keys())
#print(sheet.ncols, sheet.nrows)  # columns number, rows number

cols_labels = sheet.row(12)

numrow = 10  # 11th row, where we find age incrementation , step = 5 years old
# numrow = 11 --> here, we find sexe labels : 1 for man && 2 for women

###   ===============   DATA PREPARATION   =======================   ###
# we create a dictionary to get people depending on the age and gender
dic = {}
# we separate the gender in the dictionary
dic["Age"] = []
dic["Men"] = []
dic["Women"] = []
age = []

start_time = time.time()

for col in range(6, sheet.ncols): # starting from the 7th column
    m = []; f = [] # initialize the list, m:male & f:female
    age.append(
        float(sheet.cell(10, col).value))  # do not worry,
    # we remove duplicates at the end

    for i in sheet.col_slice(colx=col, start_rowx=14,
                             end_rowx=sheet.nrows - 1):
        if sheet.cell(numrow + 1, col).value == 1.0:  # men
            # missing values, filled with 0
            if not i.value:  # empty string ''
                m.append(0)
            else:
                m.append(i.value)

        if sheet.cell(numrow + 1, col).value == 2.0:  # women
            # missing values, filled with 0
            if not i.value:  # empty string ''
                f.append(0)
            else:
                f.append(i.value)

    # an empty m is created for women set, however,
    # below, we ignore this empty list already created by boocle for
    # because we want to separate F and M in the dictionary

    ###  =========================================   ###
    if m:  # List is not empty
        dic["Men"].append(m)

    if f:  # List is not empty
        dic["Women"].append(f)
    ###   =========================================   ###

# age list includes duplicated elements : same age for man & women
###   ==================== remove duplicates elts =====================   ###
dic["Age"] = list(dict.fromkeys(age))

print("--- %s seconds ---" % (time.time() - start_time))


###   ===============   PLOTS (figure1)  =======================   ###

sum_men = [sum(dic["Men"][i]) for i in np.arange(len(dic["Age"]))]
sum_women = [-sum(dic["Women"][i]) for i in np.arange(len(dic["Age"]))]

# create age labels for the plot
yticklabels = []
for i in dic['Age'][:-1]:
    yticklabels.append(str(int(i)) + "-" + str(int(i) + 4))
yticklabels.append(">=95")
#yticklabels--> ['0-4','5-9',...,'90-94','>=95']

#fig = plt.figure(1)
fig, ax = plt.subplots()
ax.barh(dic["Age"], sum_men, 5.3,
        label="Men", color='b', linewidth=0.5,
        align='center', alpha=0.5)

ax.barh(dic["Age"], sum_women, 5.3,
        label="Women", color='r', linewidth=0.5,
        align='center', alpha=0.5)

ax.set_title("Population pyramid -- 2017 , France")
ax.set_ylabel("Ages")
ax.set_xlabel("Population")
ax.legend(('Men', 'Women'))

ind = 5 * np.arange(len(yticklabels))  # the x locations for the groups
ax.set_yticks(ind)
ax.set_yticklabels(yticklabels, minor=False)  #plt.gca()
plt.show()
#plt.savefig("figure1.png") to save the plot
#img = Image.open("figure1.png")
#img.show()

###   ===============   PLOTS (figure2)  =======================   ###

def sum_list(list1, list2):
    return list1 + list2

y = [ sum_list(dic["Men"][i], dic["Women"][i])
    for i in np.arange(len(dic["Age"])) ]
y_men = [dic["Men"][i] for i in np.arange(len(dic["Age"]))]
y_women = [dic["Women"][i] for i in np.arange(len(dic["Age"]))]

#20 lists in y
res = []
for i in y:
    res = res + i

# separate the gender in the histogram
res_men = []
res_women = []
for i in y_men:
    res_men = res_men + i
for i in y_women:
    res_women = res_women + i

# plot figure2
fig = plt.figure(2)
length = len(res)
num_bins = round(1 + math.log2(length))  # Sturge rule
#print(num_bins) --> 22
plt.hist(res, num_bins, facecolor='blue', alpha=0.5)
plt.title("Absolute frequencies of people")
plt.xlabel("people(Intervals)")
plt.ylabel("Absolute frequencies")
plt.show()


# plot figure3
# we use log, to better visualize the plot

fig = plt.figure(3)
length = len(res)
num_bins = round(1 + math.log2(length))  # Sturge rule
#print(num_bins)
plt.hist(res, num_bins, color='blue',
         edgecolor='red', density=True,
         orientation='horizontal',
         rwidth=0.8, alpha=.5, log=True)
plt.title("Logarithmic absolute frequencies of people")
plt.xlabel("Absolute frequencies (log)")
plt.ylabel("people(Intervals)")
plt.show()

# plot figure4
fig = plt.figure(4)

length = len(res_men)
num_bins = round(1 + math.log2(length))  # Sturge rule
#print(num_bins)
plt.hist([res_men, res_women], num_bins, label=["Men", "Women"], alpha=0.5)
plt.legend()

plt.title("Absolute frequencies of people")
plt.xlabel("people(Intervals)")
plt.ylabel("Absolute frequencies")
plt.show()

# plot figure5
# # we use log, to better visualize the plot
fig = plt.figure(5)

length = len(res)
num_bins = round(1 + math.log2(length))  # Sturge rule
#print(num_bins)
plt.hist([res_men, res_women],
         num_bins,
         edgecolor='red',
         density=True,
         orientation='horizontal',
         rwidth=0.8,
         alpha=.5,
         log=True,
         label=["Men", "Women"])

plt.legend()
plt.title("Logarithmic absolute frequencies of people")
plt.xlabel("Absolute frequencies (log)")
plt.ylabel("people(Intervals)")
plt.show()

# Percentage of people aged 15–24 years, for each municipality
# ============================================================================
res_comune = []  # % of people (all ages), by comune
pop_comune = []  # total pop by comune
slice_comune = []  # % of people aged 15-24 years, by comune

for com in range(len(dic["Men"][0])):
    A = [dic["Men"][k][com] for k in range(len(dic["Age"]))]
    B = [dic["Women"][k][com] for k in range(len(dic["Age"]))]
    total_pop = sum(A + B)
    slice_A = A[3:5]  # men 15-24
    slice_B = B[3:5]  # women 15-24
    sliced = sum(slice_A) + sum(slice_B)  # population age: 15-24
    if total_pop != 0:
        percent = (sliced / total_pop) * 100
    else:
        percent = np.nan  # we can remove the row or replace nan by neg value,
        # or ignore the commune when visualizing

    pop_comune.append(total_pop)  # total pop by comune
    res_comune.append(percent)  # total %pop aged 15-24 years, by comune
    slice_comune.append(sliced)  # total 15-24 by comune

print(" First 5 elements : ", res_comune[:5])

# Mean (avg) and Standard deviation (std)
# ============================================================================
avg_comune = [] # mean
std_comune = []  # std

for com in range(len(dic["Men"][0])):
    A = [dic["Men"][k][com] for k in range(len(dic["Age"]))]
    B = [dic["Women"][k][com] for k in range(len(dic["Age"]))]
    slice_A = A[3:5]
    slice_B = B[3:5]
    avg = np.mean(slice_A + slice_B)
    std = (np.var(slice_A + slice_B))**0.5  # or np.sqrt
    avg_comune.append(avg)
    std_comune.append(std)

print(" First 5 elements (mean): ", avg_comune[:5])
print(" First 5 elements (std): ", std_comune[:5])


# Percentage of people aged 15-24 years in France
avg_France = np.mean(avg_comune)
print(" '%'of people aged 15-24 years in France ", avg_France)

# Largest and lowest extreme values
# ============================================================================

w = np.asarray(res_comune)[
                     np.logical_not(np.isnan(res_comune))]
# removing NaNs, not needed if we used a larger value as 555
# with NaNs, we do not need to look at the results
# to fix the value (555 or more than)

v = np.unique(w)
max_val = np.unique(v)[-1]  #[-1] using NaN,
#[-2] with a value : the last one is 555,
# which corresponds to municipalities with missing values

extrem_bas = np.where(np.array(res_comune) == min(v))[0]
# or np.flatnonzero(np.array(res_comune) == min(v))
extrem_haut = np.where(np.array(res_comune) == max_val)[0]

print(" ==============  Largest extreme value  =============== ")

com = extrem_haut[0] # comune with the largest % of people aged 15-24 years
comune_name = sheet.col_slice(colx=5,
                              start_rowx=14 + com,
                              end_rowx=14 + com + 1)
comune_code_insee_DR = sheet.col_slice(colx=1,
                                       start_rowx=14 + com,
                                       end_rowx=14 + com + 1)
comune_code_insee_CR = sheet.col_slice(colx=2,
                                       start_rowx=14 + com,
                                       end_rowx=14 + com + 1)
comune_code_insee = comune_code_insee_DR[0].value + \
                        comune_code_insee_CR[0].value

print("---> Comune Name    :", comune_name[0].value, "\n")
print("---> Comune Code INSEE   :", comune_code_insee, "\n")
print("---> '%' of people aged 15-24 years by comune   :",
      round(res_comune[com], 2), " %", "\n") # or res_comune[com]
print("--->  Comune Population 'all ages'   :", pop_comune[com], "\n")
print("--->  Comune Population, people aged 15-24 years   :",
      slice_comune[com], "\n")
print(" ==============  Largest extreme value  (END) =============== ")

print(" ==============  Lowest extreme value  =============== ")

for com in list(extrem_bas):# comune with the lowest % of 15-24 years old people
    comune_name = sheet.col_slice(colx=5,
                                  start_rowx=14 + com,
                                  end_rowx=14 + com + 1)
    comune_code_insee_DR = sheet.col_slice(colx=1,
                                           start_rowx=14 + com,
                                           end_rowx=14 + com + 1)
    comune_code_insee_CR = sheet.col_slice(colx=2,
                                           start_rowx=14 + com,
                                           end_rowx=14 + com + 1)
    comune_code_insee = comune_code_insee_DR[0].value + \
                        comune_code_insee_CR[0].value

    print("---> Comune Name    :", comune_name[0].value, "\n")
    print("---> Comune Code INSEE   :", comune_code_insee, "\n")
    print("---> '%' of people aged 15-24 years by comune   :",
          round(res_comune[com], 2), " %", "\n") # or res_comune[com]
    print("--->  Comune Population 'all ages'   :", pop_comune[com], "\n")
    print("--->  Comune Population, people aged 15-24 years   :",
          slice_comune[com], "\n")

# ================(option)=================
# bulding a dictionay including lowest extreme % and the corresponding
# comune name, total population, etc
dict_1 = {}
for com in list(extrem_bas):
    comune_name = sheet.col_slice(colx=5,
                                  start_rowx=14 + com,
                                  end_rowx=14 + com + 1)
    comune_code_insee_DR = sheet.col_slice(colx=1,
                                           start_rowx=14 + com,
                                           end_rowx=14 + com + 1)
    comune_code_insee_CR = sheet.col_slice(colx=2,
                                           start_rowx=14 + com,
                                           end_rowx=14 + com + 1)
    comune_code_insee = comune_code_insee_DR[0].value + \
                        comune_code_insee_CR[0].value

    dict_1[comune_name[0].value] = {}
    dict_1[comune_name[0].value]["comune_code_insee"] = comune_code_insee
    dict_1[comune_name[0].value] \
                    ["'%' of people aged 15-24 years"]=round(res_comune[com], 2)
    dict_1[comune_name[0].value] \
                    ["Comune Population 'all ages'"] = pop_comune[com]
    dict_1[comune_name[0].value] \
                    ["Comune Population, aged 15-24 years"] = slice_comune[com]

#print(dict_1) # printing the dictionary
#comunes_list = list(dict_1.keys()) # list of all comunes with the lowest %
# print(comunes_list)
#comune_index = list(dict_1.keys())[0] # the first one
#dict_1[comune_index] # the required corresponding data
print(" ==============  Lowest extreme value  (END) =============== ")



# %%
