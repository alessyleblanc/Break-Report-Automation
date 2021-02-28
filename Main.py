import pandas as pd
import xlsxwriter

WB = pd.ExcelFile('Break Report')
workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet('Master Sheet')

worksheet.write('A1', 'Client Name')
worksheet.write('B1', '# of portfolios')
worksheet.write('C1', 'total # of breaks')
worksheet.write('D1', 'Breaks 0 -30')
worksheet.write('E1', 'Breaks 31- 60')
worksheet.write('F1', 'Breaks > 60')

df1 = pd.read_excel(WB, 0)
df1_imms = df1[df1['System'] == 'IMMS']
breaks = [len(df1_imms['AgeGroup']), len(df1_imms[df1_imms['AgeGroup'] == '0-30']),
          len(df1_imms[df1_imms['AgeGroup'] == '31-60']), len(df1_imms[df1_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(9, i+2, breaks[i])

df2 = pd.read_excel(WB, 1)
df2_imms = df2[df2['System'] == 'IMMS']
breaks = [len(df2_imms['AgeGroup']), len(df2_imms[df2_imms['AgeGroup'] == '0-30']),
          len(df2_imms[df2_imms['AgeGroup'] == '31-60']), len(df2_imms[df2_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(10, i+2, breaks[i])

df3 = pd.read_excel(WB, 2)
df3_imms = df3[df3['System'] == 'IMMS']
breaks = [len(df3_imms['AgeGroup']), len(df3_imms[df3_imms['AgeGroup'] == '0-30']),
          len(df3_imms[df3_imms['AgeGroup'] == '31-60']), len(df3_imms[df3_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(11, i+2, breaks[i])

df4 = pd.read_excel(WB, 3)
df4_imms = df4[df4['System'] == 'IMMS']
breaks = [len(df4_imms['AgeGroup']), len(df4_imms[df4_imms['AgeGroup'] == '0-30']),
          len(df4_imms[df4_imms['AgeGroup'] == '31-60']), len(df4_imms[df4_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(12, i+2, breaks[i])

df5 = pd.read_excel(WB, 4)
df5_imms = df5[df5['System'] == 'IMMS']
breaks = [len(df5_imms['AgeGroup']), len(df5_imms[df5_imms['AgeGroup'] == '0-30']),
          len(df5_imms[df5_imms['AgeGroup'] == '31-60']), len(df5_imms[df5_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(13, i+2, breaks[i])

df6 = pd.read_excel(WB, 5)
df6_imms = df6[df6['System'] == 'IMMS']
breaks = [len(df6_imms['AgeGroup']), len(df6_imms[df6_imms['AgeGroup'] == '0-30']),
          len(df6_imms[df6_imms['AgeGroup'] == '31-60']), len(df6_imms[df6_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(14, i+2, breaks[i])

df7 = pd.read_excel(WB, 6)
df7_imms = df7[df7['System'] == 'IMMS']
breaks = [len(df7_imms['AgeGroup']), len(df7_imms[df7_imms['AgeGroup'] == '0-30']),
          len(df7_imms[df7_imms['AgeGroup'] == '31-60']), len(df7_imms[df7_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(15, i+2, breaks[i])

df8 = pd.read_excel(WB, 7)
df8_imms = df8[df8['System'] == 'IMMS']
breaks = [len(df8_imms['AgeGroup']), len(df8_imms[df8_imms['AgeGroup'] == '0-30']),
          len(df8_imms[df8_imms['AgeGroup'] == '31-60']), len(df8_imms[df8_imms['AgeGroup'] == '>60'])]

for i in range(len(breaks)):
    worksheet.write(16, i+2, breaks[i])

workbook.close()
