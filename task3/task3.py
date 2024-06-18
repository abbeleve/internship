import pandas as pd
import collections

file_path = "./task3xslx.xlsx"

product = []
period = []
VG = []

with open(file_path, encoding='utf-8')  as r_file:
    df = pd.read_excel(file_path)
    product = df['Продукт'].tolist()
    period = df['Период'].tolist()
    VG = df['ВГ'].tolist()

critical = []
temp = []

for i in range(len(product)):
    vg1 = None
    vg2 = None
    if product[i] not in temp:
        temp.append(product[i])
        vg1 = VG[i]
        for j in range(len(product) - 1, i, -1):
            if product[i] == product[j]:
                vg2 = VG[j]
                break
        if vg2 == None:
            if vg1 < 90:
                critical.append(1)
            else:
                critical.append(0)
        else:
            if period[i][6] < period[j][6]:
                if vg1 - vg2 > 5 and vg2 < 90:
                    critical.append(1)
                else:
                    critical.append(0)
            else:
                if vg2 - vg1 > 5 and vg1 < 90:
                    critical.append(1)
                else:
                    critical.append(0)
index = 0 
result = []
for i in product:
    result.append(critical[temp.index(i)])

df['Критичная позиция'] = result
output_file_path = "./task3xslx_обновленный.xlsx"
df.to_excel(output_file_path, index=False)