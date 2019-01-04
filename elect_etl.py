import pandas as pd
import re
from docx.api import Document
from unidecode import unidecode
import numpy as np


def ConvertOnlyNumeric(string):
    newString = ''
    for s in string:
        if s.isdigit():
            s = unidecode(s)
        newString += s
    return newString


document = Document('electzone.docx')

data = []
for j in range(len(document.tables)):
    table = document.tables[j]
    keys = None
    for i, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)

df = pd.DataFrame(data)
df = df.applymap(ConvertOnlyNumeric)
df=df.apply(lambda x: x.str.strip())
df=df.rename(columns=lambda x: re.sub('\n| ','',x))

for i in df['ลำดับที่'].index:
    if df['ลำดับที่'].iloc[i]:
        val = df['ลำดับที่'].iloc[i]
    else:
        df['ลำดับที่'].iloc[i] = val

checkdf = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].str.contains('\n|[ก-๙]\s', regex=True)]
n=0
for i in checkdf.index:
    i+=n
    df = pd.concat(
        [df.loc[:i],
         pd.DataFrame([[''] * len(df.columns)], columns=df.columns, index=[i + 1]),
         df.loc[i + 1:]]) \
        .reset_index(drop=True)
    splitlist = re.findall('[ก-๙()]+', checkdf.loc[i-n])
    df.loc[i:i + len(splitlist) - 1, 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] = splitlist
    df.loc[i + 1, ['ลำดับที่', 'เขตเลือกตั้งที่']] = df.loc[i, ['ลำดับที่', 'เขตเลือกตั้งที่']]
    n+=1

for i in range(1, 78):
    i=str(i)
    containcheck = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i].str.contains('\n')
    vals = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i][containcheck].unique()
    for val in vals:
        dffind = df['เขตเลือกตั้งที่'][(df['เขตเลือกตั้งที่'] == val) & (df['ลำดับที่'] == i)]
        df['เขตเลือกตั้งที่'][(df['เขตเลือกตั้งที่'] == val) & (df['ลำดับที่'] == i)] = pd.Series(
            list(val) + [val[-1]] * (len(dffind) - len(val)), index=dffind.index)

def dup_blank_row(dfcol):
    for i in df.index:
        if dfcol.iloc[i] in ['\n', '', '.']:
            dfcol.iloc[i] = dfcol.iloc[i - 1]

dup_blank_row(df['เขตเลือกตั้งที่'])
dup_blank_row(df['จังหวัด'])

havecondition = False
for i in df.index:
    zonename = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i]
    if '(' in zonename:
        havecondition = True
        zonenum = df['เขตเลือกตั้งที่'].iloc[i]
    if havecondition:
        df['เขตเลือกตั้งที่'].iloc[i] = zonenum
    if havecondition or re.match('\(.*|^[2-9].\s.*|.*\)', df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i]):
        df['เขตเลือกตั้งที่'].iloc[i] = df['เขตเลือกตั้งที่'].iloc[i - 1]
    if ')' in zonename:
        havecondition = False

def get_amphor_index(string):
    return df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].str.contains(string)].index[0]

df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอเวียงสา'] = '2'
df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอปัว'] = '3'
df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอเชียงคำ'] = '2'
df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอดอกคำใต้'] = '3'
df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอภูกามยาว'] = '3'
df['เขตเลือกตั้งที่'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอกาบัง'] = '2'
# df['เขตเลือกตั้งที่'].loc[989:997] = '2'
startindex=get_amphor_index('อำเภอโพธาราม')
df['เขตเลือกตั้งที่'].loc[startindex:startindex+4] = '3'
df['เขตเลือกตั้งที่'].loc[startindex+5] = '4'
startindex=get_amphor_index('อำเภอโพธิ์ชัย')
df['เขตเลือกตั้งที่'].loc[startindex:startindex+12] = '2'
# df['เขตเลือกตั้งที่'].loc[1321:1323] = '6'
startindex=get_amphor_index('อำเภอเขวาสินรินทร์')
df['เขตเลือกตั้งที่'].loc[startindex:startindex+13] = '2'
df['เขตเลือกตั้งที่'].loc[startindex+32:startindex+44] = '4'
# df['เขตเลือกตั้งที่'].loc[1498:1501] = '3'
# df['เขตเลือกตั้งที่'].loc[1502:1503] = '4'

df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].replace('\d{1,2}.\s|^และ', '', regex=True, inplace=True)
df.replace('เทศบาลตำบล', 'ตำบล', regex=True, inplace=True)
df.replace('อำเภอ|^ตำบล|เขต|แขวง', '', regex=True, inplace=True)### for test
# for i in range(1, 78):
#     i = str(i)
#     containcheck = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i]
#     vals = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i].unique()
#     print(f'{i}\n======\n' + df['จังหวัด'].loc[df['ลำดับที่'] == i].iloc[0])
#     for val in vals:
#         print(df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][(df['เขตเลือกตั้งที่'] == val) & (df['ลำดับที่'] == i)].iloc[0])


def clean_brace(dframe, delstart):
    havecondition = False
    for i in df.index:
        zonename = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i]
        if '('+delstart in zonename:
            havecondition = True
        if havecondition:
            dframe.iloc[i] = zonename
            df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i] = ''
        if ')' in zonename:
            havecondition = False
    return dframe.str.replace(f'(\({delstart}|^และ|\)$)', '', regex=True)


df['only'] = ''
df['only']=clean_brace(df['only'], 'เฉพาะ')

df['exclude'] = ''
df['exclude']=clean_brace(df['exclude'], 'ยกเว้น')


def clean_tesban(df):
    for i in df[~df.str.startswith('ตำบล') & df.str.contains("^\w", regex=True)].index:
        if not any(df.iloc[i+1].startswith(x) for x in ['เทศบาล', 'แขวง']) and \
                len(df.iloc[i+1])>0:
            df.iloc[i]=df.iloc[i]+df.iloc[i+1]
            df.iloc[i+1]=''
    isstart = False
    for i in df.index:
        if df.iloc[i]:
            isstart = True
            startloc=i-1
        else:
            startloc=0
        conditionlist = []
        k = 0
        while isstart:
            if df.iloc[i + k]:
                conditionlist.append(df.iloc[i + k])
            else:
                isstart = False
            df.iloc[i + k] = ''
            k += 1
        if startloc!=0:
            df.iloc[startloc] = conditionlist

clean_tesban(df.only)
clean_tesban(df.exclude)

df = df[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].map(len) > 0].reset_index(drop=True)
df.replace('[()]', '', regex=True, inplace=True)
df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']=df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].str.strip()
df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']=='กันทรลักษณ์'] = 'กันทรลักษ์'
# df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']=='ป้อมปราบศัตรูพ่า'] = 'ป้อมปราบศัตรูพ่า'

df.to_csv('electzone-processed.csv')


from shptocsv import shptodf

shpdf=shptodf('/Users/phoneee/Downloads/tambon 1/tambon.shp')


shpdf['ap_tn']=shpdf['ap_tn'].str.strip()
shpdf['ap_tn'][shpdf['ap_tn']=='ป้อมปราบศัตรูพ่า'] = 'ป้อมปราบศัตรูพ่า'
shpdf['ap_tn'][shpdf['ap_tn']=='เมืองนครศรีธรรมร'] = 'เมืองนครศรีธรรมราช'
shpdf['ap_tn'][shpdf['ap_tn']=='เมืองสุราษฎร์ธาน'] = 'เมืองสุราษฎร์ธานี'



# shpdf['AMP_NAME'][shpdf['AMP_NAME']=='เมือง']=shpdf['AMP_NAME'][shpdf['AMP_NAME']=='เมือง']+shpdf['PRV_NAME'][shpdf['AMP_NAME']=='เมือง']
merge=pd.merge(df, shpdf, how='left', left_on=['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'], right_on=['pv_tn', 'ap_tn'])
merrgefail=merge[['จังหวัด','ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']][merge['tb_tn'].isna()]




shpdf=shptodf('/Users/phoneee/Downloads/tambon/tambon.shp')
shpdf['A_NAME_T']=shpdf['A_NAME_T'].str.strip()
merge=pd.merge(df, shpdf, how='left', left_on=['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'], right_on=['P_NAME_T', 'A_NAME_T'])
merrgefail=merge[['จังหวัด','ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']][merge['T_NAME_T'].isna()]
