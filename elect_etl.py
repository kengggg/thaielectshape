import pandas as pd
import re
from docx.api import Document
from unidecode import unidecode
import numpy as np


def ConvertNumeric(string):
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
df = df.applymap(ConvertNumeric)
df = df.apply(lambda x: x.str.strip())
df = df.rename(columns=lambda x: re.sub('\n| ', '', x))

for i in df['ลำดับที่'].index:
    if df['ลำดับที่'].iloc[i]:
        val = df['ลำดับที่'].iloc[i]
    else:
        df['ลำดับที่'].iloc[i] = val

checkdf = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][
    df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].str.contains('\n|[ก-๙]\s', regex=True)]
n = 0
for i in checkdf.index:
    i += n
    df = pd.concat(
        [df.loc[:i],
         pd.DataFrame([[''] * len(df.columns)], columns=df.columns, index=[i + 1]),
         df.loc[i + 1:]]) \
        .reset_index(drop=True)
    splitlist = re.findall('[ก-๙()]+', checkdf.loc[i - n])
    df.loc[i:i + len(splitlist) - 1, 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] = splitlist
    df.loc[i + 1, ['ลำดับที่', 'เขตเลือกตั้งที่']] = df.loc[i, ['ลำดับที่', 'เขตเลือกตั้งที่']]
    n += 1

for i in range(1, 78):
    i = str(i)
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


df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอเวียงสา', 'เขตเลือกตั้งที่'] = '2'
df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอปัว', 'เขตเลือกตั้งที่'] = '3'
df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอเชียงคำ', 'เขตเลือกตั้งที่'] = '2'
df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอดอกคำใต้', 'เขตเลือกตั้งที่'] = '3'
df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอภูกามยาว', 'เขตเลือกตั้งที่'] = '3'
df.loc[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'อำเภอกาบัง', 'เขตเลือกตั้งที่'] = '2'
# df['เขตเลือกตั้งที่'].loc[989:997] = '2'
startindex = get_amphor_index('อำเภอโพธาราม')
df['เขตเลือกตั้งที่'].loc[startindex:startindex + 4] = '3'
df['เขตเลือกตั้งที่'].loc[startindex + 5] = '4'

startindex = get_amphor_index('อำเภอโพธิ์ชัย')
df['เขตเลือกตั้งที่'].loc[startindex:startindex + 12] = '2'
# df['เขตเลือกตั้งที่'].loc[1321:1323] = '6'
startindex = get_amphor_index('อำเภอเขวาสินรินทร์')
df['เขตเลือกตั้งที่'].loc[startindex:startindex + 13] = '2'
df['เขตเลือกตั้งที่'].loc[startindex + 32:startindex + 44] = '4'
# df['เขตเลือกตั้งที่'].loc[1498:1501] = '3'
# df['เขตเลือกตั้งที่'].loc[1502:1503] = '4'


df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].replace('\d{1,2}.\s|^และ', '', regex=True, inplace=True)
df.replace('^ตำบล', '', regex=True, inplace=True)
df.replace('อำเภอ|^ตำบล|เขต|แขวง', '', regex=True, inplace=True)  ### for test


# for i in range(1, 78):
#     i = str(i)
#     containcheck = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i]
#     vals = df['เขตเลือกตั้งที่'].loc[df['ลำดับที่'] == i].unique()
#     print(f'{i}\n======\n' + df['จังหวัด'].loc[df['ลำดับที่'] == i].iloc[0])
#     for val in vals:
#         print(df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][(df['เขตเลือกตั้งที่'] == val) & (df['ลำดับที่'] == i)].iloc[0])


def clean_brace(dframe, delstart):
    havecondition = False
    findbrace = 0
    for i in df.index:
        zonename = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i]
        if '(' in zonename:
            findbrace += 1
        if any(('(' + x in zonename) for x in delstart):
            havecondition = True
        if havecondition:
            dframe.iloc[i] = zonename
            df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].iloc[i] = ''
        if ')' in zonename:
            findbrace -= 1
        if ')' in zonename and findbrace == 0:
            havecondition = False
    return dframe.str.replace(f'(\({"|".join(delstart)}|^และ|\)$)', '', regex=True)


df['interior'] = ''
df['interior'] = clean_brace(df['interior'], ['เฉพาะตำบล', 'เฉพาะ'])
# df['interior']=clean_brace(df['interior'], 'ใน')

df['exterior'] = ''
df['exterior'] = clean_brace(df['exterior'], ['ยกเว้นตำบล', 'ยกเว้น'])


#
def clean_tesban(df):
    # for i in df[~df.str.startswith('ตำบล') & df.str.contains("^\w", regex=True)].index:
    #     if not any(df.iloc[i+1].startswith(x) for x in ['เทศบาล', 'แขวง']) and \
    #             len(df.iloc[i+1])>0:
    #         df.iloc[i]=df.iloc[i]+df.iloc[i+1]
    #         df.iloc[i+1]=''
    for i in df[df.str.endswith('เทศบาลเมือง') | df.str.endswith('เทศบาลนคร')].index:
        if not any(df.iloc[i + 1].startswith(x) for x in ['เทศบาล', 'ตำบล']) and \
                len(df.iloc[i + 1]) > 0:
            df.iloc[i] = df.iloc[i] + df.iloc[i + 1]
            df.iloc[i + 1] = ''
        df.iloc[i] = df.iloc[i].replace('(', '')
    isstart = False
    for i in df.index:
        if df.iloc[i]:
            isstart = True
            startloc = i - 1
        else:
            startloc = 0
        conditionlist = []
        k = 0
        while isstart:
            if df.iloc[i + k]:
                conditionlist.append(df.iloc[i + k])
            else:
                isstart = False
            df.iloc[i + k] = ''
            k += 1
        if startloc != 0:
            df.iloc[startloc] = conditionlist


df.replace('[()]', '', regex=True, inplace=True)
clean_tesban(df.interior)
clean_tesban(df.exterior)

df = df[df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].map(len) > 0].reset_index(drop=True)

df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] = df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'].str.strip()
df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'กันทรลักษณ์'] = 'กันทรลักษ์'
df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'] == 'สนามชัย'] = 'สนามชัยเขต'

startindex = df[(df['ลำดับที่'] == '58') & df.ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง.str.contains('เมืองนครสวรรค์')].index
df['interior'].loc[startindex[0]] = ["นครสวรรค์ตก(ในเทศบาลนครนครสวรรค์)",
                                     "นครสวรรค์ออก(ในเทศบาลนครนครสวรรค์)",
                                     "แควใหญ่(ในเทศบาลนครนครสวรรค์)",
                                     "ปากน้ำโพ",
                                     "วัดไทร",
                                     "บางม่วง",
                                     "บ้านมะเกลือ",
                                     "บ้านแก่ง",
                                     "หนองกรด",
                                     "หนองกระโดน",
                                     "บึงเสนาท"]
df['exterior'].loc[startindex[1]] = df['interior'].loc[startindex[0]]
# df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'][df['ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']=='ป้อมปราบศัตรูพ่า'] = 'ป้อมปราบศัตรูพ่า'

df.to_csv('electzone-processed.csv')

from shptocsv import shptodf

# shpdf = shptodf('tambon_old/tambon_old.shp')
shpdf = shptodf('tambon/TH_Tambon.shp')

shpdf['A_NAME_T'] = shpdf['A_NAME_T'].str.strip()
shpdf.loc[shpdf['A_NAME_T'] == 'ป้อมปราบศัตรูพ่า', 'A_NAME_T'] = 'ป้อมปราบศัตรูพ่าย'
shpdf.loc[shpdf['A_NAME_T'] == 'เมืองนครศรีธรรมร', 'A_NAME_T'] = 'เมืองนครศรีธรรมราช'
shpdf.loc[shpdf['A_NAME_T'] == 'เมืองสุราษฎร์ธาน', 'A_NAME_T'] = 'เมืองสุราษฎร์ธานี'
shpdf.loc[shpdf['A_NAME_T'] == 'กรงปีนัง', 'A_NAME_T'] = 'กรงปินัง'
shpdf.loc[shpdf['A_NAME_T'] == 'ป่าพยอม', 'A_NAME_T'] = 'ป่าพะยอม'
shpdf.loc[shpdf['A_NAME_T'] == 'เมืองประจวบคีรีข', 'A_NAME_T'] = 'เมืองประจวบคีรีขันธ์'
shpdf.loc[shpdf['A_NAME_T'] == 'ว่านใหญ่', 'A_NAME_T'] = 'หว้านใหญ่'
shpdf.loc[shpdf['A_NAME_T'] == 'บึงกาฬ', 'A_NAME_T'] = 'เมืองบึงกาฬ'

# shpdf['AMP_NAME'][shpdf['AMP_NAME']=='เมือง']=shpdf['AMP_NAME'][shpdf['AMP_NAME']=='เมือง']+shpdf['PRV_NAME'][shpdf['AMP_NAME']=='เมือง']
# merge = pd.merge(df, shpdf, how='left', left_on=['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'],
#                  right_on=['pv_tn', 'ap_tn'])
# merrgefail = merge[['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']][merge['ap_tn'].isna()]


# shpdf['A_NAME_T']=shpdf['A_NAME_T'].str.strip()
merge = pd.merge(df, shpdf, how='left', left_on=['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง'],
                 right_on=['P_NAME_T', 'A_NAME_T'])
merrgefail = merge[['จังหวัด', 'ท้องที่ที่ประกอบเป็นเขตเลือกตั้ง']][merge['T_NAME_T'].isna()]

shpdf['A_NAME_T'][(shpdf['P_NAME_T'] == 'ยะลา') & shpdf['A_NAME_T'].str.contains('สนาม')]  ##fortest

# subdistbkkshp = shptodf('subdist_bma/subdist_bma.shp')
# subdistbkkshp["SNAME"] = subdistbkkshp["SNAME"].replace("^แขวง", "", regex=True)
# subdistbkkshp["DNAME"] = subdistbkkshp["DNAME"].replace("^เขต", "", regex=True)

onlydf=pd.DataFrame(df.interior[df.interior != ''].tolist(), index=df.จังหวัด[df.interior != '']).stack().reset_index(name='interior')[['interior','จังหวัด']]
onlymerge = pd.merge(onlydf, shpdf, how='left', left_on=['จังหวัด', 'interior'],
                 right_on=['P_NAME_T', 'T_NAME_T'])
onlymerrgefail = onlymerge[['จังหวัด', 'interior']][onlymerge['A_NAME_T'].isna()]


[{"id": "6001", "name": "เมืองนครสวรรค์",
  "condition": {"id": "600106", "name": "นครสวรรค์ตก", "divisonType": "ตำบล", "conditionType": "interior",
                "condition": {"id": "600106", "name": "นครนครสวรรค์"}}}]

th_map_df=pd.read_csv("th_map.csv")
th_map_df['id'] = th_map_df['id'].apply(str)
def getGeoCode(P, A=None, T=None):
    if P and A==None and T==None:
        return th_map_df.loc[(th_map_df['id'].str.len() == 2) &
                             (th_map_df['name'] == P), 'id'].iloc[0]
    if P and A and T==None:
        p_id=getGeoCode(P)
        return th_map_df.loc[(th_map_df['id'].str.len() == 4) &
                             (th_map_df['id'].str.startswith(p_id)) &
                             (th_map_df['name'] == A), 'id'].iloc[0]
    if P and A and T:
        a_id = getGeoCode(P, A)
        return th_map_df.loc[(th_map_df['id'].str.len() == 6) &
                             (th_map_df['id'].str.startswith(a_id)) &
                             (th_map_df['name'] == T), 'id'].iloc[0]


