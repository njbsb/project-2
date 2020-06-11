import pandas as pd
import os
from datetime import date, datetime
from dateutil.relativedelta import relativedelta


def date_to_duration(date):
    dip = date.date()
    rdelta = relativedelta(date.today(), dip)
    dateYM = str(rdelta.years) + "y" + str(rdelta.months) + "m"
    return dateYM


def drop_unusedColumn(df, kind):
    print("dropping unused column...")
    droplist = []
    if kind == "zh":
        droplist = [6, 7, 9, 31, 36, 39, 40, 41, 42]
    else:
        pass
        # droplist = [6]
        # keeplist = [0, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11,
        #             12, 13, 21, 22, 23, 24, 25, 26, 27, 28]
        # droplist = [16, 17, 18, 19, 20]
    df = df.drop(columns=droplist)
    return df


def reindex_column(df):
    colname = list(df.columns)
    for i in range(len(colname)):
        colname[i] = i
    df.columns = colname
    # print(df.head())
    return df


def changeColumnName(df, typ):
    print("changing column name...")
    z_column = [ec[0] for i, ec in df.iteritems()]
    df = df.drop([0]).reset_index(drop=True)
    df.columns = z_column
    if typ == 'zp':
        df = df.rename(columns={"Personnel Number": "Staff No", "Formatted Name of Employee or Applicant": "Staff Name", "Sal. Grade": "SG", "Job Grade": "Conso JG",
                                "Company": "Comp. Position", "Pos. SKG": "Pos. SKG Zpdev", "Home SKG": "Home SKG Zpdev", "Gender Key Desc": "Gender"})
    return df


def rearrange_column(df_zh, df_zp):
    print("rearranging columns...")
    print(datetime.now())
    # print(df_zp.columns)
    df_zh = changeColumnName(df_zh, 'zh')
    df_zp = changeColumnName(df_zp, 'zp')
    # print(df_zh.head(10))
    # print(df_zp.head(10))
    print("merging both dataframes...")
    df = pd.merge(df_zh, df_zp, on=['Staff No', 'Staff Name', 'SG', 'Lvl', 'Conso JG', 'Position',
                                    'Section', 'Department', 'Division', 'Sector', 'Comp. Position', 'Gender'], how='outer')
    print("merging done")
    # df1 = pd.merge(df_zh, df_zp, on=[
    #                'Staff No', 'Staff Name', 'SG'], how='outer')
    # df2 = pd.merge(df_zh, df_zp, on=[
    #                'Staff No', 'Staff Name', 'SG', 'Lvl'], how='outer')

    # df = pd.merge(df_zh, df_zp, on=['Staff No', 'Staff Name', 'SG', 'Lvl', 'JG', 'Position',
    #                                 'Section', 'Department', 'Division', 'Sector', 'Comp. Position', 'Gender'], how='outer')
    # df = df_zh.merge(df_zp, 'outer')
    '''
    ori_colname = list(df.columns)
    df = reindex_column(df)  # convert column name to index
    final = [0, 1, 33, 25, 26, 35, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 51, 63, 64, 36, 37, 38, 39, 40, 41, 52, 53, 54, 55, 56, 42, 43, 44,
             45, 46, 47, 48, 49, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 27, 28, 57, 58, 29, 30, 31, 32, 59, 34, 60, 61, 62]
    if(len(ori_colname) == len(final)):
        final_colname = [ori_colname[i] for i in final]
        final_colname[2] = 'OT Status'
    df = df.reindex(columns=final)  # arrange the columns
    df = reindex_column(df)
    # print('dated', df[[15, 25, 26, 27, 28, 29, 30, 64]])
    # need to reindex column first
    indexwithdate = [15, 25, 26, 27, 28, 29, 30, 64]
    df = datetime_toDate(df, indexwithdate)
    df.columns = final_colname
    '''
    df = df.sort_values(by=['Pos ID'], ascending=True)
    return df


def datetime_toDate(df, indexwithdate):
    print('adjusting date')
    for i, column in df.iteritems():
        if(i in indexwithdate):
            # loop through column
            for j, row in column.iteritems():
                if type(row) == datetime:
                    column.iloc[j] = row.date()
                else:
                    if type(row) == str and ' ' in row:
                        column.iloc[j] = row.replace(' ', '')
                    # else would be str or float for nan
        elif i == 2:
            for j, row in column.iteritems():
                column.iloc[j] = ''


def correct_JG_ZP(df):
    print("Correcting JG...")
    print(df.head(10))
    for i, row in df.iloc[:, 6].iteritems():
        if not i == 0:
            if type(row) == str:
                slash = row.find('/')
                jg1, jg2 = row[:slash], row[slash+1:]
                if(jg1 == jg2):
                    df.iloc[i, 6] = jg1
    print(df)
    return df


directory = os.getcwd()
opt = input("insert input: \n1. Full DB\n2. Partial DB\nAnswer: ")
if opt == '1':
    zhpla_path = r'D:\Documents\Python\files\ZHPLA_MAY20.xlsx'
    zpdev_path = r'D:\Documents\Python\files\ZPDEV_MAY20.xlsx'
    df_zh = pd.read_excel(zhpla_path, header=None,
                          sheet_name='Sheet1', skiprows=3)
    df_zp = pd.read_excel(zpdev_path, header=None,
                          sheet_name='Sheet1', skiprows=3)
    filnam = "TA_input1.csv"
elif opt == '2':
    zhpla_path = r'D:\Documents\Python\files\zh.xlsx'
    zpdev_path = r'D:\Documents\Python\files\zp.xlsx'
    df_zh = pd.read_excel(zhpla_path, header=None, sheet_name='Sheet1')
    df_zp = pd.read_excel(zpdev_path, header=None, sheet_name='Sheet1')
    filnam = "TA_input2.csv"


df_zp[40] = df_zp[29]
df_zp[41] = df_zp[30]
for i, column in df_zp.iteritems():
    if i == 40 or i == 41:
        for index, row in column.iteritems():
            if not index == 0:
                column[index] = date_to_duration(row)
            else:
                if(i == 40):
                    column[index] = "YiPetronas"
                else:
                    column[index] = "YiPosition"

df_zh = drop_unusedColumn(df_zh, 'zh')
df_zp = correct_JG_ZP(df_zp)

df_merge = rearrange_column(df_zh, df_zp)

# print(df_merge)
print('writing csv...')
df_merge.to_csv(filnam, header=True, index=False)
print('DONE')

# to recorrect company name
# incorrect_list = ['AET Inc. Limited', 'Akademi Laut Malaysia', 'Dewan Filharmonik PETRONAS', 'Eaglestar Marine Holdings (L) Pte. Ltd.', 'Eaglestar Shipmanagement (L) Pte.Ltd.', 'Engen DRC SARL (0097)',
#                   'Engen Ghana Limited (0074)', 'Engen Kenya Ltd (0087)', 'Engen Petroleum (Burundi) SA (0101)', 'Engen Petroleum (Mauritius) Ltd (0105)', 'Engen Petroleum (Zimbabwe) Ltd (0015)',
#                   'Engen Petroleum Ltd (0011)', 'Engen Petroleum Ltd (0012)', 'Engen Reunion SA (0106)', 'Engen Swaziland (Pty) Ltd (0044)',
#                   'Gas District Cooling (M) Sdn Bhd', 'Gas District Cooling (Putrajaya) Sdn Bhd', 'Japan Malaysia LNG Co. Ltd.', 'Kimanis Power Sdn Bhd',
#                   'KLCC Parking Management Sdn Bhd', 'KLCC Projeks Sdn Bhd', 'KLCC Property Holdings Berhad', 'KLCC Urusharta Sdn Bhd', 'Malaysia Marine & Heavy Eng. Holding',
#                   'Malaysia Marine & Heavy Engineering-New', 'Malaysian Philharmonic Orchestra', 'Malaysian Refining Co Sdn Bhd', 'MISC Berhad',
#                   'MISC Maritime Services Sdn.Bhd.', 'MMHE EPIC Marine & Services Sdn Bhd', 'PC Gabon Upstream SA', 'PC Mauritania Pty Ltd',
#                   'PC Muriah Ltd', 'Pengerang Power Sdn Bhd', 'Pengerang Refining Company Sdn. Bhd.', 'PETCO Trading DMCC',
#                   'PETCO Trading Labuan Co. Ltd', 'PETRONAS (E&P) Overseas Ventures Sdn Bhd', 'PETRONAS Chemicals Ammonia Sdn Bhd', 'PETRONAS Chemicals Aromatics Sdn Bhd',
#                   'PETRONAS Chemicals Derivatives Sdn Bhd', 'PETRONAS Chemicals Fertiliser Sabah S B', 'PETRONAS Chemicals Glycols Sdn Bhd', 'PETRONAS Chemicals Isononanol SB',
#                   'PETRONAS Chemicals LDPE Sdn Bhd', 'PETRONAS Chemicals Marketing (Labuan) Lt', 'PETRONAS Chemicals Marketing Sdn Bhd', 'PETRONAS Chemicals Methanol Sdn Bhd',
#                   'PETRONAS Chemicals Olefins Sdn Bhd', 'PETRONAS Chemicals Polyethylene Sdn Bhd', 'PETRONAS E&P Argentina SA', 'PETRONAS Energy Trading Ltd.',
#                   'PETRONAS Intnl Marketng (Thailand)Co Ltd', 'PETRONAS LNG Ltd', 'PETRONAS LNG Sdn Bhd', 'PETRONAS Lub Marketing Malaysia SB',
#                   'PETRONAS Research Sdn Bhd', 'PETRONAS Suriname E&P BV', 'PETRONAS Technical Training Sdn Bhd', 'Petrosains Sdn Bhd',
#                   'PICL (Egypt) Corporation Ltd', 'PRPC Utilities and Facilities Sdn. Bhd.', 'Putrajaya Ventures Sdn Bhd', 'Regas. Terminal (Sg. Udang) Sdn. Bhd',
#                   'Techno Indah', 'VPSB'
#                   ]
# correct_list = ['@AET', '@ALAM', 'DFP', '@EAGLESTARMA', '@EAGLESTARSH', '@Engen103G', '@Engen103E', '@Engen103F', '@Engen374H', '@Engen374I', '@Engen374F', '@Engen374B', '@Engen374C', '@Engen103J', '@Engen103D', 'GDCSB',
#                 'GDC(PJ)', 'JAMALCO', '@KPSB', 'KLCCPM', 'KLCC', 'KLCCPHB', 'KLCCUH', '@MHB', '@MMHE', 'MPO', 'MRCSB', '@MISC', '@MMS', '@MEMS', '@PCGUSA', '@PCMPL', 'PCINO', 'PePSB', 'Pengerang RC', 'PTDMCC', 'PTLCL', 'PEPOV',
#                 'PC Ammonia', 'PCARO', 'PCDSB', 'PCFSSB', 'PCGSB', 'PC INA', 'PC LDPE', 'PCML', 'PCM', 'PC Methanol', 'PCOSB', 'PCPSB', '@PEPASA', 'PETL', '@PIMTCL', 'PLL', 'PLSB', 'PLMMSB', 'PRSB', '@PSEPBV', 'PTTSB', 'PSB',
#                 '@PICL EGYPT', 'PRPC UF SB', 'PJVSB', 'RGT(SU)SB', '@TechnoIndah', '@VPSB']
# print(len(incorrect_list), len(correct_list))
# company_dict = {incorrect_list[i]: correct_list[i]
#                 for i in range(len(incorrect_list))}
# df_zp = df_zp.replace({10: company_dict})
# print("rename completed")
# to remove the semicolon from ZP dataframe, also include one position
# for i, column in df_zp.iteritems():
#     if i in [5, 6, 7, 8, 9]:  # for position, section, dept, div, sector
#         for j, row in column.iteritems():
#             if not j == 0 and type(row) == str:
#                 if row == 'Chargehand, B;asting and Painting':
#                     column.iloc[j] = row.replace(';', 'l')
#                 else:
#                     column.iloc[j] = row.replace(';', ',')
# print('remove semicolon completed')


# print("original number of ZH row: ", len(df_zh) - 1)
# print("original number of ZP row: ", len(df_zp) - 1)

# id_zh = [row for i, row in df_zh.iloc[:, 0].iteritems() if not i ==
#          0 and not type(row) == float]
# id_zp = [row for i, row in df_zp.iloc[:, 0].iteritems() if not i ==
#          0 and not type(row) == float]
# print("ZH | no of valid id: {}, invalid/nan: {}".format(len(id_zh),
#                                                         (len(df_zh.iloc[:, 0]) - 1 - len(id_zh))))
# print("ZP | no of valid id: {}, invalid: {}".format(
#     len(id_zp), (len(df_zp.iloc[:, 0]) - 1 - len(id_zp))))

# zh_special = [staffz for staffz in id_zh if staffz not in id_zp]
# zp_special = [staffz for staffz in id_zp if staffz not in id_zh]
# print("\nOnly exist in ZH: ", len(zh_special))
# print("Only exist in ZP: ", len(zp_special))
