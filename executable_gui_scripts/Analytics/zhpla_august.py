import pandas as pd


def addPossum(df):
    df['Pos SUM'] = df[['Position', 'Section', 'Department', 'Division', 'Sector',
                        'Comp. Position', 'Business']].apply(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
    return df


def addcolumn_superior(df):
    df_staff = df[['Pos ID', 'Staff No', 'Staff Name']].drop_duplicates()
    df_staff.rename(columns={'Pos ID': 'Superior Pos ID',
                             'Staff No': 'Superior ID', 'Staff Name': 'Superior Name'}, inplace=True)
    newdf = pd.merge(df, df_staff, on='Superior Pos ID', how='left')
    return newdf


def addcolumn_consoJG(df):
    column_consoJG = []
    jg_head = ['', 'Est. ', 'Eqv. ']
    df_jg = df[['JG', 'Est.JG', 'EQV JG']].fillna('-')
    for i, eachrow in df_jg.iterrows():
        for index, val in enumerate(list(eachrow)):
            if not val == '-':
                conso = jg_head[index] + str(val)
                column_consoJG.append(conso)
                break
            else:
                if(index == 2 and val == '-'):
                    column_consoJG.append('No JG')
                else:
                    continue
    df['Conso JG'] = column_consoJG
    return df


path = r'D:\Documents\Python\project-2\inputfile\ZHPLA_august 2020.xlsx'
df_zh = pd.read_excel(path)

sub_df = df_zh[df_zh.iloc[:, 0] == 'ZHPLA508 Enhanced Report']
list_index = sub_df.index.tolist()
df_group = []
for i, index in enumerate(list_index):
    if i == len(list_index) - 1:
        sub_df = df_zh.iloc[index + 1:, 1:]
    else:
        sub_df = df_zh.iloc[index + 1: list_index[i + 1], 1:]
    for j in [0, 1]:
        sub_df = sub_df.dropna(how='all', axis=j)
    sub_df = sub_df.rename(columns=sub_df.iloc[0]).drop(sub_df.index[0])

    sub_df = addPossum(sub_df)
    sub_df = addcolumn_superior(sub_df)
    sub_df = addcolumn_consoJG(sub_df)
    sub_df = sub_df.reindex(columns=['Staff No', 'Staff Name', 'SG', 'Lvl', 'Tier', 'JG', 'Est.JG', 'EQV JG', 'Conso JG', 'OLIVE JG',
                                     'Pos ID', 'Pos SUM', 'Superior Pos ID', 'Superior ID', 'Superior Name', 'Position', 'Sec ID',
                                     'Section', 'Dept ID', 'Department', 'Division ID', 'Division', 'Sector ID', 'Sector', 'Comp. ID Position',
                                     'Comp. Position', 'Buss ID', 'Business', 'C.Center', 'Gender', 'Race', 'Zone', 'Home SKG', 'Pos.SKG', 'JobID( C)',
                                     'Level', 'NE Area', 'Loc ID', 'Location Text', 'Vac Date', 'Start Date', 'End Date', 'Vacancy Status',
                                     'Email Address', 'Mirror Pos ID', 'Mirror Position', 'EG Post. ID', 'EG Position', 'ESG Post. ID', 'ESG Position',
                                     'PT1', 'PT2', 'Overseas Staff No', 'Overseas Staff Name', 'Overseas Staff Comp. ID', 'Overseas Staff Comp.',
                                     'Staff Comp. ID', 'Staff Comp.', 'Staff EG ID', 'Staff EG', 'Staff ESG ID', 'Staff ESG', 'Staff Work Contract ID', 'Staff Work Contract'
                                     ])
    df_group.append(sub_df)

df_zhclean = pd.concat(df_group).reset_index(drop=True)
df_zhclean.to_excel('august_zhpla.xlsx', index=None)
print(df_zhclean)
