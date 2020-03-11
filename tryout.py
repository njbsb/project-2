from createsp import df_list

for df in df_list:
    print(max(enumerate(df), key = lambda tup: len(tup[1])))