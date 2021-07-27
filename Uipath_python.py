def split_excel_equally(filepath, split_number):
    try:
        import pandas as pd
        from time import strftime

        file_location = '\\'.join(filepath.split('\\')[:-1])

        log_list = [strftime("%d/%m/%Y %H:%M:%S") + '- Inside split_excel_equally']
        df = pd.read_excel(filepath)

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Step 1: Input Excel is read')
        limit = 0

        df1 = pd.DataFrame()  # Initialize DF

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- Step 2 : DF is initialized")

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- Step 3 : Loop through the split_number count begins")

        for count in range(1, split_number + 1):
            df1 = df1[0:0]  # for making the DF null

            if count != split_number:
                for i in range(len(df) // split_number):
                    # df1 = df1.append(df.loc[[limit]].squeeze())
                    df1 = df1.append(df.loc[[limit]])
                    limit = limit + 1
            else:
                df1 = df[limit:]  # For pushing the last batch of excel
            # df1.to_excel(r"C:\Users\Dell\Documents\UiPath\Uipath_python\output" + str(count) + ".xlsx", index=False)
            writer = pd.ExcelWriter(file_location + "\\output" + str(count) + ".xlsx", engine='xlsxwriter')
            df1.to_excel(writer, sheet_name='output', index=False)
            writer.save()

            log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- Output file created : output" + str(count) + ".xlsx")
            log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- Total number of records : " + str(len(df1)))

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- File splitting completed successfully")

        with open(file_location + '\\app_log.log', 'w+') as f:
            for item in log_list:
                f.write('%s\n' % item)

        f.close()

        return log_list
    except Exception as e:
        log_list.append("Exception caught : " + e)
        return log_list


def split_excel_percent(filepath, percent_list):
    import pandas as pd
    import logging
    from time import strftime

    file_location = '\\'.join(filepath.split('\\')[:-1])

    df = pd.read_excel(filepath)
    limit = 0

    log_list = [strftime("%d/%m/%Y %H:%M:%S") + " Step 1 : Input excel is read"]

    df_actual = pd.DataFrame()  # Initialize DF

    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " Step 2 : DF is initialized")
    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " Step 3 : Loop through the percent value count begins")

    for count in range(len(percent_list)):
        df_actual = df_actual[0:0]  # for making the DF null

        if count != len(percent_list) - 1:
            for i in range(int(int(percent_list[count]) / 100 * len(df))):
                df_actual = df_actual.append((df.loc[[limit]].squeeze()))
                limit = limit + 1
        else:
            df_actual = df[limit:]

        # df_actual.to_excel(r"C:\Users\Dell\Documents\UiPath\Uipath_python\output" + str(count) + ".xlsx", index=False)
        writer = pd.ExcelWriter(file_location+"\\output"+str(count)+".xlsx", engine='xlsxwriter')
        df_actual.to_excel(writer, sheet_name='output', index=False)
        writer.save()

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " Output file created : output" + str(count) + ".xlsx")

        log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " No. of records in the file : " + str(len(df_actual)))

    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " Total number of output files created : " + str(count-1))

    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + " File splitting completed successfully")

    with open(file_location + '\\app_log.log', 'w+') as f:
        for item in log_list:
            f.write('%s\n' % item)

    f.close()

    return log_list


def filter_excel_manual(filepath, column_name, column_values):
    import pandas as pd

    file_location = '\\'.join(filepath.split('\\')[:-1])

    # logging.basicConfig(filename=r"" + '\\'.join(filepath.split('\\')[:-1]) + "\\app.log", filemode='w',
    #                     format='%(name)s - %(levelname)s - %(message)s')

    log_list = ['read df']
    df = pd.read_excel(filepath)

    log_list.append('starting log')

    if int(column_name[-1]) == 0:

        log_list.append('Generating single file as ' + column_name[1] + 'is selected in Config')
        df_filtered = df[df[column_name[0]].isin(column_values)]
        # df_filtered.to_excel(file_location + "\\filtered_output.xlsx", index=False)
        writer = pd.ExcelWriter(file_location + "\\filtered_output.xlsx", engine='xlsxwriter')
        df_filtered.to_excel(writer, sheet_name='output', index=False)
        writer.save()

        log_list.append('Output file generated : filtered_output.xlsx')
        log_list.append('Total records : ' + str(len(df_filtered)))

    else:
        log_list.append('Generating multiple files as ' + column_name[1] + ' is selected in Config')
        count = 1
        for value in column_values:
            df_filtered = df[df[column_name[0]] == value]
            # df_filtered.to_excel(file_location + "\\filtered_output" + str(count) + ".xlsx",
            #                      index=False)
            writer = pd.ExcelWriter(file_location + "\\filtered_output" + str(count) + ".xlsx", engine='xlsxwriter')
            df_filtered.to_excel(writer, sheet_name='output', index=False)
            writer.save()

            log_list.append('Output file generated : ' + "filtered_output" + str(count) + ".xlsx")
            log_list.append('Total records : ' + str(len(df_filtered)))
            count = count + 1
        log_list.append('Total files generated : ' + str(count - 1))

    with open(file_location + '\\app_log.log', 'w+') as f:
        for item in log_list:
            f.write('%s\n' % item)

    f.close()

    return log_list


def filter_excel_sql(filepath, sql_query):
    import sqlite3
    import pandas as pd
    from time import strftime

    file_location = '\\'.join(filepath.split('\\')[:-1])

    # logging.basicConfig(filename=r"" + file_location + "\\app.log", filemode='w',
    #                     format='%(name)s - %(levelname)s - %(message)s')

    log_list = ['Python logs']

    con = sqlite3.connect(file_location + '\\database.db')
    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Connecting to staging database - Complete')

    df = pd.read_excel(filepath)
    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + "- read excel to df - Complete")

    df.to_sql(name='temp_table', con=con, if_exists='replace', index=True)
    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Convert df to sql table - Complete')

    df_sql = pd.read_sql_query(sql_query, con)
    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Running the sql query - Complete')

    # df_sql.to_excel(file_location + '\\sql_output.xlsx', index=False)
    writer = pd.ExcelWriter(file_location + "\\filtered_output.xlsx", engine='xlsxwriter')
    df_sql.to_excel(writer, sheet_name='output', index=False)
    writer.save()

    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Converting the sql table to excel - Complete')

    log_list.append(strftime("%d/%m/%Y %H:%M:%S") + '- Number of records in the output file : ' + str(len(df_sql)))

    with open(file_location + '\\app_log.log', 'w+') as f:
        for item in log_list:
            f.write('%s\n' % item)

    f.close()

    con.commit()
    con.close()

    return log_list
