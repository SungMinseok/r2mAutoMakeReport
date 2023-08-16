import pandas as pd
from datetime import datetime
import os

def save_qa_info_to_csv(excel_filename):
    #excel_filename = f'CL_{country}_{project_name}_{date.strftime("%y%m%d")} {patchtype} QA.xlsx'
    
    # Load Excel sheet
    df = pd.read_excel(excel_filename, sheet_name='TEST REPORT', header=None)

    # Extract information
    project_name = df.loc[4, 2]
    #start_date = pd.to_datetime(df.loc[6, 2])
    korean_days = ["일", "월", "화", "수", "목", "금", "토"]
    start_date = df.loc[6, 2]
    day_of_week = start_date.strftime("%w")
    start_date = start_date.strftime("%Y.%m.%d") + f"({korean_days[int(day_of_week)]})"
    end_date = df.loc[6, 6]
    day_of_week = end_date.strftime("%w")
    end_date = end_date.strftime("%Y.%m.%d") + f"({korean_days[int(day_of_week)]})"
    # team_size = df.loc[7, 2]
    # alpha_server = df.loc[29, 2]
    # live_server = df.loc[31, 2]
    # alpha_client_0 = df.loc[29, 5]
    # alpha_client_1 = df.loc[30, 5]
    # live_client_0 = df.loc[31, 5]
    # live_client_1 = df.loc[32, 5]
    # live_client_2 = df.loc[33, 5]
    # success_rate = "{:.2%}".format(round(df.loc[11, 2], 4))#df.loc[11, 2]
    # execution_rate = "{:.2%}".format(round(df.loc[11, 3], 4))#df.loc[11, 3]
    # pass_count = df.loc[11, 4]
    # fail_count = df.loc[11, 5]
    # na_count = df.loc[11, 6]
    # nt_count = df.loc[11, 7]
    # ne_count = df.loc[11, 8]
    # total_count = df.loc[11, 9]


    # Create a dictionary to hold the data
    qa_info = {
        'PROJECT_NAME': df.loc[4, 2],
        'PROJECT_NAME_1': df.loc[4, 2].replace(' QA',''),
        'START_DATE': start_date,
        'END_DATE': end_date,
        'MEMBERS': df.loc[7, 2],
        'SERVER_ALPHA': df.loc[11, 2],
        'SERVER_LIVE': df.loc[13, 2],
        'CLIENT_ALPHA_0': df.loc[11, 5],
        'CLIENT_ALPHA_1': df.loc[12, 5],
        'CLIENT_LIVE_0': df.loc[13, 5],
        'CLIENT_LIVE_1': df.loc[14, 5],
        'CLIENT_LIVE_2': df.loc[15, 5],
        'SUCCESS_RATE': "{:.2%}".format(round(df.loc[20, 2], 4)),
        'EXECUTION_RATE': "{:.2%}".format(round(df.loc[20, 3], 4)),
        'PASS_COUNT': df.loc[20, 4],
        'FAIL_COUNT': df.loc[20, 5],
        'NA_COUNT': df.loc[20, 6],
        'NT_COUNT': df.loc[20, 7],
        'NE_COUNT': df.loc[20, 8],
        'TOTAL_COUNT': df.loc[20, 9]
    }

    # Convert the dictionary to a DataFrame and save to CSV
    qa_df = pd.DataFrame.from_dict(qa_info, orient='index', columns=['Value'])
    qa_df.to_csv('qa_info.csv', index_label='Key', encoding='utf-8-sig')
    import os
    #os.startfile('qa_info.csv')

if __name__  == "__main__" :
    # 호출 예시
    #country = 'TW'
    #date = pd.to_datetime('2023-08-16')
    save_qa_info_to_csv('CL_TW_R2M_230816 업데이트 QA.xlsx')
