import pandas as pd
from xlsx_processing import *
import os
import shutil
import psutil
import time 
import numpy as np
def merge_rows_by_category1(df):
    grouped_df = df.groupby('Category1').agg(lambda x: ' '.join(x)).reset_index()
    return grouped_df

def create_checklist(sheet_name , input_file, output_file, criterion, required_parts):
    '''
    체크리스트생성
    '''
    # Read the input Excel file into a Pandas DataFrame
    #df = pd.read_excel(input_file, sheet_name)

    df = pd.read_excel(input_file, sheet_name)
    #xls = pd.ExcelFile(input_file)
    #df = pd.read_excel(xls, sheet_name='퀘스트')


        # "a"열의 길이를 확인합니다
    df[criterion] = df[criterion].replace(r'^\s*$', np.nan, regex=True)

# Count non-NaN values in the modified column
    first_column_length = df[criterion].notna().sum()
    #first_column_length = (df[criterion].notna() & (df[criterion] != ' ')).sum()
    print(first_column_length5)
    # "a"열의 길이만큼 데이터를 추출합니다
    df = df.iloc[:first_column_length, :]




    df[criterion] = df[criterion].fillna(method='ffill')
    
    # Check if all columns are required
    if required_parts == 'all':
        required_parts = df.columns

    # Convert the DataFrame into the desired checklist format
    result_data = []
    for _, row in df.iterrows():
        name = row[criterion]
        for part in required_parts:
            if isinstance(part, tuple):
                # Handle tuples by combining their values in one row
                combined_value = ' x'.join(str(row[p]) for p in part )
                result_data.append([name, ' & '.join(part), combined_value])
            else:
                if part != criterion :
                    value = row[part]
                    result_data.append([name, part, value])

    # Create the resulting DataFrame
    result_df = pd.DataFrame(result_data, columns=['Category1', 'Category2', 'Category3'])
    
    # Merge rows with the same values in Category1
    #result_df = merge_rows_by_category1(result_df)

    # Save the result to an output Excel file
    result_df.to_excel(output_file, index=False)

    postprocess_cashshop(output_file)

def create_checklist2(input_file, output_file, criterion, required_parts):
    '''
    체크리스트생성
    '''
    # Read the input Excel file into a Pandas DataFrame
    df = pd.read_excel(input_file)

    id_list = df[criterion].dropna(axis=0)
    id_index_list = id_list.index
    totalCount = len(id_index_list)
    #df[criterion] = df[criterion].fillna(method='ffill')
    
    # Check if all columns are required
    if required_parts == 'all':
        required_parts = df.columns

    result_data = []

    for i in tqdm(range(0,totalCount)):
        #print(cashShopIdIndexList[j], j+1)

        if (i+1) >= totalCount :
            tempDf = df[id_index_list[i]:]
        else :
            tempDf = df[id_index_list[i]:id_index_list[i+1]]
        tempDf = tempDf.reset_index()

        split_char = 'x'
        temp_val = split_char.join(tempDf['능력치'].dropna().values)

        print(tempDf['능력치'].dropna().values)

        tempDf = tempDf[:1]
        tempDf['능력치'] = temp_val

        
        result_data.append(tempDf)


    # Create the resulting DataFrame
    #result_df = pd.DataFrame(result_data, columns=['Category1', 'Category2', 'Category3'])
    
    # Merge rows with the same values in Category1
    #result_df = merge_rows_by_category1(result_df)

    # Save the result to an output Excel file
    #result_df.to_excel(output_file, index=False)

    postprocess_cashshop(output_file)


#file_to_close = "insert2.xlsx"  # 종료하려는 파일명

def 파일종료(파일명) :
    for process in psutil.process_iter(attrs=['pid', 'name']):
        try:
            process_info = process.info()
            if process_info['name'] == "EXCEL.EXE":
                open_files = process.open_files()
                for file_info in open_files:
                    if file_info.path.endswith(파일명):
                        # 주어진 파일을 열고 있는 엑셀 프로세스를 종료
                        pid = process_info['pid']
                        process = psutil.Process(pid)
                        process.terminate()
                        print(f"{파일명}을(를) 종료합니다.")
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass



# Example usage:
if __name__ == "__main__":
    input_file = fr"c:\Users\mssung\Documents\insert3.xlsx"
    output_file = f"result{time.strftime('_%y%m%d_%H%M%S')}.xlsx"
    criterion = '이름'
    criterion = '아이템 명'
    required_parts = 'all'#['미믹 코인','대성공1', '대성공2', '대성공3']#[('능력치', '수치'), '등록 성공 확률']
    #required_parts = ['처치 수', '경험치',('아이템명1','수량1'),('아이템명2','수량2'),('아이템명3','수량3'),'스택획득확률']#[('능력치', '수치'), '등록 성공 확률']
    #required_parts = ['이름', '경험치']
    #required_parts = ['아이템 설명','아이템 사용','비고']#'all'#['등급', '나이', '가격']
    #required_parts = ['등급','설명','아이템 사용','무브리밋']#'all'#['등급', '나이', '가격']

    sheet_name = "아이템CL(수동)"
    sheet_name = "제작CL"
    sheet_name = "상자CL"
    sheet_name = "스킬CL"
    #sheet_name = "새거 (2)"
    sheet_name = "영혼부여"
    #sheet_name = '세트효과'
    #sheet_name = "아이템CL(자동)"
    #sheet_name = "퀘스트"
    #sheet_name = "아이템CL(자동)"
    #sheet_name = "제작CL"

    if sheet_name == "아이템CL(수동)" :
        '''
        아이템CL(수동시트)
        '''
        criterion = '이름'
        required_parts = ['등급','이미지','설명','무브리밋','아이템 사용']#아이템CL
    elif sheet_name == "제작CL" :
        '''
        제작CL
        '''
        criterion = '제작 아이템'
        required_parts = ['카테고리','제작 수량','제작 제한','골드 비용','재료','성공 확률','실제 제작']#아이템CL
        required_parts = ['제작 제한','골드 비용','성공 확률','실제 제작','거래소여부','재료']#아이템CL
        required_parts = ['카테고리', '제작 제한','골드 비용','재료','성공 확률','실제 제작','제작 아이템 사용','완성아이템 거래가능여부']#아이템CL
    elif sheet_name == "상자CL" :
        '''
        상자CL
        '''
        criterion = '이름'
        required_parts = ['구성품 획득','뽑기 구성']#아이템CL
    elif sheet_name == "스킬CL" :
        '''
        스킬CL
        '''
        criterion = '스킬명'
        required_parts = ['스킬명','스킬 아이콘','스킬 획득 가능 레벨','대상','소모 HP(1미만=%)','소모품','쿨타임(초)','지속시간(초)','능력치 텍스트','상태이상(능력치) 적용', '상태이상 버프 아이콘']#아이템CL
        
    elif sheet_name == "스킬CL" :
        '''
        스킬강화CL
        '''
        criterion = '스킬명'
        required_parts = ['강화재료','필요재화량','강화 확률(%)']#아이템CL
        
    # elif sheet_name == "새거 (2)" :
    #     '''
    #     스킬강화CL
    #     '''
    #     criterion = '클래스 구분'
    #     required_parts = ['상자','획득 아이템']#아이템CL
        
    elif sheet_name == "영혼부여" :
        '''
        영혼부여
        '''
        criterion = '결과 장비'
        required_parts = ['필요 영혼석','재료 장비','영혼 부여 비용','영혼 부여 시도']#아이템CL
        #required_parts = ['영혼 부여 확률']#아이템CL
        
    elif sheet_name == "세트효과" :
        '''
        영혼부여
        '''
        criterion = '이름'
        required_parts = ['세트효과명', '장비 세트 능력치','세트 아이템']#아이템CL
        
    elif sheet_name == "아이템CL(자동)" :
        '''
        아이템CL(자동시트)
        '''
        criterion = '이름'
        required_parts = ['등급','이미지','설명','무브리밋','아이템 사용']#아이템CL
    elif sheet_name == "퀘스트" :
        '''
        퀘스트
        '''
        criterion = '이름'
        required_parts = ['퀘스트 목표','퀘스트 내용','경험치', '보상1', '보상2']#'all'#['등급','이미지','설명','무브리밋','아이템 사용']#아이템CL
    #elif sheet_name == "제작CL" :
    # elif sheet_name == "새거 (2)" :

    #     '''
    #     기본
    #     '''
        
    #     criterion = '아이템 명'
    #     required_parts ='all'
        
    # criterion = '이름'
    # required_parts = 'all'
    #create_checklist2(input_file, output_file, criterion, required_parts)

    # # 원본 엑셀 파일과 복사본 파일 경로 설정
    # original_file = 'insert2.xlsx'
    # copy_file = 'copy.xlsx'

    # # 원본 파일을 복사본 파일로 복사
    # shutil.copy(original_file, copy_file)

    # 복사본 파일을 읽어옴
    #df = pd.read_excel(copy_file, sheet_name)
    #파일종료('insert2.xlsx')
    create_checklist(sheet_name,input_file, output_file, criterion, required_parts)
    #os.startfile('insert2.xlsx')
    os.startfile(output_file)

