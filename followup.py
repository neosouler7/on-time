
from collections import Counter, OrderedDict

import openpyxl

from common import STORE_MAP, PASS, SUCCESS, JINNY_THINK, JOY_CONFIRM, NO_6M_SALES
from common import get_current_time, get_column_number, get_last_data_idx, convert_to_zero


SOURCE_DIR = "C:/Users/jinny.hur/Desktop/업무파일/01. 리얼로케이션" # 마스터 파일 저장 경로
SOURCE_FILE_NAME = "Retail BTQ reallocation (0207_rank)_JH" # 마스터 파일명(확장자 제외)


def get_retail_info_template(store_info_list):
    store_keys_list = list(STORE_MAP.keys())

    store_info = {
        store_keys_list[i]: {
            "L12M_sales": convert_to_zero(store_info_list[i * 10]),
            "L6M_sales": convert_to_zero(store_info_list[i * 10 + 1]),
            "L3M_sales": convert_to_zero(store_info_list[i * 10 + 2]),
            "MS": convert_to_zero(store_info_list[i * 10 + 3]),
            "stock": convert_to_zero(store_info_list[i * 10 + 4]),
            "in_transit": convert_to_zero(store_info_list[i * 10 + 5]),
            "wish_list": convert_to_zero(store_info_list[i * 10 + 6]),
            "reallocation": convert_to_zero(store_info_list[i * 10 + 7]),
            "coverage": convert_to_zero(store_info_list[i * 10 + 8]),
            "rotation": store_info_list[i * 10 + 9], # rotation의 경우 수식값으로 None과 0.0 이 명확하게 구분되는 바 filter 적용하지 않음
        } for i in range(len(store_keys_list))
    }

    ranking = {
        "L12M_label": "",
        "L12M_sales_sum": 0,
        "L6M_label": "",
        "L6M_sales_sum": 0,
        "L3M_label": "",
        "L3M_sales_sum": 0
    }

    return {"model": "", "entry": "", "function": "", "shortage_store_count": 0, "store_info": store_info, "ranking": ranking}

def get_product_detail(ref_no):
    if len(ref_no) >= 8 and ref_no[3] == "4":
        model = ref_no[:8] + "00"
    elif len(ref_no) >= 8 and ref_no[3] == "6":
        model = ref_no[:8] + "00"
    else:
        model = ref_no

    if len(ref_no) >= 3 and ref_no[2] == "8":
        entry = "BIJOUX"
    elif len(ref_no) >= 3 and ref_no[2] == "B":
        entry = "BIJOUX"
    elif len(ref_no) >= 3 and ref_no[2] == "N":
        entry = "NJ"
    else:
        entry = "X"

    if len(ref_no) >= 4 and ref_no[3] == "4":
        function = "RING"
    elif len(ref_no) >= 4 and ref_no[3] == "6":
        function = "BRAC"
    elif len(ref_no) >= 4 and ref_no[3] == "7":
        function = "NECK"
    elif len(ref_no) >= 4 and ref_no[3] == "8":
        function = "EAR"
    elif len(ref_no) >= 3 and ref_no[2] == "8":
        function = "EAR"
    elif len(ref_no) >= 4 and ref_no[3] == "3":
        function = "NECK"
    else:
        function = "X"

    return model, entry, function

def get_sales_sum(store_info):
    L12M_sales_sum, L6M_sales_sum, L3M_sales_sum = 0.0, 0.0, 0.0
    for store_id, info in store_info.items():
        L12M_sales_sum += info["L12M_sales"]
        L6M_sales_sum += info["L6M_sales"]
        L3M_sales_sum += info["L3M_sales"]
    return L12M_sales_sum, L6M_sales_sum, L3M_sales_sum


class Main:
    def __init__(self):
        pass
    
    def run(self):
        """
            reallocation.py 내 포함하려고 하였으나, JINNY_THINK / JOY_CONFIRM 과 같은 작업 이후에 rotation 값이 변경 되므로, 부득이하게 별도 파일 생성하게 되었음.
            즉, 1. reallocation.py 실행 / 2. JINNY 엑셀 작업 / 3. followup.py 와 같은 순서로 진행하면 됨.

            followup.py 는 크게 아래의 기능을 가짐
            1. shortage_store_count
                a. 품번 기준, rotation 값이 0인 store 개수를 저장
            2. ranking
                a. dddd
                b. ddd
                
            3. 엑셀 반환

            # data_set
            {
                "CRN7413800": {
                    "model": a,
                    "entry": a,
                    "function": a,
                    "shortage_store_count": a,
                    "store_info": {
                        "6400: {
                            "L12M_sales": a,
                            "L6M_sales": a,
                            "L3M sales": b,
                            "MS": c,
                            "stock": d,
                            "in_transit": e,
                            "wish_list": f,
                            "reallocation": g,
                            "coverage": h,
                            "rotation": i,
                        },
                        ...
                    }
                    "ranking: {
                        "L12M_label": b,
                        "L12M_sales_sum": a,
                        "L6M_label": b,
                        "L6M_sales_sum": a,
                        "L3M_label": b,
                        "L3M_sales_sum": a,
                    }
                },
                ...
            }
        """

        # 마스터 파일의 "Allocation" 시트를 참조하여, 위 형태의 data-set을 구성
        file_name = f'{SOURCE_DIR}/{SOURCE_FILE_NAME}.xlsx'
        print(f"Reading {file_name} ...")
        wb = openpyxl.load_workbook(filename=file_name, data_only=True)
        ws = wb["Allocation"]

        retail_info = dict()
        for row in ws.iter_rows(min_row=8, max_row=get_last_data_idx(ws), values_only=True):
            temp = get_retail_info_template(list(row)[get_column_number("AK"):get_column_number("EP") + 1]) # store_info
            retail_info[row[1]] = temp

        # retail_info에 필요한 정보 저장
        print("----- DATA GATHERING START -----\n")
        for ref_no, retail in retail_info.items():
            # 품번 기준, rotation 값이 0인 store 개수를 저장한다.
            shortage_store_count = 0
            for item in retail.get("store_info").items():
                if item[1].get("rotation") == 0.0:
                    shortage_store_count += 1
            retail["shortage_store_count"] = shortage_store_count

            # TODO.
            # model, entry, function = get_product_detail(ref_no)

            # L12M_sales_sum, L6M_sales_sum, L3M_sales_sum = get_sales_sum(retail.get("store_info"))

            # retail_info[ref_no]["model"] = model
            # retail_info[ref_no]["entry"] = entry
            # retail_info[ref_no]["function"] = function
            # retail_info[ref_no]["ranking"]["L12M_sales_sum"] = L12M_sales_sum
            # retail_info[ref_no]["ranking"]["L6M_sales_sum"] = L6M_sales_sum
            # retail_info[ref_no]["ranking"]["L3M_sales_sum"] = L3M_sales_sum

        print("----- DATA GATHERING END -----\n")
        
        print("----- STATISTICS START -----\n")
        print("----- STATISTICS END -----\n")


        # 엑셀 반환
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active

        column_names = ["ref_no", "shortage_store_count"]
        new_ws.append(column_names)

        for ref_no, retail in retail_info.items():
            new_ws.append([ref_no, retail.get("shortage_store_count")])
            
        new_wb.save(filename=f'{SOURCE_DIR}/{SOURCE_FILE_NAME}_followup_{get_current_time("%Y%m%d_%H%M%S")}.xlsx')
        print(f"\nExcel successfully created!")
        print("")

        wb.close()
        new_wb.close()

if __name__ == "__main__":
    m = Main()
    m.run()