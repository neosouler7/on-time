from datetime import datetime, timedelta
from collections import Counter, OrderedDict

import time
import openpyxl


SOURCE_DIR = "/Users/jh/pythonProjects/on-time" # 마스터 파일 저장 경로
SOURCE_FILE_NAME = "reallocation_mst4" # 마스터 파일명(확장자 제외)

STORE_MAP = {
    "6400": "Hyundai AKJD", # AK ~ AS
    "6401": "Galleria", # AT ~ BB
    "6402": "Hyundai COEX", # BC ~ BK
    "6403": "Lotte Avenuel", # BL ~ BT
    "6404": "SSG Kangnam", # BU ~ CC
    "6405": "SSG Time square", # CD ~ CL
    "6407": "SSG Centum", # CM ~ CU
    "6408": "Maison Cheongdam", # CV ~ DD
    "6409": "Lotte Avenuel WT", # DE ~ DM
    "6410": "SSG Daegu", # DN ~ DV
    "6411": "Hyundai Pangyo" # DW ~ EE
}

PASS, SUCCESS, JINNY_NEEDED = "pass", "success", "jinny_needed"

def get_current_time(date_format=None, day_delta=None):
    today = datetime.utcnow() + timedelta(hours=9)
    today += timedelta(days=day_delta) if day_delta is not None else timedelta()

    return today if date_format is None else today.strftime(date_format)

def get_column_number(column_id):
    return sum((ord(char) - ord('A') + 1) * 26 ** i for i, char in enumerate(reversed(column_id.upper()))) - 1

def get_last_data_idx(ws):
    start_row = end_row = 8
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        end_row += 1 if row[1] is not None else 0
        if row[1] is None:
            break
    return end_row - 1

def convert_to_zero(value):
    try:
        return float(value) if value not in (None, '-') else 0.0
    except ValueError:
        return 0.0

def get_retail_info_template(stock_info_list, store_info_list):
    store_keys_list = list(STORE_MAP.keys())

    stock_info = {
        "total_available": stock_info_list[0], 
        "warehouse": {
            "KRD4/1001": stock_info_list[4],
            "from_KRD4/1001": 0
        },
        "shipping": {
            f"shipment{i}": stock_info_list[(i-1) * 2 + 8] for i in range(1, 4)
        }
    }
    stock_info["shipping"].update({
        f"from_shipment{i}": 0 for i in range(1, 4)
    })

    store_info = {
        store_keys_list[i]: {
            "L6M_sales": convert_to_zero(store_info_list[i * 9]),
            "L3M_sales": convert_to_zero(store_info_list[i * 9 + 1]),
            "MS": convert_to_zero(store_info_list[i * 9 + 2]),
            "stock": convert_to_zero(store_info_list[i * 9 + 3]),
            "in_transit": convert_to_zero(store_info_list[i * 9 + 4]),
            "wish_list": convert_to_zero(store_info_list[i * 9 + 5]),
            "reallocation": convert_to_zero(store_info_list[i * 9 + 6]),
            "coverage": convert_to_zero(store_info_list[i * 9 + 7]),
            "rotation": convert_to_zero(store_info_list[i * 9 + 8]),
        } for i in range(len(store_keys_list))
    }

    return {"result": "", "stock_info": stock_info, "store_info": store_info}


class Main:
    def __init__(self):
        pass
    
    def run(self):
        """
            1. 데이터 조회
                : 아래 data-set 양식으로 마스터 파일 변환
            2. reallocation
                a. store_info('24.2.11)
                    1. total_available 값이 없으면 pass
                    2. rotation 값 기준으로 지점 별 내림차순하여, 최저값 지점에 1개씩 제공
                    3. 각 재고 배부 이후 다시 전체 rotation 조회하며, 2번 반복
                    4. 단, 3번 과정 중, 현재 재고로 최저 rotation 지점 배분 불가 시(최저 rotation 지점 개수 > 현재 재고) jinny_needed으로 분류

                b. stock_info('24.2.11)
                    1. TTL Available 값에 대하여 KRD4/1001, Shipment 1, Shipment 2, Shipment 3 순으로 배분
                    2. 단, success로 판정된 품번에 한하여 stock_info 조정(pass, jinny_needed은 수행하지 않음)

            3. 엑셀 반환
                a. ref_no 별 reallocation result - pass, success, jinny_needed
                b. updated rdc value - 17~18, 21~26
                c. updated store_info - only reallocation value of each store

            # data_set
            {
                "CRN7413800": {
                    "result": PASS / SUCCESS / JINNY_NEEDED
                    "stock_info": {
                        "total_available": a
                        "warehouse": {
                            "KRD4/1001": b
                        },
                        "shipping": {
                            "shipment1": d,
                            "shipment2": e,
                            "shipment3": f
                        }
                    },
                    "store_info": {
                        "6400: {
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
                },
                ...
            }
        """

        # 마스터 파일의 "Allocation" 시트를 참조하여, 위 형태의 data-set을 구성
        wb = openpyxl.load_workbook(filename=f'{SOURCE_DIR}/{SOURCE_FILE_NAME}.xlsx')
        ws = wb["Allocation"]

        retail_info = dict()
        for row in ws.iter_rows(min_row=8, max_row=get_last_data_idx(ws), values_only=True):
            temp = get_retail_info_template(list(row)[get_column_number("N"):get_column_number("AA") + 1] # stock_info
                                          , list(row)[get_column_number("AK"):get_column_number("EE") + 1]) # store_info
            retail_info[row[1]] = temp

        # reallocation for store_info
        print("----- STORE_INFO REALLOCATION -----\n")
        for ref_no, retail in retail_info.items():
            total_available = retail.get("stock_info").get("total_available")
            current_available = total_available # reallocation 위해 재고값 복사

            print(f'\n# START {ref_no} with {total_available} stock\n')

            if total_available in ['-', None]:
                print(f'{ref_no} pass since no TTL Available')
                retail_info[ref_no]["result"] = PASS
                continue

            while True:
                if current_available == 0:
                    print(f'### store reallocation finished\n')
                    time.sleep(0.3)
                    break

                sorted_dict = OrderedDict(sorted(retail["store_info"].items(), key=lambda x: x[1].get("rotation", 0) if isinstance(x[1], dict) else 0))
                rotation_counts = Counter(value.get("rotation", 0) if isinstance(value, dict) and isinstance(value.get("rotation"), (int, float)) else 0 for key, value in sorted_dict.items())
                rotation_counts = OrderedDict(sorted(rotation_counts.items()))

                lowest_rotation_count = rotation_counts[list(rotation_counts.keys())[0]]
                if lowest_rotation_count > current_available: # 현 재고에 대하여 더 이상 지점별로 동등하게 배분 못할 시
                    print(f'### {JINNY_NEEDED}\n- {rotation_counts} > {current_available}\n')
                    retail_info[ref_no]["result"] = JINNY_NEEDED
                    time.sleep(0.3)
                    break

                sorted_list = sorted(retail["store_info"].items(), key=lambda x: x[1].get("rotation", 0))
                store_id, store_info = sorted_list[0] # 가장 rotation이 낮은 store을 대상으로 함

                current_reallocation = store_info["reallocation"]
                current_rotation = store_info["rotation"]

                retail_info[ref_no]["store_info"][store_id]["reallocation"] += 1
                retail_info[ref_no]["store_info"][store_id]["rotation"] = round((store_info["stock"] + store_info["in_transit"] + current_reallocation) / store_info["L6M_sales"] / 6, 1)
                retail_info[ref_no]["result"] = SUCCESS

                current_available -= 1 # reallocation 가능 재고 차감

                msg = f'## {STORE_MAP.get(store_id)}({store_id})\n'
                msg += f'- reallocation: {current_reallocation} → {retail_info[ref_no]["store_info"][store_id]["reallocation"]}\n'
                msg += f'- rotation: {current_rotation} → {retail_info[ref_no]["store_info"][store_id]["rotation"]}'
                
                print(msg)
                print("")

        # reallocation for stock_info
        print("----- STOCK_INFO REALLOCATION -----\n")
        for ref_no, retail in retail_info.items():
            if retail.get("result") in [PASS, JINNY_NEEDED]: # 대상 재고가 없거나, JINNY_NEEED 의 경우 재고정보 변경 X 
                continue

            total_available = retail.get("stock_info").get("total_available")
            print(f'\n# START {ref_no} with {total_available} stock\n')

            while total_available > 0:
                KRD4_1001, from_KRD4_1001 = convert_to_zero(retail["stock_info"]["warehouse"]["KRD4/1001"]), convert_to_zero(retail["stock_info"]["warehouse"]["from_KRD4/1001"])
                if KRD4_1001 > from_KRD4_1001:
                    print(f'## stock added on warehouse-from_KRD4/1001')
                    retail["stock_info"]["warehouse"]["from_KRD4/1001"] = from_KRD4_1001 + 1
                    total_available -= 1

                shipment1, from_shipment1 = convert_to_zero(retail["stock_info"]["shipping"]["shipment1"]), convert_to_zero(retail["stock_info"]["shipping"]["from_shipment1"])
                if shipment1 > from_shipment1:
                    print(f'## stock added on shipping-from_shipment1')
                    retail["stock_info"]["shipping"]["from_shipment1"] = from_shipment1 + 1
                    total_available -= 1
                
                shipment2, from_shipment2 = convert_to_zero(retail["stock_info"]["shipping"]["shipment2"]), convert_to_zero(retail["stock_info"]["shipping"]["from_shipment2"])
                if shipment2 > from_shipment2:
                    print(f'## stock added on shipping-from_shipment2')
                    retail["stock_info"]["shipping"]["from_shipment2"] = from_shipment2 + 1
                    total_available -= 1
                
                shipment3, from_shipment3 = convert_to_zero(retail["stock_info"]["shipping"]["shipment3"]), convert_to_zero(retail["stock_info"]["shipping"]["from_shipment3"])
                if shipment3 > from_shipment3:
                    print(f'## stock added on shipping-from_shipment3')
                    retail["stock_info"]["shipping"]["from_shipment3"] = from_shipment3 + 1
                    total_available -= 1
                
            print(f'\n### stock reallocation finished\n')

        # 엑셀 반환
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active

        column_names = ["ref_no", "result"
                      , "rdc-from KRD4/1001", "rdc-from Shipment 1", "rdc-from Shipment 2", "rdc-from Shipment 3"
                      , "6400", "6401", "6402", "6403", "6404", "6405", "6407", "6408", "6409", "6410", "6411"]
        new_ws.append(column_names)

        success_cnt, pass_cnt, jinny_needed_cnt = 0, 0, 0
        for ref_no, retail in retail_info.items():
            result = retail.get("result")
            if result == PASS:
                pass_cnt += 1
            if result == SUCCESS:
                success_cnt += 1
            if result == JINNY_NEEDED:
                jinny_needed_cnt += 1

            stock_info, store_info = retail.get("stock_info"), retail.get("store_info")
            new_ws.append([ref_no, retail.get("result"), stock_info.get("warehouse").get("from_KRD4/1001")
                          , stock_info.get("shipping").get("from_shipment1"), stock_info.get("shipping").get("from_shipment2"), stock_info.get("shipping").get("from_shipment3")
                          , store_info.get("6400").get("reallocation"), store_info.get("6401").get("reallocation")
                          , store_info.get("6402").get("reallocation"), store_info.get("6403").get("reallocation")
                          , store_info.get("6404").get("reallocation"), store_info.get("6405").get("reallocation")
                          , store_info.get("6407").get("reallocation"), store_info.get("6408").get("reallocation")
                          , store_info.get("6409").get("reallocation"), store_info.get("6410").get("reallocation")
                          , store_info.get("6411").get("reallocation")])
            
        new_wb.save(filename=f'{SOURCE_DIR}/{SOURCE_FILE_NAME}_{get_current_time("%Y%m%d_%H%M%S")}.xlsx')
        print(f"\nExcel successfully created!")
        
        total_items = len(retail_info.items())
        print(f"\nTotal: {total_items} - PASS: {pass_cnt}({round(pass_cnt/total_items, 1)}%), SUCCESS: {success_cnt}({round(success_cnt/total_items, 1)})%, JINNY: {jinny_needed_cnt}({round(jinny_needed_cnt/total_items, 1)})%")

        wb.close()
        new_wb.close()

if __name__ == "__main__":
    m = Main()
    m.run()