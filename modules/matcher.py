"""
영수증-주문내역 매칭 시스템 스켈레톤 코드
"""

from datetime import datetime, timedelta
import pandas as pd
from typing import Dict, List, Optional, Tuple, Any
import re
import os

from .excel_handler_with_pyxl import excel_serial_to_datetime, excel_serial_to_str, ExcelHandlerPyXL

DEFAULT_EXCEL_PATH = "./매출리포트-250810203219_1 - Sample.xlsx"

class OrderMatcher:
    """영수증 정보와 주문 내역을 매칭하는 클래스"""
    
    def __init__(self, excel_handler):
        """
        매칭 시스템 초기화
        @param excel_handler: ExcelHandlerPyXL 인스턴스
        """
        self.excel_handler = excel_handler
        self.time_tolerance_seconds = 10  # 시간 허용 오차 (초)
    
    # ===== 메인 매칭 프로세스 =====
    
    def match_order(self, receipt_data: Dict, customer_info: Dict) -> Dict:
        """
        영수증과 고객 정보를 기반으로 주문 매칭 및 정보 입력
        @param receipt_data: 영수증에서 추출한 정보 (JSON)
        @param customer_info: 고객 정보 (name, phone, address)
        @returns: 매칭 결과 및 처리 상태
        """
        # 구현 예정:
        # 1. 입력 데이터 유효성 검증
        # 2. 엑셀에서 필터링된 주문 데이터 가져오기
        # 3. 매칭 조건별 순차 검색
        # 4. 매칭 결과에 따른 고객 정보 입력
        # 5. 처리 결과 반환 (성공/실패/다중매칭 등)
        try:
            # 1. 입력 데이터 유효성 검증
            if not self.validate_receipt_data(receipt_data):
                return {'status': 'error', 'message': '영수증 데이터가 유효하지 않습니다.'}
            
            if not self.validate_customer_info(customer_info):
                return {'status': 'error', 'message': '고객 정보가 유효하지 않습니다.'}
            
            # 2. 엑셀에서 필터링된 주문 데이터 가져오기
            if not self.excel_handler.worksheet:
                return {'status': 'error', 'message': '엑셀 워크시트가 로드되지 않았습니다.'}
            
            # 택배/채널 관련 주문들만 필터링
            order_df = self.excel_handler._sheet_to_dataframe_raw(self.excel_handler.worksheet)
            option_col = self.excel_handler._find_option_colname(order_df)
            
            if not option_col:
                return {'status': 'error', 'message': '옵션 컬럼을 찾을 수 없습니다.'}
            
            # 택배요청 또는 채널추가무료배송 포함 행만 필터링
            mask = order_df[option_col].fillna("").astype(str).str.contains("택배요청|채널추가무료배송", na=False)
            filtered_df = order_df[mask].copy()
            
            if filtered_df.empty:
                return {'status': 'no_orders', 'message': '택배 배송 대상 주문이 없습니다.'}
            
            # 3. 매칭 조건별 순차 검색
            matching_orders, debug_info = self.find_matching_orders(receipt_data, filtered_df)
            
            # 4. 매칭 결과에 따른 분기 처리
            if not matching_orders:
                return {
                    'status': 'no_match',
                    'message': '매칭되는 주문을 찾지 못했습니다.',
                    'debug_info': debug_info,
                    'receipt_info': {
                        'datetime': receipt_data.get('approved_at'),
                        'product': receipt_data.get('items', [{}])[0].get('name', '') if receipt_data.get('items') else ''
                    }
                }
            
            elif len(matching_orders) == 1:
                # 단일 매칭: 정보 입력
                matched_idx = matching_orders[0]['index']
                updated_count = self.update_customer_info([matched_idx], customer_info, receipt_data)  # ← receipt_data 추가
                
                return {
                    'status': 'success',
                    'message': '주문 매칭 및 고객 정보 입력 완료',
                    'matched_order': matching_orders[0],
                    'updated_order_blocks': updated_count,
                    'debug_info': debug_info
                }
            
            else:
                # 다중 매칭: 가장 높은 점수(첫 번째) 주문에 자동 입력
                best_match = matching_orders[0]
                matched_idx = best_match['index']
                updated_count = self.update_customer_info([matched_idx], customer_info, receipt_data)  # ← receipt_data 추가
                
                return {
                    'status': 'success',
                    'message': f'다중 매칭({len(matching_orders)}개) 중 최적 매칭에 고객 정보 입력 완료',
                    'matched_order': best_match,
                    'updated_order_blocks': updated_count,
                    'multiple_candidates': len(matching_orders),
                    'debug_info': debug_info
                }
        
        except Exception as e:
            return {
                'status': 'error',
                'message': f'처리 중 오류 발생: {str(e)}'
            }
        
    
    def find_matching_orders(self, receipt_data: Dict, order_df: pd.DataFrame) -> Tuple[List[Dict], Dict]:
        """
        영수증 정보로 매칭되는 주문들 검색
        @param debug: True시 디버그 정보도 함께 반환
        @returns: (매칭된 주문들, 디버그 정보)
        """
        matching_orders = []
        debug_info = {
            'total_rows': len(order_df),
            'checked_rows': 0,
            'date_pass': 0,
            'time_pass': 0,
            'product_pass': 0,
            'failed_reasons': [],
            'all_attempts': []
        }

        # 영수증에서 필요한 정보 추출
        receipt_datetime = receipt_data.get('approved_at', '')
        receipt_items = receipt_data.get('items', [])
        
        if not receipt_datetime or not receipt_items:
            debug_info['failed_reasons'].append("영수증 정보 부족: 시간 또는 상품 정보 없음")
            return matching_orders, debug_info
        
        receipt_product_name = receipt_items[0].get('name', '') if receipt_items else ''
        debug_info['receipt_info'] = {
            'datetime': receipt_datetime,
            'product': receipt_product_name
        }
        
        # DataFrame의 각 행을 순회하며 매칭 검사
        for idx, row in order_df.iterrows():
            debug_info['checked_rows'] += 1
            
            order_date_serial = row['주문기준일자']
            order_time_serial = row['주문시작시각']
            order_product_name = row['상품명']
            
            attempt_info = {
                'index': idx,
                'order_product': order_product_name,
                'date_match': False,
                'time_match': False,
                'product_match': False,
                'skip_reason': None
            }
            
            # NaN 값 처리
            if pd.isna(order_date_serial) or pd.isna(order_time_serial) or pd.isna(order_product_name):
                attempt_info['skip_reason'] = "NaN 값 존재"
                debug_info['all_attempts'].append(attempt_info)
                continue
            
            # 1. 날짜 매칭으로 1차 필터링
            date_match = self.match_date(receipt_datetime, order_date_serial)
            attempt_info['date_match'] = date_match
            
            if not date_match:
                attempt_info['skip_reason'] = "날짜 불일치"
                debug_info['all_attempts'].append(attempt_info)
                continue
            
            debug_info['date_pass'] += 1
            
            # 2. 시간 매칭으로 2차 필터링
            time_match = self.match_time(receipt_datetime, order_time_serial)
            attempt_info['time_match'] = time_match
            
            if not time_match:
                attempt_info['skip_reason'] = "시간 불일치"
                debug_info['all_attempts'].append(attempt_info)
                continue
                
            debug_info['time_pass'] += 1
            
            # 3. 상품명 매칭으로 3차 필터링
            product_match, product_similarity = self.match_product_name(receipt_product_name, order_product_name)
            attempt_info['product_match'] = product_match
            attempt_info['product_similarity'] = product_similarity
            
            if product_match:
                debug_info['product_pass'] += 1
                
                # 매칭 점수 계산
                total_score = 0.3 + 0.3 + (0.4 * product_similarity)
                
                matching_orders.append({
                    'index': idx,
                    'score': total_score,
                    'product_similarity': product_similarity,
                    'order_data': {
                        '주문기준일자': order_date_serial,
                        '주문시작시각': order_time_serial, 
                        '상품명': order_product_name,
                        '옵션': row.get('옵션', '')
                    }
                })
                
                attempt_info['matched'] = True
                attempt_info['score'] = total_score
            else:
                attempt_info['skip_reason'] = f"상품명 불일치 (유사도: {product_similarity:.3f})"
            
            debug_info['all_attempts'].append(attempt_info)

            # 3. 시간 매칭 통과 시 즉시 매칭 성공
            # debug_info['product_pass'] += 1

            # # 매칭 점수 계산 (날짜:0.5, 시간:0.5)
            # total_score = 0.5 + 0.5  # 날짜와 시간만으로 100% 점수

            # matching_orders.append({
            #     'index': idx,
            #     'score': total_score,
            #     'product_similarity': 1.0,  # 상품 매칭 안 함
            #     'order_data': {
            #         '주문기준일자': order_date_serial,
            #         '주문시작시각': order_time_serial, 
            #         '상품명': order_product_name,
            #         '옵션': row.get('옵션', '')
            #     }
            # })

            # attempt_info['matched'] = True
            # attempt_info['score'] = total_score
            # attempt_info['product_match'] = True  # 상품 매칭 건너뛰므로 True로 설정
            # attempt_info['product_similarity'] = 1.0

            # debug_info['all_attempts'].append(attempt_info)
        
        # 매칭 결과 정렬
        matching_orders.sort(key=lambda x: x['score'], reverse=True)
        
        return matching_orders, debug_info
    
    # ===== 개별 매칭 조건 검사 =====
    
    def match_date(self, receipt_datetime: str, order_serial) -> bool:
        """
        날짜 매칭 검사 (정확한 일치 필요)
        @param receipt_datetime: 영수증 날짜시간 ("2025-08-01 11:14:31")
        @param order_serial: 엑셀 시리얼 숫자 (45871.0)
        @returns: 날짜 일치 여부
        """
        try:
            # 데이터 타입과 값 확인
            print(f"[DEBUG] order_serial 타입: {type(order_serial)}")
            print(f"[DEBUG] order_serial 값: {order_serial}")
            print(f"[DEBUG] order_serial repr: {repr(order_serial)}")

            # 1. 영수증 datetime을 날짜만 추출
            receipt_dt = self.parse_receipt_datetime(receipt_datetime)
            receipt_date = receipt_dt.date()

            # # float로 변환 시도
            # if isinstance(order_serial, str):
            #     print(f"[DEBUG] 문자열을 float로 변환 시도: '{order_serial}'")
            #     order_serial = float(order_serial)
            
            # # 2. 엑셀 시리얼을 datetime으로 변환 후 날짜 추출
            # order_dt = excel_serial_to_datetime(order_serial)
            # order_date = order_dt.date()

            # Timestamp 객체 처리
            if hasattr(order_serial, 'date'):  # pandas Timestamp
                order_date = order_serial.date()
            else:  # float (시리얼 숫자)
                order_dt = excel_serial_to_datetime(float(order_serial))
                order_date = order_dt.date()

            print(f"[DEBUG] 영수증 날짜: {receipt_date}")
            print(f"[DEBUG] 주문 날짜: {order_date}")
            
            # 3. 날짜 정확 일치 확인
            return receipt_date == order_date
            
        except Exception as e:
            # 변환 실패 시 매칭 실패로 처리
            print(f"[DEBUG] 매칭 실패: {e}")
            return False
    
    def match_time(self, receipt_datetime: str, order_serial) -> bool:
        """
        시간 매칭 검사 (±10초 허용오차)
        @param receipt_datetime: 영수증 날짜시간
        @param order_serial: 엑셀 시리얼 숫자 (날짜+시간 포함)
        @returns: 시간 매칭 여부
        """
        try:
            # 1. 영수증 datetime을 datetime 객체로 변환
            receipt_dt = self.parse_receipt_datetime(receipt_datetime)
            
            # # 2. 엑셀 시리얼을 datetime 객체로 변환
            # order_dt = excel_serial_to_datetime(order_serial)
            
            # # 3. 시간 차이를 초 단위로 계산
            # time_diff = abs((receipt_dt - order_dt).total_seconds())

            # Timestamp 객체 처리
            if hasattr(order_serial, 'to_pydatetime'):  # pandas Timestamp
                order_dt = order_serial.to_pydatetime()
            else:  # float (시리얼 숫자)
                order_dt = excel_serial_to_datetime(float(order_serial))

            time_diff = abs((receipt_dt - order_dt).total_seconds())

            print(f"[TIME DEBUG] 영수증: {receipt_dt}, 주문: {order_dt}")
            print(f"[TIME DEBUG] 시간차: {time_diff}초, 허용: {self.time_tolerance_seconds}초, 통과: {time_diff <= self.time_tolerance_seconds}")
            
            # 4. ±10초 범위 내 확인
            return time_diff <= self.time_tolerance_seconds
            
        except Exception as e:
            # 변환 실패 시 매칭 실패로 처리
            print(f"[TIME ERROR] {e}")
            return False
    
    def match_product_name(self, receipt_name: str, order_name: str) -> Tuple[bool, float]:
        """
        상품명 매칭 검사 (편집 거리 기반)
        @param receipt_name: 영수증 상품명
        @param order_name: 주문 내역 상품명
        @returns: (매칭 여부, 유사도 점수)
        """
        if not receipt_name or not order_name:
            return False, 0.0
        
        import difflib
        
        # 대소문자 무시, 공백 제거
        receipt_clean = receipt_name.lower().replace(' ', '')
        order_clean = order_name.lower().replace(' ', '')
        
        # 편집 거리 기반 유사도 계산 (0.0 ~ 1.0)
        # 1.0 = 완전 일치, 0.0 = 완전 불일치
        similarity = difflib.SequenceMatcher(None, receipt_clean, order_clean).ratio()
        
        # 임계값 설정 (예: 0.8 이상이면 매칭으로 판단)
        # 한글자 차이는 대부분 통과 (8글자 중 1글자 = 87.5%)
        is_match = similarity >= 0.75
        
        return is_match, similarity
    
    # ===== 데이터 변환 및 전처리 =====
    
    def parse_receipt_datetime(self, datetime_str: str) -> datetime:
        """
        영수증 날짜시간 문자열을 datetime 객체로 변환
        @param datetime_str: "2025-08-01 11:14:31" 형식
        @returns: datetime 객체
        """
        # 1. 문자열 형식 검증
        # 2. datetime.strptime()으로 파싱 
        try:
            # 기본 형식: "YYYY-MM-DD HH:MM:SS"
            return datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            try:
                # 대안 형식: "YYYY-MM-DD"만 있는 경우
                return datetime.strptime(datetime_str, "%Y-%m-%d")
            except ValueError:
                # 3. 예외 처리 (잘못된 형식)
                raise ValueError(f"Invalid datetime format: {datetime_str}. Expected: 'YYYY-MM-DD HH:MM:SS'")
    
    def extract_product_keywords(self, product_name: str) -> List[str]:
        """
        상품명에서 핵심 키워드 추출
        @param product_name: 상품명
        @returns: 핵심 키워드 리스트
        """
        # 구현 예정:
        # 1. 괄호, 특수문자 제거
        # 2. 핵심 단어 추출 (예: "주이패턴", "이불", "냉감" 등)
        # 3. 불용어 제거
        # 4. 키워드 리스트 반환
        pass
    
    # ===== 고객 정보 입력 =====
    
    def update_customer_info(self, matched_indices: List[int], customer_info: Dict, receipt_data: Dict = None) -> int:
        """
        매칭된 주문에 고객 정보 입력 (워크시트 직접 수정)
        @param matched_indices: 매칭된 행 인덱스들  
        @param customer_info: 고객 정보
        @param receipt_data: 영수증 데이터 (품목명용)
        @returns: 업데이트된 행 수
        """
        if not self.excel_handler.worksheet:
            return 0
        
        # 1. DataFrame에서 그룹핑 정보만 추출
        temp_df = self.excel_handler._sheet_to_dataframe_raw(self.excel_handler.worksheet)
        time_groups = self.group_by_order_time(temp_df, matched_indices)
        
        updated_count = 0
        
        # 2. 워크시트에서 컬럼 인덱스 찾기
        header_row = 1  # 보통 첫 번째 행이 헤더
        col_mapping = {}
        all_headers = []  # 디버깅용
        for col_idx in range(1, self.excel_handler.worksheet.max_column + 1):
            cell_value = self.excel_handler.worksheet.cell(header_row, col_idx).value
            if cell_value in ['수하인명', '수하인전화번호', '수하인핸드폰번호', '수하인주소', '품목명']:
                col_mapping[cell_value] = col_idx

        # 디버깅 출력
        print(f"[DEBUG] 전체 헤더: {all_headers}")
        print(f"[DEBUG] 찾은 컬럼: {col_mapping}")
        
        # 3. 품목명 텍스트 생성 (영수증 데이터가 있을 때만)
        items_text = ""
        if receipt_data and receipt_data.get('items'):
            items_text = self.format_items_for_description(receipt_data['items'])
            print(f"[DEBUG] 생성된 품목명 텍스트: '{items_text}'")
        else:
            print(f"[DEBUG] 영수증 데이터 없음: receipt_data={receipt_data}")
        
        # 4. 각 그룹의 첫 번째 행에만 정보 입력
        for order_time, indices in time_groups.items():
            if not indices:
                continue
            
            # DataFrame 인덱스를 실제 엑셀 행 번호로 변환 (헤더 고려)
            first_excel_row = indices[0] + 2  # DataFrame 인덱스 + 헤더(1) + 0-based 보정(1)
            print(f"[DEBUG] 행 {first_excel_row}에 정보 입력 시도")

            # 고객 정보 입력
            if '수하인명' in col_mapping and customer_info.get('name'):
                self.excel_handler.worksheet.cell(first_excel_row, col_mapping['수하인명'], customer_info['name'])
                print(f"[DEBUG] 수하인명 입력: {customer_info['name']}")
            
            if '수하인전화번호' in col_mapping and customer_info.get('phone'):
                self.excel_handler.worksheet.cell(first_excel_row, col_mapping['수하인전화번호'], customer_info['phone'])
            
            if '수하인핸드폰번호' in col_mapping and customer_info.get('phone'):
                self.excel_handler.worksheet.cell(first_excel_row, col_mapping['수하인핸드폰번호'], customer_info['phone'])
            
            if '수하인주소' in col_mapping and customer_info.get('address'):
                self.excel_handler.worksheet.cell(first_excel_row, col_mapping['수하인주소'], customer_info['address'])
            
            # 품목명 입력 (새로 추가)
            if '품목명' in col_mapping and items_text:
                self.excel_handler.worksheet.cell(first_excel_row, col_mapping['품목명'], items_text)
            
            updated_count += 1
        
        return updated_count

    def format_items_for_description(self, items: List[Dict]) -> str:
        """
        영수증 items를 품목명 형식으로 변환
        @param items: 영수증 아이템 리스트
        @returns: 형식화된 품목명 텍스트
        """
        print(f"[DEBUG] format_items_for_description 호출: {items}")
        if not items:
            return ""
        
        # 택배 관련 옵션이 있는 아이템만 필터링
        delivery_items = []
        for item in items:
            options = item.get('options', '') or ''  # None 처리
            print(f"[DEBUG] 아이템 '{item.get('name')}' 옵션: '{options}'")
            
            # 옵션에 택배 관련 키워드가 있는지 확인
            if '채널추가무료배송' in options or '택배요청' in options:
                delivery_items.append(item)
                print(f"[DEBUG] 택배 아이템으로 선택: {item.get('name')}")
            else:
                print(f"[DEBUG] 택배 아이템 아님: {item.get('name')}")
        
        if not delivery_items:
            print("[DEBUG] 택배 관련 아이템이 없음")
            return ""

        # 총 수량 계산
        total_quantity = sum(item.get('quantity', 1) for item in delivery_items)
        print(f"[DEBUG] 총 수량: {total_quantity}")
        # 아이템별 텍스트 생성
        item_texts = []
        for item in delivery_items:
            name = item.get('name', '')
            quantity = item.get('quantity', 1)
            
            # 수량이 1개면 "1개" 생략 가능, 1개 초과면 명시
            if quantity == 1:
                item_texts.append(f"{name}")
            else:
                item_texts.append(f"{name} {quantity}개")
        
        print(f"[DEBUG] 생성된 아이템 텍스트: {item_texts}")

        # 최종 형식: "총X개) item1/item2/item3"
        if len(item_texts) == 1:
            # 단일 아이템인 경우
            result = f"총{total_quantity}개) {item_texts[0]}"
            print(f"[DEBUG] 단일 아이템 결과: '{result}'")
            return result
        else:
            # 복수 아이템인 경우
            result = f"총{total_quantity}개) {'/'.join(item_texts)}"
            print(f"[DEBUG] 복수 아이템 결과: '{result}'")
            return result
        
    def group_by_order_time(self, order_df: pd.DataFrame, indices: List[int]) -> Dict[float, List[int]]:
        """
        주문시작시각별로 행 인덱스 그룹핑
        @param order_df: 주문 내역 DataFrame
        @param indices: 매칭된 행 인덱스들
        @returns: {주문시작시각: [행인덱스들]} 딕셔너리
        """
        # 구현 예정:
        # 1. 각 인덱스의 주문시작시각 값 추출
        # 2. 동일한 시각별로 그룹핑
        # 3. 그룹별 딕셔너리 반환
        time_groups = {}
    
        for idx in indices:
            # 해당 행의 주문시작시각 값 추출
            order_time = order_df.loc[idx, '주문시작시각']
            
            # NaN 값 처리
            if pd.isna(order_time):
                continue
                
            # float 타입으로 변환 (시리얼 숫자)
            # order_time = float(order_time)
            # Timestamp 객체 처리
            if hasattr(order_time, 'timestamp'):  # pandas Timestamp
                # Timestamp를 float로 변환 (유닉스 타임스탬프 사용)
                time_key = order_time.timestamp()
            else:  # float (시리얼 숫자)
                time_key = float(order_time)
            
            # 해당 시각의 그룹에 인덱스 추가
            if order_time not in time_groups:
                time_groups[order_time] = []
            time_groups[order_time].append(idx)
        
        # 각 그룹 내에서 인덱스 정렬 (DataFrame 행 순서대로)
        for time_key in time_groups:
            time_groups[time_key].sort()
        
        return time_groups
    
    # ===== 유틸리티 및 검증 =====
    
    def validate_receipt_data(self, receipt_data: Dict) -> bool:
        """
        영수증 데이터 유효성 검증
        @param receipt_data: 영수증 정보
        @returns: 유효성 여부
        """
        if not receipt_data:
            return False
        
        # 1. 필수 필드 존재 확인 (approved_at, items 등)
        approved_at = receipt_data.get('approved_at')
        if not approved_at or not isinstance(approved_at, str):
            return False
        
        # 2. 데이터 형식 검증
        # 날짜 형식 검증
        try:
            self.parse_receipt_datetime(approved_at)
        except ValueError:
            return False
        
        # 필수 필드: items 존재 및 내용 확인
        items = receipt_data.get('items')
        if not items or not isinstance(items, list) or len(items) == 0:
            return False
        
        # 첫 번째 아이템에 name 필드 존재 확인
        first_item = items[0]
        if not isinstance(first_item, dict) or not first_item.get('name'):
            return False
        
        # 3. 유효성 검사 결과 반환
        return True
    

    def validate_customer_info(self, customer_info: Dict) -> bool:
        """
        고객 정보 유효성 검증
        @param customer_info: 고객 정보
        @returns: 유효성 여부
        """
        if not customer_info:
            return False
        
        # 1. 필수 필드 존재 확인 (name, phone, address)
        required_fields = ['name', 'phone', 'address']
        for field in required_fields:
            value = customer_info.get(field)
            if not value or not isinstance(value, str) or not value.strip():
                return False
        
        # 2. 전화번호 형식 검증 (010-XXXX-XXXX 또는 숫자만)
        phone = customer_info['phone'].strip()
        # 숫자와 하이픈만 허용
        if not re.match(r'^[\d\-]+$', phone):
            return False
        
        # 최소 길이 확인 (010-1234-5678 = 13자, 01012345678 = 11자)
        digits_only = re.sub(r'[^\d]', '', phone)
        if len(digits_only) < 10 or len(digits_only) > 11:
            return False
        
        # 3. 유효성 검사 결과 반환
        return True
    

    def calculate_match_score(self, date_match: bool, time_match: bool, 
                            product_similarity: float) -> float:
        """
        매칭 점수 계산
        @param date_match: 날짜 일치 여부
        @param time_match: 시간 일치 여부  
        @param product_similarity: 상품명 유사도
        @returns: 종합 매칭 점수 (0.0 ~ 1.0)
        """
        # 구현 예정:
        # 1. 각 조건별 가중치 적용
        # 2. 종합 점수 계산
        # 3. 0.0 ~ 1.0 범위로 정규화
        pass


# ===== 통합 실행 함수 =====

def process_receipt_and_customer(excel_path: str, password: str, receipt_data: Dict, 
                               customer_info: Dict) -> Dict:
    """
    영수증과 고객 정보를 처리하여 엑셀 업데이트
    @param excel_path: 엑셀 파일 경로
    @param password: 엑셀 암호 (없으면 None)
    @param receipt_data: 영수증 정보
    @param customer_info: 고객 정보
    @returns: 처리 결과
    """
    # 구현 예정:
    # 1. ExcelHandlerPyXL 인스턴스 생성 및 파일 로드
    # 2. OrderMatcher 인스턴스 생성
    # 3. 매칭 및 정보 업데이트 실행
    # 4. 엑셀 파일 저장
    # 5. 처리 결과 반환
    excel_handler = None
    
    try:
        # 1. ExcelHandlerPyXL 인스턴스 생성 및 파일 로드
        excel_handler = ExcelHandlerPyXL(excel_path, password)
        
        if not excel_handler.read_excel_basic():
            return {
                'status': 'error',
                'message': '엑셀 파일을 읽을 수 없습니다.',
                'file_path': excel_path
            }
        
        # 2. 필터링된 새 시트 생성
        keywords = ["채널추가무료배송", "택배요청"]
        new_sheet_name = "필터링_결과"
        
        excel_handler.filter_to_new_sheet_raw(
            keywords=keywords,
            new_sheet_name=new_sheet_name,
            mode="any",
            extra_cols={"배송처리상태": "대기", "메모": ""},
        )
        
        # 3. 워킹 시트를 새 시트로 변경
        if not excel_handler.switch_to_sheet(new_sheet_name):
            return {
                'status': 'error',
                'message': f'새 시트 "{new_sheet_name}"로 변경할 수 없습니다.'
            }
        
        # 4. OrderMatcher 인스턴스 생성 (새 시트 기준)
        matcher = OrderMatcher(excel_handler)
        
        # 5. 매칭 및 정보 업데이트 실행 (새 시트에서)
        match_result = matcher.match_order(receipt_data, customer_info)
        
        # 6. 성공한 경우에만 엑셀 파일 저장
        if match_result.get('status') == 'success':
            try:
                convert_date_columns_for_display(excel_handler.worksheet)
                excel_handler.workbook.save(excel_path)
                match_result['saved'] = True
                match_result['file_path'] = excel_path
                match_result['target_sheet'] = new_sheet_name
            except Exception as save_error:
                match_result['status'] = 'partial_success'
                match_result['message'] += f' (저장 실패: {str(save_error)})'
                match_result['saved'] = False
        
        return match_result
        
    except Exception as e:
        return {
            'status': 'error',
            'message': f'처리 중 예상치 못한 오류가 발생했습니다: {str(e)}'
        }

def convert_date_columns_for_display(worksheet):
    """저장 전에 날짜 컬럼을 사용자 친화적 형식으로 변환"""
    # 헤더에서 날짜 컬럼 찾기
    date_cols = {}
    for col_idx in range(1, worksheet.max_column + 1):
        cell_value = worksheet.cell(1, col_idx).value
        if cell_value in ['주문기준일자', '주문시작시각']:
            date_cols[cell_value] = col_idx
    
    # 각 행의 날짜 값 변환
    for row_idx in range(2, worksheet.max_row + 1):
        for col_name, col_idx in date_cols.items():
            cell = worksheet.cell(row_idx, col_idx)
            if cell.value and isinstance(cell.value, (int, float)):
                try:
                    with_time = (col_name == "주문시작시각")
                    cell.value = excel_serial_to_str(float(cell.value), with_time=with_time)
                except:
                    pass  # 변환 실패 시 원본 유지

def process_single_receipt_with_handler(excel_handler, receipt_data: Dict, 
                                      customer_info: Dict, target_sheet_name: str) -> Dict:
    """
    단일 영수증 처리 (기존 ExcelHandler 재사용)
    @param excel_handler: 기존 ExcelHandlerPyXL 인스턴스
    @param target_sheet_name: 필터링 시트명
    """
    try:
        # 1. 필터링 시트로 변경 (이미 생성된 시트)
        if not excel_handler.switch_to_sheet(target_sheet_name):
            return {
                'status': 'error',
                'message': f'필터링 시트 "{target_sheet_name}"를 찾을 수 없습니다.'
            }
        
        # 2. OrderMatcher 인스턴스 생성
        matcher = OrderMatcher(excel_handler)
        
        # 3. 매칭 및 정보 업데이트 실행 (영수증 데이터도 전달)
        match_result = matcher.match_order(receipt_data, customer_info)
        
        # 4. 결과 반환 (저장은 배치 완료 후 한 번에)
        if match_result.get('status') == 'success':
            match_result['saved'] = False  # 아직 저장 안함
            match_result['target_sheet'] = target_sheet_name
        
        return match_result
        
    except Exception as e:
        return {
            'status': 'error',
            'message': f'처리 중 예상치 못한 오류가 발생했습니다: {str(e)}'
        }

def convert_date_columns_for_display(worksheet):
    """저장 전에 날짜 컬럼을 사용자 친화적 형식으로 변환"""
    from .excel_handler_with_pyxl import excel_serial_to_str
    
    # 헤더에서 날짜 컬럼 찾기
    date_cols = {}
    for col_idx in range(1, worksheet.max_column + 1):
        cell_value = worksheet.cell(1, col_idx).value
        if cell_value in ['주문기준일자', '주문시작시각']:
            date_cols[cell_value] = col_idx
    
    # 각 행의 날짜 값 변환
    for row_idx in range(2, worksheet.max_row + 1):
        for col_name, col_idx in date_cols.items():
            cell = worksheet.cell(row_idx, col_idx)
            if cell.value and isinstance(cell.value, (int, float)):
                try:
                    with_time = (col_name == "주문시작시각")
                    cell.value = excel_serial_to_str(float(cell.value), with_time=with_time)
                except:
                    pass  # 변환 실패 시 원본 유지

# ===== 테스트 함수 =====

def test_order_matching():
    """주문 매칭 시스템 테스트"""
    # 구현 예정:
    # 1. 샘플 영수증 데이터 준비
    # 2. 샘플 고객 정보 준비  
    # 3. 매칭 프로세스 실행
    # 4. 결과 검증 및 출력
    pass

def test_data_conversion():
    """데이터 변환 함수 테스트"""
    print("=" * 40)
    print("데이터 변환 함수 테스트")
    print("=" * 40)
    
    # 임시 매처 인스턴스 (excel_handler 없이)
    matcher = OrderMatcher(None)
    
    # 1. 영수증 datetime 파싱 테스트
    test_cases = [
        "2025-08-01 11:14:31",
        "2025-08-01",
        "invalid-format"  # 에러 케이스
    ]
    
    print("\n[1] 영수증 datetime 파싱 테스트:")
    for case in test_cases:
        try:
            result = matcher.parse_receipt_datetime(case)
            print(f"   ✅ '{case}' → {result}")
        except Exception as e:
            print(f"   ❌ '{case}' → Error: {e}")
    
    # 2. 엑셀 시리얼 변환 테스트
    serial_cases = [
        45871.0,  # 날짜만
        45871.74030092593,  # 날짜+시간
        45871.46631944444,  # 다른 시간 (11:14:31에 해당하는 시리얼)
        -1  # 에러 케이스
    ]
    
    print(f"\n[2] 엑셀 시리얼 변환 테스트:")
    for case in serial_cases:
        try:
            result = excel_serial_to_datetime(case)
            print(f"   ✅ {case} → {result}")
        except Exception as e:
            print(f"   ❌ {case} → Error: {e}")
    
    # 3. 실제 매칭 시나리오 테스트
    print(f"\n[3] 실제 데이터 매칭 테스트:")
    try:
        # 영수증: 2025-08-01 11:14:31
        receipt_dt = matcher.parse_receipt_datetime("2025-08-01 11:14:31")
        print(f"   영수증 시간: {receipt_dt}")
        
        # 11:14:31을 시리얼로 계산 (45871 + 11시14분31초)
        time_fraction = (11*3600 + 14*60 + 31) / 86400  # 초를 하루 기준 분수로
        expected_serial = 45871 + time_fraction
        print(f"   예상 시리얼: {expected_serial}")
        
        # 시리얼을 다시 datetime으로
        converted_dt = excel_serial_to_datetime(expected_serial)
        print(f"   변환된 시간: {converted_dt}")
        
        # 차이 계산
        diff = abs((receipt_dt - converted_dt).total_seconds())
        print(f"   시간 차이: {diff:.2f}초")
        
    except Exception as e:
        print(f"   ❌ 매칭 테스트 실패: {e}")

def test_match_date():
    matcher = OrderMatcher(None)
    result = matcher.match_date("2025-08-01 11:14:31", 45870.0)  # 2025-08-01에 해당하는 시리얼
    print(f"매칭 결과: {result}")
    print(excel_serial_to_datetime(45870.0))  # 2025-08-01인지 확인
    print(excel_serial_to_datetime(45871.0))  # 2025-08-02인지 확인

def test_match_time():
    matcher = OrderMatcher(None)
    result = matcher.match_time("2025-08-01 11:14:31", 45870.46841435185)  # 같은 날 11:14:31 예상
    print(excel_serial_to_datetime(45870.46841435185))
    print(f"시간 매칭 결과: {result}")

def test_order_matcher_basic():
    """OrderMatcher 기본 기능 테스트"""
    print("=" * 60)
    print("OrderMatcher 기본 기능 테스트")
    print("=" * 60)
    
    # 1. Excel Handler 생성
    handler = ExcelHandlerPyXL(DEFAULT_EXCEL_PATH, None)
    if not handler.read_excel_basic():
        print("Excel 로드 실패")
        return
    
    # 2. OrderMatcher 생성
    matcher = OrderMatcher(handler)
    
    # 3. 개별 함수 테스트
    print("\n[테스트 1] 데이터 변환 함수:")
    try:
        dt = matcher.parse_receipt_datetime("2025-08-01 11:14:31")
        print(f"  parse_receipt_datetime: {dt}")
        
        serial_dt = excel_serial_to_datetime(45870.46841435185)
        print(f"  excel_serial_to_datetime: {serial_dt}")
    except Exception as e:
        print(f"  오류: {e}")
    
    # 4. 매칭 함수 테스트
    print("\n[테스트 2] 매칭 함수:")
    date_match = matcher.match_date("2025-08-01 11:14:31", 45870.0)
    time_match = matcher.match_time("2025-08-01 11:14:31", 45870.46841435185)
    product_match = matcher.match_product_name("주이패턴이불", "뜨왈주이패턴베개커버")
    
    print(f"  날짜 매칭: {date_match}")
    print(f"  시간 매칭: {time_match}")
    print(f"  상품 매칭: {product_match}")


def test_full_matching():
    """전체 매칭 프로세스 테스트 - 새 시트 워크플로우"""
    print("=" * 60)
    print("전체 매칭 프로세스 테스트 (새 시트 워크플로우)")
    print("=" * 60)
    
    # 샘플 데이터
    receipt_data = {
        "approved_at": "2025-08-01 11:14:31",
        "items": [{"name": "주이패턴이불(냉감나일론)"}]
    }
    
    customer_info = {
        "name": "홍길동",
        "phone": "010-1234-5678", 
        "address": "서울시 강남구 테헤란로 123"
    }
    
    print(f"[입력] 영수증 정보:")
    print(f"  - 시간: {receipt_data['approved_at']}")
    print(f"  - 상품: {receipt_data['items'][0]['name']}")
    
    print(f"\n[입력] 고객 정보:")
    print(f"  - 이름: {customer_info['name']}")
    print(f"  - 전화: {customer_info['phone']}")
    print(f"  - 주소: {customer_info['address']}")
    
    print(f"\n[실행] 매칭 프로세스 시작...")
    
    result = process_receipt_and_customer(
        DEFAULT_EXCEL_PATH, None, receipt_data, customer_info
    )
    
    print(f"\n[결과] 매칭 결과:")
    print(f"  - 상태: {result.get('status')}")
    print(f"  - 메시지: {result.get('message')}")
    
    if result.get('status') == 'success':
        print(f"  - 저장됨: {result.get('saved')}")
        print(f"  - 대상 시트: {result.get('target_sheet')}")
        print(f"  - 업데이트된 주문 블록: {result.get('updated_order_blocks')}")
    
    return result


def debug_matching_data():
    """매칭 실패 원인 디버깅"""
    print("=" * 60)
    print("매칭 데이터 디버깅")
    print("=" * 60)
    
    # 1. Excel 데이터 로드
    handler = ExcelHandlerPyXL(DEFAULT_EXCEL_PATH, None)
    if not handler.read_excel_basic():
        print("Excel 로드 실패")
        return
    
    # 2. 필터링된 데이터 확인
    order_df = handler._sheet_to_dataframe_raw(handler.worksheet)
    option_col = handler._find_option_colname(order_df)
    
    mask = order_df[option_col].fillna("").astype(str).str.contains("택배요청|채널추가무료배송", na=False)
    filtered_df = order_df[mask].copy()
    
    print(f"[필터링] 총 {len(filtered_df)}개 행이 택배/채널 조건에 해당")
    
    # 3. 영수증 정보
    receipt_data = {
        "approved_at": "2025-08-01 11:14:31",
        "items": [{"name": "주이패턴이불(냉감나일론)"}]
    }
    
    print(f"\n[영수증] 시간: {receipt_data['approved_at']}")
    print(f"[영수증] 상품: {receipt_data['items'][0]['name']}")
    
    # 4. 실제 엑셀 데이터와 비교
    print(f"\n[엑셀 데이터] 상위 5개 행:")
    for idx, row in filtered_df.head(5).iterrows():
        order_date = row.get('주문기준일자', 'N/A')
        order_time = row.get('주문시작시각', 'N/A')
        product_name = row.get('상품명', 'N/A')
        
        # 시리얼을 날짜로 변환
        try:
            if pd.notna(order_date):
                date_str = excel_serial_to_str(float(order_date), with_time=False)
            else:
                date_str = "N/A"
        except:
            date_str = f"Raw: {order_date}"
            
        try:
            if pd.notna(order_time):
                time_str = excel_serial_to_str(float(order_time), with_time=True)
            else:
                time_str = "N/A"
        except:
            time_str = f"Raw: {order_time}"
        
        print(f"  행 {idx}: 날짜={date_str}, 시간={time_str}")
        print(f"         상품='{product_name}'")
        print()

def debug_specific_matching():
    """특정 행(19번째)의 매칭 과정 상세 추적"""
    print("=" * 60)
    print("18번째 행 매칭 과정 디버깅")
    print("=" * 60)
    
    # 1. Excel 데이터 로드 및 필터링
    handler = ExcelHandlerPyXL(DEFAULT_EXCEL_PATH, None)
    handler.read_excel_basic()
    
    order_df = handler._sheet_to_dataframe_raw(handler.worksheet)
    option_col = handler._find_option_colname(order_df)
    mask = order_df[option_col].fillna("").astype(str).str.contains("택배요청|채널추가무료배송", na=False)
    filtered_df = order_df[mask].copy()
    
    # 2. 18번째 행 데이터 확인
    if len(filtered_df) < 18:
        print(f"[ERROR] 필터링된 데이터에 18번째 행이 없습니다. (총 {len(filtered_df)}행)")
        return
    
    target_row = filtered_df.iloc[17]  # 18번째 행 (0-based index)
    print(f"[18번째 행 데이터]")
    print(f"  주문기준일자: {target_row.get('주문기준일자')} (type: {type(target_row.get('주문기준일자'))})")
    print(f"  주문시작시각: {target_row.get('주문시작시각')} (type: {type(target_row.get('주문시작시각'))})")
    print(f"  상품명: '{target_row.get('상품명')}'")
    print(f"  옵션: '{target_row.get('옵션')}'")
    
    # 3. 시리얼을 날짜/시간으로 변환
    try:
        order_date_serial = float(target_row.get('주문기준일자'))
        order_time_serial = float(target_row.get('주문시작시각'))
        
        date_converted = excel_serial_to_str(order_date_serial, with_time=False)
        time_converted = excel_serial_to_str(order_time_serial, with_time=True)
        
        print(f"\n[변환된 값]")
        print(f"  날짜: {date_converted}")
        print(f"  시간: {time_converted}")
    except Exception as e:
        print(f"[ERROR] 시리얼 변환 실패: {e}")
        return
    
    # 4. 영수증 데이터와 매칭 테스트
    receipt_data = {
        "approved_at": "2025-08-01 11:14:31",
        "items": [{"name": "주이패턴이불(냉감나일론)"}]
    }
    
    matcher = OrderMatcher(handler)
    
    print(f"\n[매칭 테스트]")
    print(f"영수증 시간: {receipt_data['approved_at']}")
    print(f"영수증 상품: {receipt_data['items'][0]['name']}")
    
    # 개별 매칭 조건 테스트
    date_match = matcher.match_date(receipt_data['approved_at'], order_date_serial)
    time_match = matcher.match_time(receipt_data['approved_at'], order_time_serial)
    product_match, similarity = matcher.match_product_name(
        receipt_data['items'][0]['name'], 
        target_row.get('상품명')
    )
    
    print(f"\n[매칭 결과]")
    print(f"  날짜 매칭: {date_match}")
    print(f"  시간 매칭: {time_match}")
    print(f"  상품 매칭: {product_match} (유사도: {similarity:.3f})")
    
    if not date_match:
        # 날짜 차이 계산
        try:
            receipt_dt = matcher.parse_receipt_datetime(receipt_data['approved_at'])
            order_dt = excel_serial_to_datetime(order_date_serial)
            print(f"  날짜 차이: 영수증={receipt_dt.date()}, 주문={order_dt.date()}")
        except Exception as e:
            print(f"  날짜 비교 오류: {e}")
    
    if not time_match:
        # 시간 차이 계산
        try:
            receipt_dt = matcher.parse_receipt_datetime(receipt_data['approved_at'])
            order_dt = excel_serial_to_datetime(order_time_serial)
            time_diff = abs((receipt_dt - order_dt).total_seconds())
            print(f"  시간 차이: {time_diff:.1f}초 (허용: {matcher.time_tolerance_seconds}초)")
        except Exception as e:
            print(f"  시간 비교 오류: {e}")

def debug_new_sheet_matching():
    """새 시트에서의 매칭 과정 디버깅"""
    print("=" * 60)
    print("새 시트에서의 매칭 과정 디버깅")
    print("=" * 60)
    
    # 1. 전체 워크플로우 재현
    handler = ExcelHandlerPyXL(DEFAULT_EXCEL_PATH, None)
    handler.read_excel_basic()
    
    # 2. 필터링된 새 시트 생성
    keywords = ["채널추가무료배송", "택배요청"]
    new_sheet_name = "디버그_필터링"
    
    handler.filter_to_new_sheet_raw(
        keywords=keywords,
        new_sheet_name=new_sheet_name,
        mode="any",
        extra_cols={"배송처리상태": "대기", "메모": ""},
    )
    
    # 3. 새 시트로 변경
    if not handler.switch_to_sheet(new_sheet_name):
        print("시트 변경 실패")
        return
    
    # 4. 새 시트에서 데이터 확인
    new_df = handler._sheet_to_dataframe_raw(handler.worksheet)
    print(f"[새 시트] 총 {len(new_df)}개 행")
    print(f"[새 시트] 컬럼: {list(new_df.columns)}")
    
    # 5. 영수증 데이터로 실제 매칭 시도
    receipt_data = {
        "approved_at": "2025-08-01 11:14:31",
        "items": [{"name": "주이패턴이불(냉감나일론)"}]
    }
    
    matcher = OrderMatcher(handler)
    matching_orders = matcher.find_matching_orders(receipt_data, new_df)
    
    print(f"\n[매칭 결과] {len(matching_orders)}개 매칭")
    
    if matching_orders:
        for i, match in enumerate(matching_orders):
            print(f"  매칭 {i+1}: 인덱스={match['index']}, 점수={match['score']:.3f}")
    else:
        print("  매칭 실패 - 각 행별 상세 확인:")
        
        # 각 행별로 왜 실패했는지 확인
        for idx, row in new_df.head(20).iterrows():
            print(f"\n  행 {idx}:")
            try:
                date_serial = row.get('주문기준일자')
                time_serial = row.get('주문시작시각')
                product_name = row.get('상품명')
                
                if pd.isna(date_serial) or pd.isna(time_serial):
                    print(f"    → SKIP: NaN 값 존재 (날짜={date_serial}, 시간={time_serial})")
                    continue
                
                date_match = matcher.match_date(receipt_data['approved_at'], date_serial)
                time_match = matcher.match_time(receipt_data['approved_at'], time_serial)
                product_match, similarity = matcher.match_product_name(
                    receipt_data['items'][0]['name'], product_name
                )
                
                print(f"    날짜매칭={date_match}, 시간매칭={time_match}, 상품매칭={product_match}")
                print(f"    상품: '{product_name}' (유사도={similarity:.3f})")
                
            except Exception as e:
                print(f"    → ERROR: {e}")


def debug_customer_info_update():
    """고객 정보 입력 과정 디버깅"""
    print("=" * 60)
    print("고객 정보 입력 과정 디버깅")
    print("=" * 60)
    
    # 전체 프로세스 재현
    handler = ExcelHandlerPyXL(DEFAULT_EXCEL_PATH, None)
    handler.read_excel_basic()
    
    # 필터링 및 시트 변경
    handler.filter_to_new_sheet_raw(
        keywords=["채널추가무료배송", "택배요청"],
        new_sheet_name="디버그_고객정보",
        mode="any",
        extra_cols={"배송처리상태": "대기", "메모": ""},
    )
    handler.switch_to_sheet("디버그_고객정보")
    
    # 매칭 실행
    receipt_data = {
        "approved_at": "2025-08-01 11:14:31",
        "items": [{"name": "주이패턴이불(냉감나일론)"}]
    }
    
    customer_info = {
        "name": "홍길동",
        "phone": "010-1234-5678", 
        "address": "서울시 강남구 테헤란로 123"
    }
    
    matcher = OrderMatcher(handler)
    
    # 매칭된 인덱스 확인
    order_df = handler._sheet_to_dataframe_raw(handler.worksheet)
    matching_orders = matcher.find_matching_orders(receipt_data, order_df)
    
    if not matching_orders:
        print("매칭 실패")
        return
        
    print(f"[매칭] {len(matching_orders)}개 매칭")
    matched_indices = [match['index'] for match in matching_orders]
    print(f"[인덱스] {matched_indices}")
    
    # 헤더 확인
    print(f"\n[헤더 확인]")
    header_row = 1
    delivery_columns = ['수하인명', '수하인전화번호', '수하인핸드폰번호', '수하인주소']
    col_mapping = {}
    
    for col_idx in range(1, handler.worksheet.max_column + 1):
        cell_value = handler.worksheet.cell(header_row, col_idx).value
        if cell_value in delivery_columns:
            col_mapping[cell_value] = col_idx
            print(f"  {cell_value}: 컬럼 {col_idx}")
    
    if not col_mapping:
        print("  ❌ 배송 관련 컬럼을 찾을 수 없음")
        return
    
    # 그룹핑 확인
    time_groups = matcher.group_by_order_time(order_df, matched_indices)
    print(f"\n[그룹핑] {len(time_groups)}개 그룹")
    for order_time, indices in time_groups.items():
        print(f"  시각 {order_time}: 인덱스 {indices}")
        
        # 첫 번째 인덱스의 실제 엑셀 행 번호
        first_excel_row = indices[0] + 2
        print(f"  → 엑셀 행 번호: {first_excel_row}")
        
        # 실제 정보 입력 테스트
        print(f"  정보 입력 전 확인:")
        for col_name, col_idx in col_mapping.items():
            old_value = handler.worksheet.cell(first_excel_row, col_idx).value
            print(f"    {col_name} (컬럼{col_idx}): '{old_value}'")
        
        # 정보 입력
        if '수하인명' in col_mapping:
            handler.worksheet.cell(first_excel_row, col_mapping['수하인명'], customer_info['name'])
        if '수하인전화번호' in col_mapping:
            handler.worksheet.cell(first_excel_row, col_mapping['수하인전화번호'], customer_info['phone'])
        if '수하인핸드폰번호' in col_mapping:
            handler.worksheet.cell(first_excel_row, col_mapping['수하인핸드폰번호'], customer_info['phone'])
        if '수하인주소' in col_mapping:
            handler.worksheet.cell(first_excel_row, col_mapping['수하인주소'], customer_info['address'])
        
        print(f"  정보 입력 후 확인:")
        for col_name, col_idx in col_mapping.items():
            new_value = handler.worksheet.cell(first_excel_row, col_idx).value
            print(f"    {col_name} (컬럼{col_idx}): '{new_value}'")
    
    # 저장 시도
    try:
        handler.workbook.save("./debug_customer_info.xlsx")
        print(f"\n[저장] debug_customer_info.xlsx 파일로 저장 완료")
    except Exception as e:
        print(f"\n[저장 오류] {e}")

# if __name__ == "__main__":
    # test_order_matching()
    # test_data_conversion()
    # test_match_date()
    # test_match_time()
    # test_order_matcher_basic()
    # test_full_matching()
    # debug_matching_data()
    # debug_specific_matching()
    # debug_new_sheet_matching()
    # debug_customer_info_update()