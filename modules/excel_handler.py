import pandas as pd
import msoffcrypto
import io
import os
import re
from datetime import datetime, timedelta

EPOCH = datetime(1899, 12, 30)          # Excel 1900 시스템 기준
LEAP_BUG_CUTOFF = datetime(1900, 3, 1)  # 1900-03-01 이전엔 +1 오류 존재
DEFAULT_EXCEL_PATH = "/매출리포트-250810203219_1 - Sample.xlsx"

def _parse_date_prefix(s: str) -> datetime | None:
    """'YYYY-MM-DD...' 형태 앞부분만 파싱(시/분/초는 무시). 실패 시 None."""
    try:
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s.strip())
        if not m:
            return None
        yyyy, mm, dd = map(int, m.groups())
        return datetime(yyyy, mm, dd)
    except Exception:
        return None
    
def _excel_serial_from_datetime(dt: datetime) -> int:
    """엑셀 시리얼 복원 (1900 윤년 버그 보정 포함)"""
    serial = (dt - EPOCH).days
    # 1900-01-01 ~ 1900-02-28 구간은 Excel이 +1로 잘못 매핑 → 1 빼서 보정
    if dt < LEAP_BUG_CUTOFF and dt >= datetime(1900, 1, 1):
        serial -= 1
    return serial

class ExcelHandler:
    def __init__(self, excel_path, password=None):
        self.excel_path = excel_path
        self.password = password
        self.df = None

        # 파일 존재 확인
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_path}")

        print(f"📁 Excel 파일: {excel_path}")
        print(f"🔐 비밀번호: {'설정됨' if password else '없음'}")

    def convert_excel_serial_to_date(self, serial_number):
        """Excel 시리얼 번호를 날짜로 변환"""
        try:
            if not serial_number or serial_number == "":
                return ""
            
            # 문자열을 float로 변환 시도
            serial_float = float(serial_number)
            
            # Excel epoch: 1900-01-01 (하지만 Excel 버그로 1900년을 윤년으로 계산)
            excel_epoch = datetime(1900, 1, 1)
            
            # 소수점이 있으면 시간까지 포함된 날짜
            if '.' in str(serial_number):
                # 날짜 + 시간
                actual_date = excel_epoch + timedelta(days=serial_float - 2)
                return actual_date.strftime('%Y-%m-%d %H:%M:%S')
            else:
                # 날짜만
                actual_date = excel_epoch + timedelta(days=serial_float - 2)
                return actual_date.strftime('%Y-%m-%d')
                
        except (ValueError, TypeError):
            # 변환 실패 시 원본 반환
            return str(serial_number)
        
    def fix_numeric_column(self, value):
        """
        숫자 컬럼 복원:
        - 1899~1901년대의 '가짜 날짜'로 읽힌 값은 엑셀 시리얼(정수)로 복원 (음수 포함)
        - 이미 숫자/숫자문자면 그대로 숫자로 반환
        - 정상적인 최근 날짜(예: 2025)는 건드리지 않음
        - 공란/누락은 빈문자열 반환(기존 동작 유지)
        반환타입: int 또는 float (숫자형), 공란은 ""
        """
        # 1) 빈값
        if value is None or value == "":
            return ""

        # 2) 숫자형은 그대로
        if isinstance(value, (int, float)):
            return int(value) if float(value).is_integer() else float(value)

        # 3) 숫자문자 처리
        if isinstance(value, str):
            s = value.strip()
            s_no_comma = s.replace(",", "")
            try:
                num = float(s_no_comma)
                return int(num) if num.is_integer() else num
            except ValueError:
                pass  # 날짜 후보로 이어감

        # 4) datetime → 시리얼 복원
        if isinstance(value, datetime):
            # 1901년 이후의 정상 날짜는 건드리지 않음
            if value.year <= 1901:
                return _excel_serial_from_datetime(value)
            return value

        # 5) 'YYYY-MM-DD...' 문자열 → 시리얼 복원(가짜 날짜만)
        if isinstance(value, str):
            dt = _parse_date_prefix(value)
            if dt and dt.year <= 1901:
                return _excel_serial_from_datetime(dt)
            return value

        # 6) 기타 타입은 최소 파괴
        return str(value)

    def fix_data_types(self):
        """읽어온 데이터의 타입 수정"""
        if self.df is None:
            return
        
        print("🔧 데이터 타입 수정 중...")
        
        # 날짜 컬럼들 수정
        date_columns = ['주문기준일자', '주문시작시각']
        
        for col in date_columns:
            if col in self.df.columns:
                print(f"  📅 {col} 변환 중...")
                self.df[col] = self.df[col].apply(self.convert_excel_serial_to_date)
        
        # 숫자 컬럼들 수정
        numeric_columns = ['수량']
        
        for col in numeric_columns:
            if col in self.df.columns:
                print(f"  🔢 {col} 수정 중...")
                self.df[col] = self.df[col].apply(self.fix_numeric_column)
        
        print("✅ 데이터 타입 수정 완료")

    def read_excel(self):
        """비밀번호 보호된 Excel 파일 읽기"""
        try:
            print("🔓 Excel 파일 복호화 중...")

            with open(self.excel_path, 'rb') as f:
                office_file = msoffcrypto.OfficeFile(f)

                if office_file.is_encrypted():
                    print("🔓 Excel 파일 복호화 중...")
                    # 비밀번호 있는 경우
                    if self.password:
                        office_file.load_key(password=self.password)
                    else:
                        raise ValueError("암호화된 파일인데 비밀번호가 제공되지 않았습니다.")

                    decrypted = io.BytesIO()
                    office_file.decrypt(decrypted)
                    excel_data = decrypted
                else:
                    print("🔓 비밀번호 없음 → 바로 로드")
                    excel_data = self.excel_path  # 원본 경로 그대로 사용

                print("📊 Excel 데이터 로딩 중...")
                self.df = pd.read_excel(
                    excel_data,
                    sheet_name=None,
                    dtype=str,           # 모든 컬럼을 문자열로 읽기 (원본 보존)
                    keep_default_na=False
                )

                print(f"✅ 시트 목록: {list(self.df.keys())}")

                # 상품 주문 상세내역 시트 선택
                if "상품 주문 상세내역" in self.df:
                    self.df = self.df["상품 주문 상세내역"]
                    print(f"📋 데이터 로딩 성공: {len(self.df)}행")
                    
                    # 데이터 타입 수정 (Excel 시리얼 번호 → 날짜)
                    self.fix_data_types()
                    
                    # 각 컬럼별 첫 5개 행 데이터 확인
                    print("\n📊 컬럼별 샘플 데이터 (첫 5행):")
                    for col in self.df.columns:
                        print(f"\n🔸 {col}:")
                        sample_data = self.df[col].head(5)
                        for i, val in enumerate(sample_data):
                            if pd.notna(val):  # 값이 있는 경우만 출력
                                val_type = type(val).__name__
                                val_str = str(val)[:50]  # 길이 제한
                                print(f"  행{i+1}: {val_str} ({val_type})")
                            else:
                                print(f"  행{i+1}: [빈값] (NaN)")
                    
                    return True
                else:
                    print(f"❌ '상품 주문 상세내역' 시트를 찾을 수 없습니다")
                    return False

        except Exception as e:
            print(f"❌ Excel 읽기 실패: {e}")
            return False
        
    def show_excel_info(self):
      """Excel 파일의 현재 구조 확인"""
      if self.df is None:
          print("❌ 데이터가 로드되지 않았습니다. read_excel()을 먼저 실행하세요.")
          return

      print("\n=== Excel 구조 정보 ===")
      print(f"📊 총 행 수: {len(self.df)}")
      print(f"📊 총 열 수: {len(self.df.columns)}")

      print(f"\n📋 컬럼 목록:")
      for i, col in enumerate(self.df.columns):
          print(f"  {i+1:2d}. {col}")

      # 주요 컬럼 찾기
      key_cols = {}
      for i, col in enumerate(self.df.columns):
          if '주문시작시각' in str(col):
              key_cols['시간'] = col
          elif '옵션' in str(col):
              key_cols['옵션'] = col

      print(f"\n🔍 주요 컬럼 위치: {key_cols}")

      # 첫 3행 데이터 미리보기
      print(f"\n👀 데이터 미리보기:")
      print(self.df.head(3).to_string(max_cols=5))

    def _find_option_col(self) -> str | None:
        """헤더 중 '옵션'이 포함된 첫 번째 컬럼명을 반환. 없으면 None."""
        if self.df is None:
            return None
        for col in self.df.columns:
            if '옵션' in str(col):
                return col
        return None

    def filter_by_option_keywords(self, keywords: list[str], mode: str = "any"):
        """
        옵션 컬럼에서 keywords가 매칭되는 행만 필터링.
        - mode='any' : 키워드 중 하나라도 포함
        - mode='all' : 키워드 전부 포함
        반환: 필터링된 DataFrame
        """
        if self.df is None:
            raise RuntimeError("데이터가 로드되지 않았습니다. read_excel()을 먼저 실행하세요.")

        option_col = self._find_option_col()
        if option_col is None:
            raise KeyError("헤더에 '옵션'이 포함된 컬럼을 찾지 못했습니다.")

        s = self.df[option_col].fillna("").astype(str)

        if mode == "all":
            mask = True
            for kw in keywords:
                mask = mask & s.str.contains(kw, na=False)
        else:  # "any"
            mask = False
            for kw in keywords:
                mask = mask | s.str.contains(kw, na=False)

        return self.df[mask].copy()

    def add_delivery_columns(self):
      """배송 정보 컬럼 추가"""
      if self.df is None:
          print("❌ 데이터가 로드되지 않았습니다.")
          return False

      delivery_cols = [
          '수하인명', '수하인주소', '수하인전화번호', '수하인핸드폰번호',
          '박스수량', '택배운임', '운임구분', '품목명', '배송메세지'
      ]

      print(f"📦 배송 컬럼 {len(delivery_cols)}개 추가 중...")

      # 컬럼 추가 (빈 값으로)
      for col in delivery_cols:
          if col not in self.df.columns:
              self.df[col] = None
              print(f"  ✅ {col}")
          else:
              print(f"  ⚠️ {col} (이미 존재)")

      print(f"🎉 총 컬럼 수: {len(self.df.columns)}개")
      return True
    
    def save_excel(self, output_path=None):
      """Excel 파일 저장"""
      if self.df is None:
          print("❌ 저장할 데이터가 없습니다.")
          return False

      # 저장 경로 결정
      if output_path:
          save_path = output_path
      else:
          # 원본 파일명에 "_배송정보추가" 추가
          base_name = self.excel_path.rsplit('.', 1)[0]
          save_path = f"{base_name}_배송정보추가.xlsx"

      try:
          print(f"💾 Excel 저장 중: {save_path}")

          # pandas로 저장 (비밀번호는 제거됨)
          self.df.to_excel(save_path, sheet_name="상품 주문 상세내역",
  index=False, engine='openpyxl')

          print(f"✅ 저장 완료!")
          print(f"📁 원본: {self.excel_path}")
          print(f"📁 새파일: {save_path}")
          print("⚠️  새 파일에는 비밀번호가 없습니다.")

          return save_path

      except Exception as e:
          print(f"❌ 저장 실패: {e}")
          return False
      
    def compare_data_before_after(self):
        """저장 전후 데이터 비교"""
        if self.df is None:
            return

        print("🔍 데이터 변경 사항 확인:")

        # 첫 5행의 주요 컬럼 확인
        key_cols = ['주문시작시각', '상품가격', '실판매금액 \n (할인, 옵션 포함)']

        for col in key_cols:
            if col in self.df.columns:
                print(f"\n📊 {col}:")
                sample_data = self.df[col].head(3).tolist()
                for i, val in enumerate(sample_data):
                    print(f"  행{i+2}: {val} ({type(val).__name__})")



def test_structure():
    """구조 확인 테스트"""
    handler = ExcelHandler(DEFAULT_EXCEL_PATH, "1202")
    if handler.read_excel():
        handler.show_excel_info()
    else:
        print("❌ 파일 읽기 실패")

def test_read():
    """파일 읽기 테스트"""
    handler = ExcelHandler(DEFAULT_EXCEL_PATH, "1202")
    if handler.read_excel():
        print("✅ 파일 읽기 성공")
        print(f"📊 컬럼 수: {len(handler.df.columns)}")
    else:
        print("❌ 파일 읽기 실패")

def test_add_columns():
      """컬럼 추가 테스트"""
      handler = ExcelHandler(DEFAULT_EXCEL_PATH, "1202")
      if handler.read_excel():
          print("📊 추가 전:", len(handler.df.columns), "개")
          handler.compare_data_before_after()
          handler.add_delivery_columns()
          print("📊 추가 후:", len(handler.df.columns), "개")
          handler.compare_data_before_after()

          # 새 컬럼들 확인
          print("\n새 컬럼들:")
          for i, col in enumerate(handler.df.columns[-9:], len(handler.df.columns)-8):
              print(f"  {i:2d}. {col}")
      else:
          print("❌ 파일 읽기 실패")


def test_save():
      """저장 테스트"""
      handler = ExcelHandler(DEFAULT_EXCEL_PATH, "1202")
      if handler.read_excel():
          handler.add_delivery_columns()
          result = handler.save_excel()
          if result:
              print(f"🎉 저장 성공: {result}")
          else:
              print("❌ 저장 실패")
      else:
          print("❌ 파일 읽기 실패")

def test_full_process():
    """전체 과정 테스트 (읽기→컬럼추가→저장)"""
    handler = ExcelHandler(DEFAULT_EXCEL_PATH, "1202")

    if handler.read_excel():
        print("1️⃣ 원본 데이터 확인:")
        print(f"   주문기준일자: {handler.df['주문기준일자'].iloc[1]}")
        print(f"   주문시작시각: {handler.df['주문시작시각'].iloc[1]}")
        print(f"   수량: {handler.df['수량'].iloc[1]}")

        handler.add_delivery_columns()
        result = handler.save_excel()

        if result:
            print(f"2️⃣ 저장 성공: {result}")
            print("3️⃣ 원본 데이터가 보존된 상태로 배송 컬럼이 추가되었습니다!")

    else:
        print("❌ 실패")


if __name__ == "__main__":
    # test_read()
    # test_structure()
    # test_add_columns()
    # test_save()
    test_full_process()