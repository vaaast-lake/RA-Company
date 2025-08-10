import os
import json
import re
from openai import OpenAI
from typing import Dict, Optional, List
from dotenv import load_dotenv
from pydantic import BaseModel

class PersonalInfo(BaseModel):
    """개인정보 추출 결과 모델"""
    name: str
    phone: str
    address: str

class PersonalInfoExtractor:
    def __init__(self, api_key: str):
        """
        GPT-5-nano를 활용한 개인정보 추출기
        """
        self.client = OpenAI(api_key=api_key)
        self.model = "gpt-5-nano"  # GPT-5-nano 모델 사용
        
    def extract_info(self, text: str) -> Dict[str, Optional[str]]:
        """
        비정형 텍스트에서 이름, 전화번호, 주소 추출
        
        Args:
            text: 개인정보가 포함된 텍스트
            
        Returns:
            {'name': str, 'phone': str, 'address': str} 형태의 딕셔너리
        """
        
        try:
            # GPT-5 Responses API with Structured Outputs 사용
            response = self.client.responses.parse(
                model=self.model,
                instructions="한국어 개인정보를 정확히 추출하는 전문가입니다. 전화번호는 010-XXXX-XXXX 형태로 정규화하고, 주소는 전체를 하나의 문자열로 통합하세요.",
                input=f"다음 텍스트에서 이름, 전화번호, 주소를 추출하세요:\n\n{text}",
                text_format=PersonalInfo,
                reasoning={"effort": "minimal"}
            )

            print(f"전체 응답 내용: {response}")
            print(f"추출 성공: {response.output_parsed}")
            
            # 전화번호 정규화
            parsed_data = response.output_parsed
            normalized_phone = self._normalize_phone(parsed_data.phone)
            
            return {
                "name": parsed_data.name,
                "phone": normalized_phone,
                "address": parsed_data.address,
                "confidence": "high"
            }
            
        except Exception as e:
            print(f"API 호출 오류: {e}")
            return {"name": None, "phone": None, "address": None, "error": str(e)}
    
    
    def _normalize_phone(self, phone: str) -> str:
        """
        전화번호를 010-XXXX-XXXX 형태로 정규화
        """
        if not phone:
            return phone
            
        # 숫자만 추출
        digits = re.sub(r'[^\d]', '', phone)
        
        # 11자리 핸드폰 번호 형태로 변환
        if len(digits) == 11 and digits.startswith('010'):
            return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
        elif len(digits) == 10 and digits.startswith('01'):
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        else:
            return phone  # 원본 반환 (정규화 실패)

    def test_with_samples(self, samples: list) -> Dict:
        """
        여러 샘플로 추출 정확도 테스트
        
        Args:
            samples: 테스트할 개인정보 텍스트 리스트
            
        Returns:
            테스트 결과 통계
        """
        results = []
        
        for i, sample in enumerate(samples):
            print(f"\n=== 샘플 {i+1} 테스트 ===")
            print(f"입력: {sample}")
            
            result = self.extract_info(sample)
            results.append(result)
            
            print(f"추출 결과:")
            print(f"  이름: {result.get('name')}")
            print(f"  전화번호: {result.get('phone')}")
            print(f"  주소: {result.get('address')}")
            
            if result.get('error'):
                print(f"  오류: {result.get('error')}")
            elif result.get('confidence'):
                print(f"  신뢰도: {result.get('confidence')}")
        
        return {
            "total_samples": len(samples),
            "results": results,
            "success_rate": len([r for r in results if not r.get('error')]) / len(samples) * 100
        }


def main():
    """
    테스트 실행 함수
    """
    # .env 파일 로드
    load_dotenv()
    
    # API 키 설정 (환경변수 또는 직접 입력)
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        api_key = input("OpenAI API 키를 입력하세요: ")
    
    # 추출기 초기화
    extractor = PersonalInfoExtractor(api_key)
    
    # 가상의 샘플 데이터로 테스트
    test_samples = [
        "홍길동\n010-1234-5678\n서울시 강남구 테헤란로 123 ABC빌딩 1001호",
        "2.김철수\n3.01098765432\n4.부산시 해운대구 해변로 456 오션뷰아파트 201호",
        "박영희 01055667788 대구시 중구 중앙대로 789 센트럴타워 302호",
        "택배 발송을\n1. 영수증 사진\n2. 주문자 성함 : 이민수\n3. 연락처 : 010 - 4567 - 8901\n4. 수령하실 주소\n인천시 연수구 송도국제대로 321 스카이캐슬 501동 1205호",
        "광주시 서구 상무대로 654 르네상스타워 25층 2501호 정수진 010-2345-6789",
        "주문자: 최동욱\n주소: 대전시 유성구 대학로 987 유성타워빌 304동 1501호\n전화: 010-7890-1234",
        "송미영\n01099887766\n울산시 남구 삼산로 135 수정아파트 106동 801호 김지훈"
    ]
    
    print("GPT-5-nano API 개인정보 추출 테스트 시작")
    print("=" * 50)
    
    # 테스트 실행
    test_results = extractor.test_with_samples(test_samples)
    
    print(f"\n\n=== 테스트 결과 요약 ===")
    print(f"총 샘플 수: {test_results['total_samples']}")
    print(f"성공률: {test_results['success_rate']:.1f}%")
    
    # 상세 분석
    successful_extractions = 0
    for i, result in enumerate(test_results['results']):
        if not result.get('error') and result.get('name') and result.get('phone'):
            successful_extractions += 1
    
    print(f"완전 추출 성공: {successful_extractions}/{len(test_samples)} ({successful_extractions/len(test_samples)*100:.1f}%)")
    

if __name__ == "__main__":
    main()