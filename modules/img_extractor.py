from PIL import Image
import base64, io, os
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
client = OpenAI()

def encode_image_to_data_url(image_path: str) -> str:
    with Image.open(image_path) as img:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        return f"data:image/png;base64,{b64}"

INSTRUCTION = """
이미지 속 영수증을 읽고 아래 규칙으로 JSON만 출력하세요.
- 금액/수량은 쉼표·'원' 등 비숫자 제거 후 정수로.
- 주소는 줄바꿈 없이 한 줄로.
- 보이지 않는 값은 null.
- 표/구분선/헤더는 무시, 실데이터만.
- JSON 외 다른 텍스트 금지.
"""

# ✅ 스키마는 format 바로 아래에 name/strict가 위치해야 합니다
JSON_SCHEMA_FORMAT = {
    "type": "json_schema",
    "name": "receipt",     # ← 꼭 필요
    "strict": True,        # ← 엄격 모드
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "merchant_name": {"type": ["string","null"]},
            "total_amount":   {"type": ["integer","null"]},
            "payment_method": {"type": ["string","null"]},
            "balance":        {"type": ["integer","null"]},
            "card_number":    {"type": ["string","null"]},
            "installment":    {"type": ["string","null"]},
            "vat":            {"type": ["integer","null"]},
            "supply_amount":  {"type": ["integer","null"]},
            "approval_no":    {"type": ["string","null"]},
            "approved_at":    {"type": ["string","null"]},
            "merchant_no":    {"type": ["string","null"]},
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "name":       {"type": ["string","null"]},
                        "unit_price": {"type": ["integer","null"]},
                        "quantity":   {"type": ["integer","null"]},
                        "amount":     {"type": ["integer","null"]},
                        "options":    {"type": ["string","null"]}
                    },
                    "required": ["name","unit_price","quantity","amount", "options"]
                }
            },
            "representative":  {"type": ["string","null"]},
            "business_no":     {"type": ["string","null"]},
            "phone":           {"type": ["string","null"]},
            "address":         {"type": ["string","null"]}
        },
        "required": [
            "merchant_name",
            "total_amount",
            "payment_method",
            "balance",
            "card_number",
            "installment",
            "vat",
            "supply_amount",
            "approval_no",
            "approved_at",
            "merchant_no",
            "items",
            "representative",
            "business_no",
            "phone",
            "address"
        ]
    }
}

# 간단 폴백: 자유 JSON 객체
JSON_OBJECT_FORMAT = {"type": "json_object"}

def extract_receipt_json(image_path: str, model: str = "gpt-5-mini") -> str:
    data_url = encode_image_to_data_url(image_path)

    # 1차: 엄격 스키마
    try:
        resp = client.responses.create(
            model=model,
            input=[{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": INSTRUCTION.strip()},
                    {"type": "input_image", "image_url": data_url},
                ],
            }],
            max_output_tokens=4096,
            text={"format": JSON_SCHEMA_FORMAT},  # ✅ 올바른 위치/형식
        )
        out = getattr(resp, "output_text", "").strip()
        if out:
            return out
    except Exception as e:
        print(f"[WARN] json_schema 실패, json_object로 재시도: {e}")

    # 2차: 간단 JSON 객체 모드
    resp = client.responses.create(
        model=model,
        input=[{
            "role": "user",
            "content": [
                {"type": "input_text", "text": INSTRUCTION.strip()},
                {"type": "input_image", "image_url": data_url},
            ],
        }],
        max_output_tokens=4096,
        text={"format": JSON_OBJECT_FORMAT},
    )
    return getattr(resp, "output_text", "").strip()

if __name__ == "__main__":
    path = "/home/vaaast_lake/work_space/RA-Company/Screenshot 2025-08-11 095410.png"
    print(extract_receipt_json(path))
