import streamlit as st
import json
import os
import pandas as pd
from PIL import Image
import tempfile
import datetime
import io
import base64

# 기존 모듈들 import
from modules.excel_handler_with_pyxl import ExcelHandlerPyXL
from modules.img_extractor import extract_receipt_json
from modules.info_extractor import PersonalInfoExtractor
from modules.matcher import convert_date_columns_for_display, process_single_receipt_with_handler

# WSL에서는 streamlit run app.py --server.headless true 로 실행
# 그냥 실행하면 ELinks가 켜짐

# ============================================
# Utils
# ============================================

def resize_image(image, max_width=400, max_height=600):
    """이미지를 적당한 크기로 리사이즈"""
    # 원본 크기
    original_width, original_height = image.size
    
    # 비율 계산
    width_ratio = max_width / original_width
    height_ratio = max_height / original_height
    ratio = min(width_ratio, height_ratio, 1.0)  # 확대는 하지 않음
    
    # 새 크기 계산
    new_width = int(original_width * ratio)
    new_height = int(original_height * ratio)
    
    # 리사이즈
    if ratio < 1.0:
        resized_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        return resized_image
    else:
        return image  # 원본이 더 작으면 그대로 반환
    
def process_batch(selected_indices, excel_file, excel_password):
    """선택된 세트들을 배치로 처리하고 최종 결과 파일 저장"""
    
    if not selected_indices:
        st.error("처리할 세트를 선택해주세요.")
        return
    
    # 진행률 표시
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 결과 요약
    success_count = 0
    fail_count = 0
    results = []
    
    # 엑셀 임시 저장 (누적 처리 방식)
    if st.session_state.batch_result_file:
        # 이전 처리 결과가 있으면 그것을 기반으로 시작
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(st.session_state.batch_result_file)
            excel_path = tmp_excel.name
        st.info("🔄 이전 처리 결과에 추가로 처리합니다.")
    else:
        # 처음 처리하는 경우만 원본 파일 사용
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(excel_file.read())
            excel_path = tmp_excel.name
        st.info("🆕 새로운 파일로 처리를 시작합니다.")
    
    try:
        total_sets = len(selected_indices)
        excel_handler = None
        
        # ========== 배치 처리 시작 시 초기화 ==========
        # ExcelHandler 한 번만 생성
        excel_handler = ExcelHandlerPyXL(excel_path, excel_password)
        if not excel_handler.read_excel_basic():
            st.error("엑셀 파일을 읽을 수 없습니다.")
            return
        
        # 필터링된 시트 생성 (한 번만)
        keywords = ["채널추가무료배송", "택배요청"]
        new_sheet_name = "필터링_결과"
        
        excel_handler.filter_to_new_sheet_raw(
            keywords=keywords,
            new_sheet_name=new_sheet_name,
            mode="any",
            extra_cols={"배송처리상태": "대기", "메모": ""},
        )
        excel_handler.switch_to_sheet(new_sheet_name)
        st.info(f"📋 필터링 시트 '{new_sheet_name}' 생성 완료")
        # =============================================
        
        for i, idx in enumerate(selected_indices):
            receipt_set = st.session_state.receipt_sets[idx]
            
            # 진행률 업데이트
            progress = (i + 1) / total_sets
            progress_bar.progress(progress)
            status_text.text(f"처리 중: {receipt_set['name']} ({i+1}/{total_sets})")
            
            # 상태 업데이트
            st.session_state.receipt_sets[idx]['status'] = '처리중'
            
            # 개별 세트 처리 - 단계별 예외 처리
            image_path = None
            receipt_data = None
            customer_info = None
            
            try:
                # 1. 이미지 임시 저장
                try:
                    import base64
                    image_data = base64.b64decode(receipt_set['image_data'])
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        tmp_img.write(image_data)
                        image_path = tmp_img.name
                except Exception as e:
                    raise Exception(f"이미지 디코딩/저장 실패: {str(e)}")
                
                # 2. 영수증 정보 추출
                try:
                    receipt_json = extract_receipt_json(image_path)
                    receipt_data = json.loads(receipt_json)
                except Exception as e:
                    raise Exception(f"영수증 정보 추출 실패: {str(e)}")
                
                # 3. 고객 정보 추출
                try:
                    extractor = PersonalInfoExtractor(os.getenv('OPENAI_API_KEY'))
                    customer_info = extractor.extract_info(receipt_set['customer_info'])
                except Exception as e:
                    raise Exception(f"고객 정보 추출 실패: {str(e)}")
                
                # 4. 매칭 수행
                try:
                    match_result = process_single_receipt_with_handler(
                        excel_handler,  # 기존 핸들러 전달
                        receipt_data,
                        customer_info,
                        new_sheet_name
                    )
                except Exception as e:
                    raise Exception(f"매칭 처리 실패: {str(e)}")
                
                # 5. 성공/실패 결과 저장
                if match_result['status'] == 'success':
                    st.session_state.receipt_sets[idx]['status'] = '완료'
                    success_count += 1
                    
                    simplified_result = {
                        'status': 'success',
                        'message': match_result['message'],
                        'matched_product': match_result['matched_order']['order_data']['상품명'],
                        'match_score': f"{match_result['matched_order']['score']:.1%}",
                        'updated_blocks': match_result['updated_order_blocks'],
                        'customer_name': customer_info.get('name', 'N/A')
                    }
                else:
                    # 매칭 실패 (시스템 오류는 아님)
                    st.session_state.receipt_sets[idx]['status'] = '실패'
                    fail_count += 1
                    
                    simplified_result = {
                        'status': 'failed',
                        'message': match_result['message'],
                        'customer_name': customer_info.get('name', 'N/A') if customer_info else 'N/A',
                        'receipt_data': receipt_data,
                        'receipt_datetime': receipt_data.get('approved_at', 'N/A') if receipt_data else 'N/A',
                        'receipt_product': receipt_data.get('items', [{}])[0].get('name', 'N/A') if receipt_data and receipt_data.get('items') else 'N/A',
                        'debug_info': match_result.get('debug_info')  # ← debug 정보 추가
                    }
                
                st.session_state.receipt_sets[idx]['result'] = simplified_result
                results.append({
                    'set_name': receipt_set['name'],
                    'result': simplified_result
                })
                
            except Exception as e:
                # 개별 세트 처리 중 시스템 오류 발생
                st.session_state.receipt_sets[idx]['status'] = '실패'
                fail_count += 1
                
                # 고객명 추출 시도 (가능한 경우)
                customer_name = 'N/A'
                if customer_info and customer_info.get('name'):
                    customer_name = customer_info['name']
                elif receipt_set.get('customer_info'):
                    # 첫 번째 줄에서 이름 추출 시도
                    first_line = receipt_set['customer_info'].split('\n')[0].strip()
                    if first_line and len(first_line) < 10:  # 이름으로 보이는 경우
                        customer_name = first_line
                
                error_result = {
                    'status': 'error',
                    'message': f'처리 중 오류 발생: {str(e)}',
                    'customer_name': customer_name,
                    'error_detail': str(e)
                }
                st.session_state.receipt_sets[idx]['result'] = error_result
                results.append({
                    'set_name': receipt_set['name'],
                    'result': error_result
                })
                
                # 실패 표시를 상태 텍스트에 반영
                status_text.text(f"⚠️ {receipt_set['name']} 실패 - 계속 진행 중... ({i+1}/{total_sets})")
            
            finally:
                # 개별 세트 처리 후 정리
                if image_path and os.path.exists(image_path):
                    try:
                        os.unlink(image_path)
                    except:
                        pass
        
        # 배치 처리 완료 후 최종 저장
        if success_count > 0:
            # 저장 전 날짜 형식 변환
            convert_date_columns_for_display(excel_handler.worksheet)
            excel_handler.workbook.save(excel_path)
            
            with open(excel_path, 'rb') as f:
                st.session_state.batch_result_file = f.read()
            st.session_state.batch_processing_complete = True
        
        # 완료 후 결과 표시
        progress_bar.progress(1.0)
        status_text.text("배치 처리 완료!")
        
        # 결과 요약
        st.success(f"🎉 배치 처리 완료!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("✅ 성공", success_count)
        with col2:
            st.metric("❌ 실패", fail_count)
        with col3:
            st.metric("📊 총 처리", total_sets)
        
        # 다운로드 버튼 (성공한 케이스가 있을 때만)
        if success_count > 0 and st.session_state.batch_result_file:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"배치처리결과_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 전체 결과 파일 다운로드",
                data=st.session_state.batch_result_file,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        # 상세 결과
        if results:
            with st.expander("처리 결과 상세"):
                for result in results:
                    if result['result']['status'] == 'success':
                        st.success(f"**{result['set_name']}**: {result['result']['message']}")
                        st.write(f"- 고객: {result['result']['customer_name']}")
                        st.write(f"- 매칭 상품: {result['result']['matched_product']}")
                        st.write(f"- 매칭 점수: {result['result']['match_score']}")
                    else:
                        st.error(f"**{result['set_name']}**: {result['result']['message']}")
    
    finally:
        # 임시 엑셀 파일 정리
        try:
            os.unlink(excel_path)
        except:
            pass

# ==============================================
# Streamlit
# ==============================================

st.set_page_config(
    page_title="영수증 매칭 시스템",
    page_icon="🧾",
    layout="wide"
)

st.title("🧾 영수증-주문내역 매칭 시스템")

# 세션 상태 초기화
if 'receipt_sets' not in st.session_state:
    st.session_state.receipt_sets = []
if 'batch_result_file' not in st.session_state:
    st.session_state.batch_result_file = None
if 'batch_processing_complete' not in st.session_state:
    st.session_state.batch_processing_complete = False

# 탭 구성
tab1, tab2, tab3 = st.tabs(["📝 데이터 입력", "🔍 배치 처리", "📊 결과 관리"])

with tab1:
    st.header("영수증 세트 추가")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("영수증 이미지")
        uploaded_image = st.file_uploader("영수증 이미지", type=['png', 'jpg', 'jpeg'], key="new_image")
        if uploaded_image:
            image = Image.open(uploaded_image)
    
            # 원본 크기 표시
            original_size = image.size
            st.caption(f"원본 크기: {original_size[0]} x {original_size[1]}px")
            
            # 리사이즈된 이미지 표시
            display_image = resize_image(image, max_width=400, max_height=500)
            resized_size = display_image.size
            
            st.image(display_image, caption=f"미리보기 ({resized_size[0]} x {resized_size[1]}px)")
    
            if original_size != resized_size:
                st.caption("💡 표시용으로 크기가 조정되었습니다. 원본은 그대로 사용됩니다.")
    
    with col2:
        st.subheader("고객 정보")
        customer_text = st.text_area(
            "고객 정보 입력",
            placeholder="홍길동\n010-1234-5678\n서울시 강남구 테헤란로 123",
            height=200,
            key="new_customer"
        )
        
        # 세트 이름 (선택사항)
        set_name = st.text_input("세트 이름 (선택)", placeholder="고객명 또는 식별자")
    
    # 세트 추가 버튼
    if st.button("➕ 세트 추가", type="primary"):
        if uploaded_image and customer_text:
            # 이미지를 base64로 저장 (세션 상태 유지를 위해)
            import base64
            import io
            
            buf = io.BytesIO()
            image.save(buf, format='PNG')
            img_base64 = base64.b64encode(buf.getvalue()).decode()
            
            new_set = {
                'id': len(st.session_state.receipt_sets) + 1,
                'name': set_name if set_name else f"세트 {len(st.session_state.receipt_sets) + 1}",
                'image_data': img_base64,
                'image_name': uploaded_image.name,
                'customer_info': customer_text,
                'status': '대기',
                'result': None
            }
            
            st.session_state.receipt_sets.append(new_set)
            st.success(f"✅ '{new_set['name']}' 세트가 추가되었습니다!")
            st.rerun()
        else:
            st.error("영수증 이미지와 고객 정보를 모두 입력해주세요.")

with tab2:
    st.header("배치 처리")
    
    # 엑셀 파일 설정
    with st.sidebar:
        st.header("설정")
        excel_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])
        excel_password = st.text_input("엑셀 비밀번호 (선택)", type="password")
    
    if st.session_state.receipt_sets:
        # 현재 세트 목록 표시 및 편집
        st.subheader(f"등록된 세트 ({len(st.session_state.receipt_sets)}개)")
        
        # 개별 세트 카드 형태로 표시
        for i, receipt_set in enumerate(st.session_state.receipt_sets):
            with st.container():
                col1, col2, col3, col4 = st.columns([3, 2, 1, 1])
                
                with col1:
                    st.write(f"**{receipt_set['name']}**")
                    # 고객 정보 미리보기 (첫 줄만)
                    preview = receipt_set['customer_info'].split('\n')[0]
                    if len(preview) > 30:
                        preview = preview[:30] + "..."
                    st.caption(f"고객정보: {preview}")
                
                with col2:
                    st.write(f"이미지: {receipt_set['image_name']}")
                    st.write(f"상태: {receipt_set['status']}")
                
                with col3:
                    # 미리보기 버튼 - 처리 중이 아닐 때만 활성화
                    processing_in_progress = any(s['status'] == '처리중' for s in st.session_state.receipt_sets)
                    
                    if st.button(
                        "👁️", 
                        key=f"preview_{i}", 
                        help="미리보기",
                        disabled=processing_in_progress  # ← 처리 중일 때 비활성화
                    ):
                        st.session_state[f'show_preview_{i}'] = True
                
                with col4:
                    # 삭제 버튼
                    if st.button("🗑️", key=f"delete_{i}", help="삭제"):
                        st.session_state.receipt_sets.pop(i)
                        st.rerun()
            
            # 미리보기 모달 (expander로 구현)
            if st.session_state.get(f'show_preview_{i}', False):
                with st.expander(f"{receipt_set['name']} 미리보기", expanded=True):
                    col_a, col_b = st.columns(2)
                    
                    with col_a:
                        # 이미지 미리보기
                        try:
                            image_data = base64.b64decode(receipt_set['image_data'])
                            image = Image.open(io.BytesIO(image_data))
                            
                            # 미리보기용 리사이즈
                            original_size = image.size
                            display_image = resize_image(image, max_width=300, max_height=400)
                            resized_size = display_image.size
                            
                            st.image(display_image, caption="영수증 이미지")
                            st.caption(f"원본: {original_size[0]}x{original_size[1]} → 표시: {resized_size[0]}x{resized_size[1]}")
                            
                        except Exception as e:
                            st.error(f"이미지를 불러올 수 없습니다: {str(e)}")
                            
                            # 디버깅 정보
                            st.write("디버깅 정보:")
                            st.write(f"- image_data 길이: {len(receipt_set.get('image_data', ''))}")
                            st.write(f"- image_data 타입: {type(receipt_set.get('image_data'))}")
                    
                    with col_b:
                        # 고객 정보 미리보기
                        st.write("**고객 정보:**")
                        st.code(receipt_set['customer_info'])
                    
                    if st.button("닫기", key=f"close_{i}"):
                        st.session_state[f'show_preview_{i}'] = False
                        st.rerun()
            
            st.divider()
        
        # 선택된 세트들만 처리
        st.subheader("배치 처리")

        # 체크박스로 선택
        selected_indices = []

        st.write("처리할 세트 선택:")
        for i, receipt_set in enumerate(st.session_state.receipt_sets):
            # 모든 세트 표시, 하지만 처리 가능한 것만 체크 가능
            can_process = receipt_set['status'] in ['대기', '실패']
            default_checked = can_process  # 처리 가능한 것만 기본 선택
            
            checked = st.checkbox(
                f"{receipt_set['name']} ({receipt_set['status']})",
                key=f"select_{i}",
                value=default_checked,
                disabled=not can_process,  # 처리 불가능한 것은 비활성화
                help="완료된 세트는 다시 처리할 수 없습니다" if not can_process else None
            )
            
            if checked and can_process:
                selected_indices.append(i)

        # 처리 불가능한 세트가 있다면 안내 메시지
        unavailable_sets = [s for s in st.session_state.receipt_sets if s['status'] not in ['대기', '실패']]
        if unavailable_sets:
            st.info(f"💡 {len(unavailable_sets)}개 세트는 이미 처리되었거나 처리 중입니다.")
        
        col_a, col_b, col_c = st.columns([2, 1, 1])
        
        with col_a:
            st.write(f"선택된 세트: {len(selected_indices)}개")
        
        with col_b:
            # 처리 중일 때는 버튼 비활성화
            processing_in_progress = any(s['status'] == '처리중' for s in st.session_state.receipt_sets)
            
            if st.button(
                "🚀 선택된 세트 처리", 
                type="primary",
                disabled=processing_in_progress  # ← 처리 중일 때 비활성화
            ):
                if excel_file and selected_indices:
                    process_batch(selected_indices, excel_file, excel_password)
                else:
                    st.error("엑셀 파일을 업로드하고 처리할 세트를 선택해주세요.")

        with col_c:
            if st.button(
                "🗑️ 전체 삭제",
                disabled=processing_in_progress  # ← 처리 중일 때 비활성화
            ):
                st.session_state.receipt_sets = []
                st.rerun()

        if processing_in_progress:
            st.warning("⚠️ 처리가 진행 중입니다. 완료될 때까지 기다려주세요.")

    else:
        st.info("📝 '데이터 입력' 탭에서 영수증 세트를 추가해주세요.")

with tab3:
    st.header("처리 결과")
    
    # 배치 처리 완료 시 다운로드 버튼
    if st.session_state.batch_processing_complete and st.session_state.batch_result_file:
        st.success("🎉 배치 처리가 완료되었습니다!")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write("**전체 결과 파일이 준비되었습니다.**")
            st.write("모든 처리된 영수증의 고객 정보가 포함된 엑셀 파일을 다운로드할 수 있습니다.")
        
        with col2:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            download_filename = f"배치처리결과_{timestamp}.xlsx"
            
            st.download_button(
                label="📥 전체 결과 다운로드",
                data=st.session_state.batch_result_file,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        st.markdown("---")
    
    # 개별 결과 상세 보기
    if st.session_state.receipt_sets:
        completed_sets = [s for s in st.session_state.receipt_sets if s['status'] == '완료']
        failed_sets = [s for s in st.session_state.receipt_sets if s['status'] == '실패']
        
        # 성공한 세트들
        if completed_sets:
            st.subheader(f"✅ 성공한 처리 ({len(completed_sets)}개)")
            
            for receipt_set in completed_sets:
                with st.expander(f"{receipt_set['name']} - 완료"):
                    if receipt_set['result']:
                        result = receipt_set['result']
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**고객명:** {result['customer_name']}")
                            st.write(f"**매칭 상품:** {result['matched_product']}")
                        with col2:
                            st.write(f"**매칭 점수:** {result['match_score']}")
                            st.write(f"**업데이트 블록:** {result['updated_blocks']}개")
                        
                        st.write(f"**메시지:** {result['message']}")
        
        # 실패한 세트들
        if failed_sets:
            st.subheader(f"❌ 실패한 처리 ({len(failed_sets)}개)")
            
            for receipt_set in failed_sets:
                with st.expander(f"{receipt_set['name']} - 실패"):
                    if receipt_set['result']:
                        st.error(receipt_set['result']['message'])
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**추출된 정보:**")
                            if receipt_set['result'].get('customer_name'):
                                st.write(f"- 고객명: {receipt_set['result']['customer_name']}")
                            if receipt_set['result'].get('receipt_datetime'):
                                st.write(f"- 영수증 시간: {receipt_set['result']['receipt_datetime']}")
                            if receipt_set['result'].get('receipt_product'):
                                st.write(f"- 영수증 상품: {receipt_set['result']['receipt_product']}")
                        
                        with col2:
                            if receipt_set['result'].get('receipt_data'):
                                if st.button(f"영수증 정보 상세", key=f"detail_{receipt_set['id']}"):
                                    st.json(receipt_set['result']['receipt_data'])
                        
                        # Debug 정보 표시 (새로 추가)
                        if receipt_set['result'].get('debug_info'):
                            st.markdown("---")
                            st.write("**매칭 과정 분석:**")
                            
                            debug_info = receipt_set['result']['debug_info']
                            
                            col_debug1, col_debug2 = st.columns(2)
                            with col_debug1:
                                st.write(f"- 전체 주문: {debug_info.get('total_rows', 'N/A')}개")
                                st.write(f"- 날짜 통과: {debug_info.get('date_pass', 'N/A')}개")
                            with col_debug2:
                                st.write(f"- 시간 통과: {debug_info.get('time_pass', 'N/A')}개")
                                st.write(f"- 상품 통과: {debug_info.get('product_pass', 'N/A')}개")
                            
                            if st.button(f"상세 매칭 로그", key=f"debug_detail_{receipt_set['id']}"):
                                st.write("**상위 10개 주문 매칭 시도:**")
                                for i, attempt in enumerate(debug_info.get('all_attempts', [])[:300]):
                                    status = "✅ 매칭 성공" if attempt.get('matched') else f"❌ {attempt.get('skip_reason', '알 수 없음')}"
                                    st.write(f"{i+1}. 주문 {attempt['index']}: {attempt['order_product']} → {status}")
        
        # 결과 초기화 버튼
        if completed_sets or failed_sets:
            st.markdown("---")
            if st.button("🗑️ 모든 결과 초기화"):
                st.session_state.receipt_sets = []
                st.session_state.batch_result_file = None
                st.session_state.batch_processing_complete = False
                st.success("✅ 모든 결과가 초기화되었습니다. 다음 처리는 원본 파일부터 시작됩니다.")
                st.rerun()
    
    else:
        st.info("처리된 결과가 없습니다.")
        st.write("'배치 처리' 탭에서 영수증들을 처리한 후 여기서 결과를 확인하고 다운로드할 수 있습니다.")



