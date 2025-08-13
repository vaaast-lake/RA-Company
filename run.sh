#!/bin/bash

# ==============================================
# Streamlit App 실행 스크립트 (Windows + Git Bash)
# UV 가상환경을 활성화하고 Streamlit 앱을 실행합니다
# ==============================================

# 색상 정의
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 로그 함수들
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 스크립트 시작
echo "================================================"
echo "🚀 Streamlit 영수증 매칭 시스템 실행기 (Git Bash)"
echo "================================================"

# 현재 디렉토리 확인
CURRENT_DIR=$(pwd)
log_info "현재 디렉토리: $CURRENT_DIR"

# 필수 파일 존재 확인
if [ ! -f "app.py" ]; then
    log_error "app.py 파일을 찾을 수 없습니다."
    log_info "스크립트를 app.py가 있는 디렉토리에서 실행해주세요."
    read -p "Press Enter to continue..."
    exit 1
fi

if [ ! -f "pyproject.toml" ]; then
    log_error "pyproject.toml 파일을 찾을 수 없습니다."
    log_info "이 스크립트는 완성된 UV 프로젝트 디렉토리에서 실행해야 합니다."
    read -p "Press Enter to continue..."
    exit 1
fi

log_success "필수 파일 확인 완료!"

# UV 설치 확인
if ! command -v uv.exe &> /dev/null && ! command -v uv &> /dev/null; then
    log_error "UV가 설치되어 있지 않습니다."
    log_info "UV 설치 방법: powershell -c \"irm https://astral.sh/uv/install.ps1 | iex\""
    read -p "Press Enter to continue..."
    exit 1
fi

# UV 명령어 설정
UV_CMD="uv"
if command -v uv.exe &> /dev/null; then
    UV_CMD="uv.exe"
fi

log_success "UV 설치 확인 완료"

# 의존성 동기화
log_info "의존성 동기화 중..."
echo "────────────────────────────────────────────────"

$UV_CMD sync

if [ $? -eq 0 ]; then
    log_success "의존성 동기화 완료"
else
    log_error "의존성 동기화 실패"
    log_info "수동으로 'uv sync'를 실행하거나 pyproject.toml을 확인해주세요."
    read -p "Press Enter to continue..."
    exit 1
fi

echo "────────────────────────────────────────────────"

# .env 파일 확인 (선택사항)
if [ ! -f ".env" ]; then
    log_warning ".env 파일이 없습니다."
    log_info "OpenAI API 키 등 환경변수가 필요한 경우 .env 파일을 생성해주세요."
    echo ""
fi

# Streamlit 실행
log_info "🎯 Streamlit 앱 실행 중..."
echo "================================================"
echo "✨ 브라우저에서 http://localhost:8501 접속"
echo "🛑 종료하려면 Ctrl+C 를 누르세요"
echo "================================================"
echo ""

# UV로 Streamlit 실행
$UV_CMD run streamlit run app.py

# 실행 완료 후 메시지
echo ""
echo "================================================"
log_info "Streamlit 앱이 종료되었습니다."
echo "================================================"
read -p "Press Enter to continue..."
