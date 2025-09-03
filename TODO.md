# Excel to Image Converter 개발 Task List

## Phase 1: 프로젝트 초기 설정

### 1.1 프로젝트 구조 설정
[O] 프로젝트 디렉토리 구조 생성 (Source/, templates/, outputs/)
[O] requirements.txt 파일 생성 및 의존성 정의
[O] README.md 파일 작성 (프로젝트 설명, 설치 가이드)

### 1.2 개발 환경 설정
[O] Python 3.10+ 환경 확인
[O] 가상환경 생성 및 활성화
[O] 핵심 라이브러리 설치 (openpyxl, jinja2, playwright, Pillow)
[O] Playwright 브라우저 설치 (playwright install chromium)

---

## Phase 2: 핵심 모듈 개발

### 2.1 Excel Parser 모듈 (excel_parser.py)
[O] openpyxl을 사용한 Excel 파일 읽기 기능 구현
[O] 셀 데이터 추출 (값, 타입, 포맷)
[O] 스타일 정보 추출 (폰트, 색상, 배경색, 테두리)
[O] 병합 셀 정보 처리
[O] 행 높이/열 너비 정보 추출
[O] 시트 선택 및 범위 지정 기능
[O] 에러 처리 (파일 없음, 손상된 파일 등)

### 2.2 HTML Renderer 모듈 (html_renderer.py)
[O] jinja2 템플릿 엔진 설정
[O] templates/sheet.html 템플릿 생성
[O] Excel 스타일을 CSS로 매핑하는 로직 구현
[O] 폰트 스타일 (크기, 굵기, 기울임, 밑줄)
[O] 색상 (텍스트 색상, 배경색)
[O] 테두리 (스타일, 두께, 색상)
[O] 정렬 (가로, 세로, 텍스트 줄바꿈)
[O] 병합 셀 HTML 구조 생성
[O] 동적 크기 조정 로직 (행 높이/열 너비)

### 2.3 Image Converter 모듈 (image_converter.py)
[O] Playwright 브라우저 설정 (헤드리스 모드)
[O] HTML 파일 렌더링 기능
[O] 동적 viewport 크기 조정 로직
[O] 스크린샷 캡처 기능
[O] Pillow를 사용한 이미지 후처리
[O] PNG/JPEG 포맷 변환
[O] 이미지 압축 및 품질 조정
[O] 투명도 처리
[O] 파일 저장 기능

---

## Phase 3: 통합 및 메인 로직

### 3.1 메인 애플리케이션 (main.py)
[O] 전체 변환 파이프라인 구현
[O] 모듈 간 데이터 전달 로직
[O] 에러 처리 및 로깅
[O] 설정 관리 (기본값, 사용자 옵션)

### 3.2 CLI 인터페이스
[O] argparse를 사용한 명령행 인터페이스 구현
[O] 입력 파일 경로 처리
[O] 출력 옵션 (포맷, 품질, 저장 경로)
[O] 시트/범위 선택 옵션
[O] 배치 처리 기능

### 3.3 추가 개발
[O] 테스트 스크립트 생성 (test_converter.py)
[O] 사용법 가이드 작성 (USAGE.md)
[O] requirements.txt 파일 정리

---

<!-- ## Phase 4: 테스트 및 품질 보증 (생략)

### 4.1 단위 테스트
[ ] excel_parser.py 단위 테스트 작성
[ ] html_renderer.py 단위 테스트 작성
[ ] image_converter.py 단위 테스트 작성
[ ] 테스트용 Excel 파일 준비 (다양한 스타일, 병합 셀 등)

### 4.2 통합 테스트
[ ] 전체 변환 파이프라인 테스트
[ ] 다양한 Excel 파일 형식 테스트 (.xlsx, .xls)
[ ] 복잡한 스타일이 적용된 파일 테스트
[ ] 대용량 파일 처리 테스트

### 4.3 성능 최적화
[ ] 메모리 사용량 최적화
[ ] 처리 속도 개선
[ ] async 기반 Playwright 구현 검토 -->

---

## Phase 5: API 및 고급 기능

### 5.1 REST API 개발 (api.py)
[O] FastAPI를 사용한 REST API 구현
[O] API Request/Response Parameter 정보 엔드포인트
[O] 파일 업로드 엔드포인트
[O] 변환 상태 조회 엔드포인트
[O] 결과 다운로드 엔드포인트
[O] 작업 목록 조회 및 삭제 엔드포인트
[O] 헬스 체크 엔드포인트

### 5.2 고급 기능
[O] 배치 처리 최적화 (batch_processor.py)
[O] 진행률 표시 기능
[O] 동시 처리 제한 (세마포어)
[O] 실시간 진행률 추적
[O] 처리 통계 및 성능 분석
[O] API 테스트 스크립트 (test_api.py)

---