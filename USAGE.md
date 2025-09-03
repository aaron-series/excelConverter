# Excel to Image Converter - 상세 사용법

## 📖 목차

1. [개요](#개요)
2. [설치 및 설정](#설치-및-설정)
3. [CLI 모드 사용법](#cli-모드-사용법)
4. [REST API 사용법](#rest-api-사용법)
5. [배치 처리](#배치-처리)
6. [리눅스 배포](#리눅스-배포)
7. [출력 파일](#출력-파일)
8. [문제 해결](#문제-해결)
9. [성능 최적화](#성능-최적화)
10. [API 문서](#api-문서)

## 개요

Excel to Image Converter는 Excel 파일을 고품질 이미지 또는 HTML로 변환하는 도구입니다. Playwright를 사용하여 정확한 렌더링을 보장하며, CLI와 REST API 두 가지 모드를 제공합니다.

### 지원 형식

**입력 형식:**
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

**출력 형식:**
- **이미지**: PNG (기본값, 투명도 지원), JPEG (압축률 조절 가능)
- **HTML**: 완전한 HTML 문서 (브라우저에서 열기 가능)

## 설치 및 설정

### 1. 기본 설치

```bash
# 저장소 클론
git clone <repository-url>
cd excel2image

# 의존성 설치
pip install -r requirements.txt

# Playwright 브라우저 설치
playwright install chromium
```

### 2. 가상환경 사용 (권장)

```bash
# 가상환경 생성
python -m venv venv

# 가상환경 활성화
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt
```

### 3. 시스템 요구사항

- **Python**: 3.10 이상
- **메모리**: 최소 2GB, 권장 4GB
- **디스크**: 최소 1GB 여유 공간
- **네트워크**: API 모드 시 인터넷 연결

## CLI 모드 사용법

### 기본 명령어

```bash
python main.py <excel_file> [options]
```

### 명령행 옵션

| 옵션 | 설명 | 기본값 | 예시 |
|------|------|--------|------|
| `input` | 입력 Excel 파일 경로 | 필수 | `sample.xlsx` |
| `-o, --output` | 출력 파일 경로 | 자동 생성 | `output.png` |
| `-s, --sheet` | 시트 이름 | 모든 시트 | `Sheet1` |
| `-i, --sheet-index` | 시트 인덱스 (0부터) | 첫 번째 시트 | `0` |
| `-r, --range` | 변환할 셀 범위 | 전체 시트 | `A1:D10` |
| `-f, --format` | 출력 형식 | `png` | `png`, `jpeg` |
| `--quality` | 이미지 품질 (1-100) | `100` | `95` |
| `--width` | 강제 이미지 너비 | 자동 | `1200` |
| `--height` | 강제 이미지 높이 | 자동 | `800` |
| `--type` | 출력 타입 | `image` | `image`, `html` |
| `--batch` | 배치 처리 모드 | `False` | `--batch` |
| `--recursive` | 하위 디렉토리 포함 | `False` | `--recursive` |

### 사용 예시

#### 1. 모든 시트를 이미지로 변환 (기본)
```bash
python main.py report.xlsx
# 결과: report_Sheet1.png, report_Sheet2.png, report_Sheet3.png, ...
```

#### 2. 모든 시트를 HTML로 변환
```bash
python main.py report.xlsx --type html
# 결과: report_Sheet1.html, report_Sheet2.html, report_Sheet3.html, ...
```

#### 3. 특정 시트만 이미지로 변환
```bash
python main.py report.xlsx -s "Summary"
# 결과: report_Summary.png
```

#### 4. 특정 시트만 HTML로 변환
```bash
python main.py report.xlsx -s "Summary" --type html
# 결과: report_Summary.html
```

#### 5. 출력 파일명 지정 (이미지)
```bash
python main.py report.xlsx -o output.png
# 결과: output_Sheet1.png, output_Sheet2.png, output_Sheet3.png, ...
```

#### 6. 출력 파일명 지정 (HTML)
```bash
python main.py report.xlsx -o output.html --type html
# 결과: output_Sheet1.html, output_Sheet2.html, output_Sheet3.html, ...
```

#### 7. 특정 범위 변환
```bash
python main.py report.xlsx -s "Data" -r "A1:D20"
# 결과: report_Data.png (A1:D20 범위만)
```

#### 8. JPEG 형식으로 변환
```bash
python main.py report.xlsx -f jpeg --quality 90
# 결과: report_Sheet1.jpg, report_Sheet2.jpg, ... (90% 품질)
```

#### 9. 배치 처리 (이미지)
```bash
python main.py *.xlsx --batch
# 모든 Excel 파일의 모든 시트를 이미지로 변환
```

#### 10. 배치 처리 (HTML)
```bash
python main.py *.xlsx --batch --type html
# 모든 Excel 파일의 모든 시트를 HTML로 변환
```

## REST API 사용법

### 서버 시작

```bash
# 개발 모드
python api.py

# 프로덕션 모드
uvicorn api:app --host 0.0.0.0 --port 8000
```

### API 엔드포인트

#### 1. 파일 업로드 및 변환 시작

**이미지 변환 (기본):**
```bash
curl -X POST "http://localhost:8000/upload" \
     -H "Content-Type: multipart/form-data" \
     -F "file=@sample.xlsx" \
     -F "output_format=png" \
     -F "quality=100"
```

**HTML 변환:**
```bash
curl -X POST "http://localhost:8000/upload" \
     -H "Content-Type: multipart/form-data" \
     -F "file=@sample.xlsx" \
     -F "type=html"
```

**응답:**
```json
{
    "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
    "status": "pending",
    "message": "변환 작업이 시작되었습니다.",
    "created_at": "2025-01-02T17:33:16.016606"
}
```

#### 2. 변환 상태 조회

```bash
curl -X GET "http://localhost:8000/status/7215cef8-712b-41f9-bf0c-a81438c21e46"
```

**이미지 변환 완료 응답:**
```json
{
    "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
    "status": "completed",
    "progress": 100,
    "message": "변환 완료 (3/4 시트)",
    "created_at": "2025-01-02T17:33:16.016606",
    "completed_at": "2025-01-02T17:33:18.123456",
    "output_file": "7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet1.png, 7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet2.png, 7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet3.png",
    "error": null
}
```

**HTML 변환 완료 응답:**
```json
{
    "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
    "status": "completed",
    "progress": 100,
    "message": "변환 완료 (3/4 시트)",
    "created_at": "2025-01-02T17:33:16.016606",
    "completed_at": "2025-01-02T17:33:18.123456",
    "output_file": "7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet1.html, 7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet2.html, 7215cef8-712b-41f9-bf0c-a81438c21e46_Sheet3.html",
    "error": null
}
```

#### 3. 결과 파일 다운로드

**이미지 파일 다운로드:**
```bash
curl -X GET "http://localhost:8000/download/7215cef8-712b-41f9-bf0c-a81438c21e46" \
     --output result.png
```

**HTML 파일 다운로드:**
```bash
curl -X GET "http://localhost:8000/download/7215cef8-712b-41f9-bf0c-a81438c21e46" \
     --output result.html
```

#### 4. 모든 작업 목록 조회

```bash
curl -X GET "http://localhost:8000/tasks"
```

#### 5. 작업 삭제

```bash
curl -X DELETE "http://localhost:8000/tasks/7215cef8-712b-41f9-bf0c-a81438c21e46"
```

### API 파라미터

| 파라미터 | 타입 | 필수 | 기본값 | 설명 |
|----------|------|------|--------|------|
| `file` | File | ✅ | - | 업로드할 Excel 파일 |
| `sheet_name` | string | ❌ | - | 변환할 시트 이름 |
| `sheet_index` | integer | ❌ | - | 변환할 시트 인덱스 (0부터) |
| `range_start` | string | ❌ | - | 변환할 범위 시작 셀 (예: "A1") |
| `range_end` | string | ❌ | - | 변환할 범위 끝 셀 (예: "D10") |
| `output_format` | string | ❌ | "png" | 출력 이미지 형식 ("png", "jpeg", "jpg") |
| `quality` | integer | ❌ | 100 | 이미지 품질 (1-100, JPEG에만 적용) |
| `width` | integer | ❌ | - | 강제 이미지 너비 (픽셀) |
| `height` | integer | ❌ | - | 강제 이미지 높이 (픽셀) |
| `type` | string | ❌ | "image" | 출력 타입 ("image" 또는 "html") |

## 배치 처리

### CLI 배치 처리

```bash
# 모든 Excel 파일을 이미지로 변환
python main.py "*.xlsx" --batch

# 모든 Excel 파일을 HTML로 변환
python main.py "*.xlsx" --batch --type html

# 하위 디렉토리 포함하여 이미지로 변환
python main.py "*.xlsx" --batch --recursive

# 특정 시트만 배치 처리
python main.py "*.xlsx" --batch -s "Summary"
```

### API 배치 처리

```bash
# 여러 파일을 순차적으로 업로드
for file in *.xlsx; do
    curl -X POST "http://localhost:8000/upload" \
         -F "file=@$file" \
         -F "type=html"
done
```

## 리눅스 배포

### 배포 패키지 구조

```
x2i/
├── conf/           # 환경설정 파일
├── bin/            # 실행파일, pid파일 등
├── lib/            # 패키지 또는 라이브러리
├── logs/           # 실행 로그 파일
└── data/
    ├── uploads/    # 업로드 파일 저장
    └── outputs/    # 결과 파일 저장
```

### 설치 및 실행

```bash
# 1. 환경 설정
cd x2i
chmod +x setup.sh
./setup.sh

# 2. 서비스 시작
chmod +x service.sh
./service.sh start

# 3. 서비스 상태 확인
./service.sh status

# 4. 서비스 중지
./service.sh stop
```

### 환경 변수

```bash
export X2I_HOME="/path/to/x2i"
export X2I_DATA_DIR="/path/to/x2i/data"
export X2I_LOG_DIR="/path/to/x2i/logs"
```

## 출력 파일

### 이미지 파일

- **형식**: PNG, JPEG
- **위치**: `outputs/` 디렉토리
- **명명 규칙**: `{원본파일명}_{시트명}.{확장자}`
- **특징**: 
  - 고해상도 렌더링
  - 투명도 지원 (PNG)
  - 압축률 조절 가능 (JPEG)

### HTML 파일

- **형식**: 완전한 HTML 문서
- **위치**: `outputs/` 디렉토리
- **명명 규칙**: `{원본파일명}_{시트명}.html`
- **특징**:
  - 브라우저에서 바로 열기 가능
  - 반응형 디자인
  - 인쇄 최적화
  - 고해상도 디스플레이 대응

### 파일 예시

```
outputs/
├── report_Sheet1.png
├── report_Sheet1.html
├── report_Sheet2.png
├── report_Sheet2.html
├── data_Summary.jpg
└── data_Summary.html
```

## 문제 해결

### 일반적인 문제

#### 1. Playwright 브라우저 오류
```bash
# 브라우저 재설치
playwright install chromium

# 또는 모든 브라우저 설치
playwright install
```

#### 2. 메모리 부족 오류
```bash
# 배치 처리 시 워커 수 줄이기
python main.py "*.xlsx" --batch --max-workers 2
```

#### 3. 권한 오류 (Linux)
```bash
# 실행 권한 부여
chmod +x main.py
chmod +x api.py
```

#### 4. 포트 충돌 (API)
```bash
# 다른 포트 사용
uvicorn api:app --port 8001
```

### 로그 확인

```bash
# CLI 로그 레벨 설정
python main.py sample.xlsx --log-level DEBUG

# API 로그 확인
tail -f logs/api.log
```

## 성능 최적화

### 1. 메모리 최적화

```bash
# 배치 처리 시 워커 수 조절
python main.py "*.xlsx" --batch --max-workers 3

# 이미지 품질 조절
python main.py sample.xlsx --quality 85
```

### 2. 속도 최적화

```bash
# 특정 범위만 변환
python main.py sample.xlsx -r "A1:D100"

# 특정 시트만 변환
python main.py sample.xlsx -s "Summary"
```

### 3. 디스크 공간 최적화

```bash
# JPEG 형식 사용 (압축률 높임)
python main.py sample.xlsx -f jpeg --quality 80

# 정기적인 출력 파일 정리
find outputs/ -name "*.png" -mtime +7 -delete
```

## API 문서

### Swagger UI

API 문서는 서버 실행 후 다음 URL에서 확인할 수 있습니다:

```
http://localhost:8000/docs
```

### OpenAPI 스키마

```bash
# OpenAPI 스키마 다운로드
curl -X GET "http://localhost:8000/openapi.json" > api_schema.json
```

### 지원하는 형식

- **입력**: `.xlsx`, `.xls`
- **출력 이미지**: `png`, `jpeg`, `jpg`
- **출력 HTML**: `html`
- **품질**: 1-100 (JPEG에만 적용)

### 응답 코드

| 코드 | 설명 |
|------|------|
| 200 | 성공 |
| 400 | 잘못된 요청 |
| 404 | 리소스를 찾을 수 없음 |
| 422 | 유효성 검사 실패 |
| 500 | 서버 내부 오류 |

### 예시 요청

#### Postman

1. **POST** `/upload`
2. **Body** → **form-data**
3. **Key**: `file` (Type: File)
4. **Key**: `type` (Type: Text, Value: `html`)
5. **Key**: `sheet_name` (Type: Text, Value: `Sheet1`)

#### cURL

```bash
curl -X POST "http://localhost:8000/upload" \
     -F "file=@sample.xlsx" \
     -F "type=html" \
     -F "sheet_name=Sheet1" \
     -F "output_format=png" \
     -F "quality=100"
```

### 헤더 정보

**요청 헤더:**
```
Content-Type: multipart/form-data
```

**응답 헤더:**
```
Content-Type: application/json
Content-Length: <size>
```

### 에러 응답 예시

```json
{
    "detail": "지원하지 않는 파일 형식입니다."
}
```

```json
{
    "detail": "type은 'image' 또는 'html'이어야 합니다."
}
```

```json
{
    "detail": "품질은 1-100 사이의 값이어야 합니다."
}
```
