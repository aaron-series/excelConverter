# Excel to Image Converter

Excel 파일을 고품질 이미지 또는 HTML로 변환하는 Python 애플리케이션입니다. CLI 모드와 REST API 모드를 모두 지원하며, Windows와 Linux 환경에서 안정적으로 작동합니다.

## 🚀 주요 기능

- **Excel 파일 지원**: `.xlsx`, `.xls` 형식 지원
- **다양한 출력 형식**: 
  - **이미지**: PNG (투명도 지원), JPEG (압축률 조절)
  - **HTML**: 완전한 HTML 문서 (브라우저에서 열기 가능)
- **고품질 변환**: Playwright를 사용한 정확한 렌더링
- **CLI 모드**: 명령줄에서 직접 사용
- **REST API**: 웹 서비스로 제공
- **배치 처리**: 여러 파일 동시 변환
- **크로스 플랫폼**: Windows, Linux 지원
- **동적 셀 너비**: 데이터 길이에 맞는 자동 셀 너비 조정
- **병합 셀 지원**: Excel의 병합 셀을 정확히 렌더링

## 📋 요구사항

- Python 3.10 이상
- Playwright 브라우저 (자동 설치)
- 최소 2GB RAM 권장

## 🛠️ 설치 방법

### 1. 저장소 클론
```bash
git clone <repository-url>
cd excel2image
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. Playwright 브라우저 설치
```bash
playwright install chromium
```

## 🎯 사용법

### CLI 모드

#### 기본 사용법 (이미지 변환)
```bash
python main.py sample.xlsx
# 결과: sample_Sheet1.png, sample_Sheet2.png, ...
```

#### HTML 변환
```bash
python main.py sample.xlsx --type html
# 결과: sample_Sheet1.html, sample_Sheet2.html, ...
```

#### 고급 옵션
```bash
python main.py \
    --input sample.xlsx \
    --output sample.png \
    --sheet "Sheet1" \
    --range "A1:D10" \
    --format png \
    --quality 100 \
    --width 1200 \
    --height 800 \
    --type image
```

#### 배치 처리
```bash
# 이미지 변환
python main.py "*.xlsx" --batch

# HTML 변환
python main.py "*.xlsx" --batch --type html
```

### REST API 모드

#### 서버 시작
```bash
python api.py
```

#### API 엔드포인트
- `POST /upload` - 파일 업로드 및 변환
- `GET /status/{task_id}` - 변환 상태 조회
- `GET /download/{task_id}` - 결과 파일 다운로드
- `GET /tasks` - 모든 작업 목록
- `DELETE /tasks/{task_id}` - 작업 삭제

#### API 사용 예시

**이미지 변환:**
```bash
curl -X POST "http://localhost:8000/upload" \
     -F "file=@sample.xlsx" \
     -F "type=image" \
     -F "output_format=png"
```

**HTML 변환:**
```bash
curl -X POST "http://localhost:8000/upload" \
     -F "file=@sample.xlsx" \
     -F "type=html"
```

#### API 문서
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## 📁 프로젝트 구조

```
excel2image/
├── Source/
│   ├── main.py              # CLI 메인 애플리케이션
│   ├── api.py               # REST API 서버
│   ├── excel_parser.py      # Excel 파싱 모듈
│   ├── html_renderer.py     # HTML 렌더링 모듈
│   ├── image_converter.py   # 이미지 변환 모듈
│   ├── batch_processor.py   # 배치 처리 모듈
│   ├── templates/
│   │   └── sheet.html       # HTML 템플릿
│   ├── requirements.txt     # Python 의존성
│   ├── README.md           # 프로젝트 문서
│   └── USAGE.md            # 상세 사용법
├── x2i/                    # Linux 배포 패키지
│   ├── conf/               # 환경설정 파일
│   ├── bin/                # 실행파일, pid파일 등
│   ├── lib/                # 패키지 또는 라이브러리
│   ├── logs/               # 실행 로그 파일
│   ├── data/               # 데이터 디렉토리
│   │   ├── uploads/        # 업로드 파일 저장
│   │   └── outputs/        # 결과 파일 저장
│   ├── setup.sh            # 환경 설정 스크립트
│   └── service.sh          # 서비스 관리 스크립트
└── outputs/                # 변환 결과 파일
```

## 🔧 주요 기능 상세

### 1. 동적 셀 너비 조정

- **데이터 기반 계산**: 각 셀의 내용 길이에 따라 자동으로 너비 조정
- **문자별 가중치**: 한글, 영문, 숫자, 특수문자별로 다른 픽셀 가중치 적용
- **개행 처리**: 개행된 텍스트에서 가장 긴 줄만을 기준으로 너비 계산
- **병합 셀 독립성**: 병합된 셀이 있어도 다른 셀들의 독립적인 너비 보장

### 2. HTML 출력 기능

- **완전한 HTML 문서**: DOCTYPE, head, body 포함
- **반응형 디자인**: 모바일 및 데스크톱 최적화
- **인쇄 지원**: CSS 미디어 쿼리로 인쇄 최적화
- **메타데이터 포함**: 시트 정보, 생성 시간, 범위 정보
- **브라우저 호환성**: 모든 최신 브라우저에서 정상 작동

### 3. 고품질 이미지 렌더링

- **Playwright 엔진**: Chromium 기반 정확한 렌더링
- **고해상도 지원**: 2x, 3x 해상도 디스플레이 대응
- **투명도 지원**: PNG 형식에서 투명 배경 지원
- **압축률 조절**: JPEG 형식에서 품질 조절 가능

### 4. 배치 처리 최적화

- **동시 처리**: 멀티프로세싱을 통한 병렬 처리
- **진행률 표시**: 실시간 진행 상황 모니터링
- **오류 처리**: 개별 파일 오류가 전체 배치에 영향 없음
- **메모리 관리**: 효율적인 리소스 사용

## 🚀 성능 최적화

### 메모리 최적화
```bash
# 배치 처리 시 워커 수 조절
python main.py "*.xlsx" --batch --max-workers 3

# 이미지 품질 조절
python main.py sample.xlsx --quality 85
```

### 속도 최적화
```bash
# 특정 범위만 변환
python main.py sample.xlsx -r "A1:D100"

# 특정 시트만 변환
python main.py sample.xlsx -s "Summary"
```

## 🔍 문제 해결

### 일반적인 문제

1. **Playwright 브라우저 오류**
   ```bash
   playwright install chromium
   ```

2. **메모리 부족**
   ```bash
   python main.py "*.xlsx" --batch --max-workers 2
   ```

3. **권한 오류 (Linux)**
   ```bash
   chmod +x main.py
   chmod +x api.py
   ```

### 로그 확인
```bash
# CLI 로그 레벨 설정
python main.py sample.xlsx --log-level DEBUG

# API 로그 확인
tail -f logs/api.log
```

## 📊 지원 형식

### 입력 형식
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

### 출력 형식
- **이미지**: PNG, JPEG, JPG
- **HTML**: 완전한 HTML 문서

### 품질 설정
- **PNG**: 무손실 압축, 투명도 지원
- **JPEG**: 1-100 품질 설정 가능

## 🌐 API 응답 코드

| 코드 | 설명 |
|------|------|
| 200 | 성공 |
| 400 | 잘못된 요청 |
| 404 | 리소스를 찾을 수 없음 |
| 422 | 유효성 검사 실패 |
| 500 | 서버 내부 오류 |

## 📝 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

## 🤝 기여

버그 리포트, 기능 요청, 풀 리퀘스트를 환영합니다!

## 📞 지원

문제가 발생하면 이슈를 생성하거나 로그 파일을 확인해주세요.

## 📚 추가 문서

- [상세 사용법 (USAGE.md)](USAGE.md) - CLI, API, 배치 처리 등 상세 가이드
- [API 문서](http://localhost:8000/docs) - 서버 실행 후 확인 가능
