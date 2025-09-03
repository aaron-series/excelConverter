"""
REST API Module

FastAPI를 사용한 Excel to Image Converter REST API
파일 업로드, 변환, 상태 조회, 결과 다운로드 기능을 제공합니다.
"""

import os, sys
import uuid
import asyncio
import logging
import subprocess
from typing import Dict, List, Optional, Any
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

from excel_parser import ExcelParser
from html_renderer import HTMLRenderer
from image_converter import ImageConverter, convert_html_to_image_async


if sys.platform.startswith('win'):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())


# 로깅 설정
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def check_and_install_playwright_browsers():
    """Playwright 브라우저가 설치되어 있는지 확인하고, 없으면 설치합니다."""
    try:
        # Windows에서 asyncio 이벤트 루프 문제 해결
        import platform
        if platform.system() == 'Windows':
            import asyncio
            try:
                # Windows에서 ProactorEventLoop 사용
                if isinstance(asyncio.get_event_loop_policy(), asyncio.WindowsProactorEventLoopPolicy):
                    pass  # 이미 ProactorEventLoop 사용 중
                else:
                    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                logger.info("Windows 이벤트 루프 정책 설정 완료")
            except Exception as e:
                logger.warning(f"Windows 이벤트 루프 정책 설정 실패: {e}")
        
        # playwright 브라우저 설치 확인
        result = subprocess.run(
            ['playwright', '--version'], 
            capture_output=True, 
            text=True, 
            timeout=10
        )
        
        if result.returncode != 0:
            logger.info("Playwright 브라우저를 설치합니다...")
            subprocess.run(['playwright', 'install', 'chromium'], check=True)
            logger.info("Playwright 브라우저 설치 완료")
        else:
            logger.info("Playwright 브라우저가 이미 설치되어 있습니다.")
            
    except subprocess.TimeoutExpired:
        logger.warning("Playwright 버전 확인 시간 초과, 브라우저 설치를 시도합니다...")
        try:
            subprocess.run(['playwright', 'install', 'chromium'], check=True)
            logger.info("Playwright 브라우저 설치 완료")
        except subprocess.CalledProcessError as e:
            logger.error(f"Playwright 브라우저 설치 실패: {e}")
            raise
    except FileNotFoundError:
        logger.error("Playwright가 설치되어 있지 않습니다. 'pip install playwright'를 실행하세요.")
        raise
    except Exception as e:
        logger.error(f"Playwright 브라우저 확인 중 오류: {e}")
        raise


# FastAPI 앱 초기화
app = FastAPI(
    title="Excel to Image Converter API",
    description="Excel 파일을 고품질 이미지로 변환하는 REST API",
    version="1.0.0",
)

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 전역 변수
TASKS: Dict[str, Dict[str, Any]] = {}
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
TEMP_DIR = Path("temp")

# 디렉토리 생성
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)


# Pydantic 모델
class ConversionRequest(BaseModel):
    """변환 요청 모델"""
    
    sheet_name: Optional[str] = None
    sheet_index: Optional[int] = None
    range_start: Optional[str] = None
    range_end: Optional[str] = None
    output_format: str = "png,jpeg,jpg"
    quality: int = 100
    width: Optional[int] = None
    height: Optional[int] = None
    type: str = "image"  # "image" 또는 "html"
    
    class Config:
        schema_extra = {
            "example": {
                "sheet_name": "Sheet1",
                "sheet_index": None,
                "range_start": "A1",
                "range_end": "D10",
                "output_format": "png,jpeg,jpg",
                "quality": 100,
                "width": 1200,
                "height": 800,
                "type": "image|html"
            }
        }


class ConversionResponse(BaseModel):
    """변환 응답 모델"""
    
    task_id: str
    status: str
    message: str
    created_at: datetime
    
    class Config:
        schema_extra = {
            "example": {
                "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
                "status": "pending",
                "message": "변환 작업이 시작되었습니다.",
                "created_at": "2025-01-02T17:33:16.016606"
            }
        }


class TaskStatus(BaseModel):
    """작업 상태 모델"""
    
    task_id: str
    status: str
    progress: int
    message: str
    created_at: datetime
    completed_at: Optional[datetime] = None
    output_file: Optional[str] = None
    error: Optional[str] = None
    
    class Config:
        schema_extra = {
            "example": {
                "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
                "status": "completed",
                "progress": 100,
                "message": "변환 완료",
                "created_at": "2025-01-02T17:33:16.016606",
                "completed_at": "2025-01-02T17:33:18.123456",
                "output_file": "7215cef8-712b-41f9-bf0c-a81438c21e46.png",
                "error": None
            }
        }


class APIInfo(BaseModel):
    """API 정보 모델"""
    
    name: str
    version: str
    description: str
    endpoints: List[str]
    supported_formats: List[str]
    
    class Config:
        schema_extra = {
            "example": {
                "name": "Excel to Image Converter API",
                "version": "1.0.0",
                "description": "Excel 파일을 고품질 이미지로 변환하는 REST API",
                "endpoints": [
                    "POST /upload - 파일 업로드 및 변환 시작",
                    "GET /status/{task_id} - 변환 상태 조회",
                    "GET /download/{task_id} - 결과 파일 다운로드",
                    "GET /tasks - 모든 작업 목록 조회",
                    "DELETE /tasks/{task_id} - 작업 삭제"
                ],
                "supported_formats": ["png", "jpeg", "jpg"]
            }
        }


async def convert_excel_to_image_task(
    task_id: str, file_path: str, request: ConversionRequest
):
    """백그라운드에서 Excel을 이미지로 변환하는 작업"""
    parser = None
    
    try:
        # Windows 이벤트 루프 정책 재확인 (백그라운드 작업 시작 시)
        import platform
        if platform.system() == 'Windows':
            try:
                # 현재 이벤트 루프 정책 확인 및 강제 설정
                current_policy = asyncio.get_event_loop_policy()
                if not isinstance(current_policy, asyncio.WindowsProactorEventLoopPolicy):
                    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                    logger.info(f"백그라운드 작업에서 Windows 이벤트 루프 정책 재설정 완료")
                else:
                    logger.info(f"백그라운드 작업에서 Windows 이벤트 루프 정책 확인됨")
            except Exception as e:
                logger.warning(f"백그라운드 작업에서 Windows 이벤트 루프 정책 설정 실패: {e}")
        
        # 파일 경로 검증
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"업로드된 파일을 찾을 수 없습니다: {file_path}")

        # 시트 이름이 지정되지 않으면 모든 시트 변환
        if request.sheet_name is None:
            await _convert_all_sheets_task(task_id, file_path, request)
        else:
            await _convert_single_sheet_task(task_id, file_path, request)
            
    except Exception as e:
        logger.error(f"작업 {task_id} 실패: {str(e)}")
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["progress"] = 0
        TASKS[task_id]["message"] = f"오류 발생: {str(e)}"
        TASKS[task_id]["completed_at"] = datetime.now()
        TASKS[task_id]["error"] = str(e)
    finally:
        if parser:
            parser.close()


async def _convert_all_sheets_task(task_id: str, file_path: str, request: ConversionRequest):
    """모든 시트를 각각의 이미지로 변환하는 작업"""
    try:
        # 작업 상태 업데이트
        TASKS[task_id]["status"] = "processing"
        TASKS[task_id]["progress"] = 5
        TASKS[task_id]["message"] = "Excel 파일 분석 중..."

        # Excel 파일에서 모든 시트 이름 가져오기
        parser = ExcelParser(file_path)
        if not parser.load_workbook():
            raise Exception("Excel 파일을 로드할 수 없습니다.")
        
        sheet_names = parser.get_sheet_names()
        parser.close()
        
        if not sheet_names:
            raise Exception("Excel 파일에 시트가 없습니다.")
        
        logger.info(f"작업 {task_id}: 총 {len(sheet_names)}개 시트를 변환합니다: {', '.join(sheet_names)}")
        
        success_count = 0
        total_count = len(sheet_names)
        output_files = []
        
        for i, sheet_name in enumerate(sheet_names):
            try:
                # 진행률 업데이트
                progress = 5 + int((i / total_count) * 90)
                TASKS[task_id]["progress"] = progress
                TASKS[task_id]["message"] = f"시트 변환 중: {sheet_name} ({i+1}/{total_count})"
                
                # 각 시트별 출력 파일명 생성
                output_filename = f"{task_id}_{sheet_name}.{request.output_format}"
                output_path = OUTPUT_DIR / output_filename
                
                # 출력 디렉토리 생성 확인
                output_path.parent.mkdir(exist_ok=True)
                
                # 단일 시트 변환
                success = await _convert_single_sheet_internal(
                    file_path, str(output_path), sheet_name, request
                )
                
                if success:
                    success_count += 1
                    output_files.append(output_filename)
                    logger.info(f"작업 {task_id}: 시트 '{sheet_name}' 변환 완료")
                else:
                    logger.error(f"작업 {task_id}: 시트 '{sheet_name}' 변환 실패")
                    
            except Exception as e:
                logger.error(f"작업 {task_id}: 시트 '{sheet_name}' 변환 중 오류: {str(e)}")
        
        # 최종 상태 업데이트
        if success_count > 0:
            TASKS[task_id]["status"] = "completed"
            TASKS[task_id]["progress"] = 100
            TASKS[task_id]["message"] = f"변환 완료 ({success_count}/{total_count} 시트)"
            TASKS[task_id]["completed_at"] = datetime.now()
            TASKS[task_id]["output_file"] = ", ".join(output_files)
            logger.info(f"작업 {task_id} 완료: {success_count}/{total_count} 시트 성공")
        else:
            TASKS[task_id]["status"] = "failed"
            TASKS[task_id]["progress"] = 0
            TASKS[task_id]["message"] = "모든 시트 변환 실패"
            TASKS[task_id]["completed_at"] = datetime.now()
            TASKS[task_id]["error"] = "모든 시트 변환에 실패했습니다."
            
    except Exception as e:
        logger.error(f"모든 시트 변환 작업 {task_id} 실패: {str(e)}")
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["progress"] = 0
        TASKS[task_id]["message"] = f"오류 발생: {str(e)}"
        TASKS[task_id]["completed_at"] = datetime.now()
        TASKS[task_id]["error"] = str(e)


async def _convert_single_sheet_task(task_id: str, file_path: str, request: ConversionRequest):
    """단일 시트를 이미지로 변환하는 작업"""
    try:
        # 작업 상태 업데이트
        TASKS[task_id]["status"] = "processing"
        TASKS[task_id]["progress"] = 10
        TASKS[task_id]["message"] = "Excel 파일 파싱 중..."

        # 출력 파일명 생성 및 경로 검증
        output_filename = f"{task_id}.{request.output_format}"
        output_path = OUTPUT_DIR / output_filename
        
        # 출력 경로 검증
        try:
            output_path_str = str(output_path)
            if not output_path_str or output_path_str.strip() == '':
                raise ValueError("출력 경로가 비어있습니다.")
            
            # 출력 디렉토리 생성 확인
            output_dir = output_path.parent
            output_dir.mkdir(exist_ok=True)
            logger.info(f"출력 디렉토리 확인/생성: {output_dir}")
            
        except Exception as e:
            logger.error(f"출력 경로 설정 실패: {str(e)}")
            raise Exception(f"출력 경로 설정 실패: {str(e)}")
        
        # 단일 시트 변환
        success = await _convert_single_sheet_internal(
            file_path, str(output_path), request.sheet_name, request
        )
        
        if success:
            # 성공 상태 업데이트
            TASKS[task_id]["status"] = "completed"
            TASKS[task_id]["progress"] = 100
            TASKS[task_id]["message"] = "변환 완료"
            TASKS[task_id]["completed_at"] = datetime.now()
            TASKS[task_id]["output_file"] = output_filename
            logger.info(f"작업 {task_id} 완료: {output_filename}")
        else:
            raise Exception("이미지 변환에 실패했습니다.")
            
    except Exception as e:
        logger.error(f"단일 시트 변환 작업 {task_id} 실패: {str(e)}")
        TASKS[task_id]["status"] = "failed"
        TASKS[task_id]["progress"] = 0
        TASKS[task_id]["message"] = f"오류 발생: {str(e)}"
        TASKS[task_id]["completed_at"] = datetime.now()
        TASKS[task_id]["error"] = str(e)


async def _convert_single_sheet_internal(file_path: str, output_path: str, sheet_name: str, request: ConversionRequest) -> bool:
    """단일 시트 변환의 내부 로직"""
    parser = None
    try:
        # 1. Excel 파싱
        parser = ExcelParser(file_path)
        sheet_data = parser.parse_sheet(
            sheet_name=sheet_name,
            sheet_index=request.sheet_index,
            range_start=request.range_start,
            range_end=request.range_end,
        )

        # 2. HTML 렌더링
        renderer = HTMLRenderer()
        html_content = renderer.render_sheet(sheet_data)

        # 3. type 파라미터에 따라 처리
        if request.type.lower() == "html":
            # HTML만 생성
            try:
                # HTML 파일로 저장
                html_output_path = output_path.replace(f".{request.output_format}", ".html")
                with open(html_output_path, "w", encoding="utf-8") as f:
                    f.write(html_content)
                logger.info(f"HTML 파일 생성 완료: {html_output_path}")
                return True
            except Exception as e:
                logger.error(f"HTML 파일 생성 실패: {str(e)}")
                return False
        else:
            # 이미지 변환 (기본값)
            success = await convert_html_to_image_async(
                html_content,
                output_path,
                image_format=request.output_format,
                quality=request.quality,
                width=request.width,
                height=request.height,
            )
            return success
        
    except Exception as e:
        logger.error(f"단일 시트 변환 내부 오류: {str(e)}")
        return False
    finally:
        if parser:
            parser.close()


@app.get("/", response_model=APIInfo)
async def get_api_info():
    """
    API 정보 조회
    
    Excel to Image Converter API의 기본 정보와 사용 가능한 엔드포인트를 제공합니다.
    
    Returns:
        APIInfo: API 기본 정보
    """
    return APIInfo(
        name="Excel to Image Converter API",
        version="1.0.0",
        description="Excel 파일을 고품질 이미지로 변환하는 REST API",
        endpoints=[
            "POST /upload - 파일 업로드 및 변환 시작",
            "GET /status/{task_id} - 변환 상태 조회",
            "GET /download/{task_id} - 결과 파일 다운로드",
            "GET /tasks - 모든 작업 목록 조회",
            "DELETE /tasks/{task_id} - 작업 삭제",
        ],
        supported_formats=["png", "jpeg", "jpg", "html"],
    )


@app.post("/upload", response_model=ConversionResponse)
async def upload_and_convert(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="업로드할 Excel 파일 (.xlsx, .xls)"),
    sheet_name: Optional[str] = None,
    sheet_index: Optional[int] = None,
    range_start: Optional[str] = None,
    range_end: Optional[str] = None,
    output_format: str = "png",
    quality: int = 100,
    width: Optional[int] = None,
    height: Optional[int] = None,
    type: str = "image",
):
    """
    Excel 파일 업로드 및 변환 시작
    
    Excel 파일을 업로드하고 이미지로 변환하는 작업을 시작합니다.
    변환은 백그라운드에서 비동기적으로 처리되며, task_id를 통해 상태를 확인할 수 있습니다.
    
    **헤더 정보:**
    - Content-Type: multipart/form-data
    - Accept: application/json
    
    **요청 예시:**
    ```bash
    curl -X POST "http://localhost:8000/upload" \
         -H "Content-Type: multipart/form-data" \
         -F "file=@sample.xlsx" \
         -F "output_format=png|jpeg|jpg" \
         -F "quality=100" \
         -F "sheet_name=Sheet1"
    ```
    
    **Python 예시:**
    ```python
    import requests
    
    files = {'file': open('sample.xlsx', 'rb')}
    data = {
        'output_format': 'png|jpeg|jpg',
        'quality': 100,
        'sheet_name': 'Sheet1'
    }
    
    response = requests.post('http://localhost:8000/upload', 
                           files=files, data=data)
    print(response.json())
    ```
    
    Args:
        background_tasks: FastAPI 백그라운드 작업
        file: 업로드할 Excel 파일
        sheet_name: 시트 이름 (기본값: 첫 번째 시트)
        sheet_index: 시트 인덱스 (0부터 시작, sheet_name과 함께 사용 불가)
        range_start: 변환할 범위 시작 셀 (예: "A1")
        range_end: 변환할 범위 끝 셀 (예: "D10")
        output_format: 출력 이미지 형식 ("png", "jpeg", "jpg")
        quality: 이미지 품질 (1-100, JPEG에만 적용)
        width: 강제 이미지 너비 (픽셀)
        height: 강제 이미지 높이 (픽셀)
        type: 출력 타입 ("image" 또는 "html", 기본값: "image")
        
    Returns:
        ConversionResponse: 변환 작업 정보
        
    Raises:
        HTTPException: 파일 형식 오류, 크기 초과, 잘못된 매개변수 등
    """
    # 파일 형식 검증
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400, detail="Excel 파일(.xlsx, .xls)만 업로드 가능합니다."
        )

    # 파일 크기 검증 (100MB 제한)
    if file.size > 100 * 1024 * 1024:
        raise HTTPException(
            status_code=400, detail="파일 크기는 100MB를 초과할 수 없습니다."
        )

    # 출력 형식 검증
    if output_format.lower() not in ["png", "jpeg", "jpg"]:
        raise HTTPException(status_code=400, detail="지원하지 않는 출력 형식입니다.")

    # type 파라미터 검증
    if type.lower() not in ["image", "html"]:
        raise HTTPException(status_code=400, detail="type은 'image' 또는 'html'이어야 합니다.")

    # 품질 검증
    if not 1 <= quality <= 100:
        raise HTTPException(
            status_code=400, detail="품질은 1-100 사이의 값이어야 합니다."
        )

    try:
        # 고유 작업 ID 생성
        task_id = str(uuid.uuid4())

        # 파일 저장
        file_path = UPLOAD_DIR / f"{task_id}_{file.filename}"
        with open(file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        # 변환 요청 객체 생성
        request = ConversionRequest(
            sheet_name=sheet_name,
            sheet_index=sheet_index,
            range_start=range_start,
            range_end=range_end,
            output_format=output_format,
            quality=quality,
            width=width,
            height=height,
            type=type,
        )

        # 작업 정보 저장
        TASKS[task_id] = {
            "task_id": task_id,
            "status": "pending",
            "progress": 0,
            "message": "작업 대기 중",
            "created_at": datetime.now(),
            "completed_at": None,
            "output_file": None,
            "error": None,
            "file_path": str(file_path),
            "request": request.dict(),
        }

        # 백그라운드 작업 시작
        background_tasks.add_task(
            convert_excel_to_image_task, task_id, str(file_path), request
        )

        logger.info(f"새로운 변환 작업 시작: {task_id}")

        return ConversionResponse(
            task_id=task_id,
            status="pending",
            message="변환 작업이 시작되었습니다.",
            created_at=TASKS[task_id]["created_at"],
        )

    except Exception as e:
        logger.error(f"파일 업로드 실패: {str(e)}")
        raise HTTPException(
            status_code=500, detail=f"파일 업로드 중 오류가 발생했습니다: {str(e)}"
        )


@app.get("/status/{task_id}", response_model=TaskStatus)
async def get_task_status(task_id: str):
    """
    변환 작업 상태 조회
    
    특정 작업의 현재 상태와 진행률을 조회합니다.
    
    **헤더 정보:**
    - Accept: application/json
    
    **요청 예시:**
    ```bash
    curl -X GET "http://localhost:8000/status/7215cef8-712b-41f9-bf0c-a81438c21e46" \
         -H "Accept: application/json"
    ```
    
    **응답 예시:**
    ```json
    {
        "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
        "status": "completed",
        "progress": 100,
        "message": "변환 완료",
        "created_at": "2025-01-02T17:33:16.016606",
        "completed_at": "2025-01-02T17:33:18.123456",
        "output_file": "7215cef8-712b-41f9-bf0c-a81438c21e46.png",
        "error": null
    }
    ```
    
    Args:
        task_id: 조회할 작업의 고유 ID
        
    Returns:
        TaskStatus: 작업 상태 정보
        
    Raises:
        HTTPException: 작업을 찾을 수 없는 경우
    """
    if task_id not in TASKS:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = TASKS[task_id]
    return TaskStatus(**task)


@app.get("/download/{task_id}")
async def download_result(task_id: str):
    """
    변환 결과 파일 다운로드
    
    완료된 변환 작업의 결과 이미지 파일을 다운로드합니다.
    
    **헤더 정보:**
    - Accept: application/octet-stream
    
    **요청 예시:**
    ```bash
    curl -X GET "http://localhost:8000/download/7215cef8-712b-41f9-bf0c-a81438c21e46" \
         -H "Accept: application/octet-stream" \
         --output result.png
    ```
    
    **Python 예시:**
    ```python
    import requests
    
    response = requests.get('http://localhost:8000/download/7215cef8-712b-41f9-bf0c-a81438c21e46')
    
    if response.status_code == 200:
        with open('result.png', 'wb') as f:
            f.write(response.content)
        print("파일 다운로드 완료")
    ```
    
    Args:
        task_id: 다운로드할 작업의 고유 ID
        
    Returns:
        FileResponse: 이미지 파일 스트림
        
    Raises:
        HTTPException: 작업을 찾을 수 없거나 완료되지 않은 경우
    """
    if task_id not in TASKS:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = TASKS[task_id]

    if task["status"] != "completed":
        raise HTTPException(status_code=400, detail="변환이 완료되지 않았습니다.")

    if not task["output_file"]:
        raise HTTPException(status_code=404, detail="출력 파일을 찾을 수 없습니다.")

    # 출력 파일 경로 생성
    output_files = task["output_file"].split(",")
    if not output_files:
        raise HTTPException(status_code=404, detail="출력 파일을 찾을 수 없습니다.")

    # 첫 번째 파일만 다운로드 (모든 시트가 같은 형식이므로)
    first_output_file = output_files[0]
    file_path = OUTPUT_DIR / first_output_file

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="출력 파일이 존재하지 않습니다.")

    return FileResponse(
        path=file_path,
        filename=first_output_file,
        media_type="application/octet-stream",
    )


@app.get("/tasks", response_model=List[TaskStatus])
async def get_all_tasks():
    """
    모든 작업 목록 조회
    
    현재 시스템에 있는 모든 변환 작업의 목록을 조회합니다.
    
    **헤더 정보:**
    - Accept: application/json
    
    **요청 예시:**
    ```bash
    curl -X GET "http://localhost:8000/tasks" \
         -H "Accept: application/json"
    ```
    
    **응답 예시:**
    ```json
    [
        {
            "task_id": "7215cef8-712b-41f9-bf0c-a81438c21e46",
            "status": "completed",
            "progress": 100,
            "message": "변환 완료",
            "created_at": "2025-01-02T17:33:16.016606",
            "completed_at": "2025-01-02T17:33:18.123456",
            "output_file": "7215cef8-712b-41f9-bf0c-a81438c21e46.png",
            "error": null
        },
        {
            "task_id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
            "status": "processing",
            "progress": 50,
            "message": "이미지 변환 중...",
            "created_at": "2025-01-02T17:35:00.000000",
            "completed_at": null,
            "output_file": null,
            "error": null
        }
    ]
    ```
    
    Returns:
        List[TaskStatus]: 모든 작업 상태 목록
    """
    return [TaskStatus(**task) for task in TASKS.values()]


@app.delete("/tasks/{task_id}")
async def delete_task(task_id: str):
    """
    작업 삭제
    
    특정 변환 작업과 관련된 모든 파일을 삭제합니다.
    
    **헤더 정보:**
    - Accept: application/json
    
    **요청 예시:**
    ```bash
    curl -X DELETE "http://localhost:8000/tasks/7215cef8-712b-41f9-bf0c-a81438c21e46" \
         -H "Accept: application/json"
    ```
    
    **Python 예시:**
    ```python
    import requests
    
    response = requests.delete('http://localhost:8000/tasks/7215cef8-712b-41f9-bf0c-a81438c21e46')
    
    if response.status_code == 200:
        print("작업 삭제 완료")
    ```
    
    Args:
        task_id: 삭제할 작업의 고유 ID
        
    Returns:
        dict: 삭제 완료 메시지
        
    Raises:
        HTTPException: 작업을 찾을 수 없는 경우
    """
    if task_id not in TASKS:
        raise HTTPException(status_code=404, detail="작업을 찾을 수 없습니다.")

    task = TASKS[task_id]

    try:
        # 업로드된 파일 삭제
        if "file_path" in task and os.path.exists(task["file_path"]):
            os.remove(task["file_path"])

        # 출력 파일 삭제
        if task["output_file"]:
            output_files = task["output_file"].split(",")
            for output_file in output_files:
                output_path = OUTPUT_DIR / output_file
                if output_path.exists():
                    output_path.unlink()

        # 작업 정보 삭제
        del TASKS[task_id]

        logger.info(f"작업 {task_id} 삭제 완료")

        return {"message": "작업이 삭제되었습니다."}

    except Exception as e:
        logger.error(f"작업 삭제 실패: {str(e)}")
        raise HTTPException(
            status_code=500, detail=f"작업 삭제 중 오류가 발생했습니다: {str(e)}"
        )


@app.get("/health")
async def health_check():
    """
    헬스 체크
    
    API 서버의 상태를 확인합니다.
    
    **헤더 정보:**
    - Accept: application/json
    
    **요청 예시:**
    ```bash
    curl -X GET "http://localhost:8000/health" \
         -H "Accept: application/json"
    ```
    
    **응답 예시:**
    ```json
    {
        "status": "healthy",
        "timestamp": "2025-01-02T17:40:00.000000"
    }
    ```
    
    Returns:
        dict: 서버 상태 정보
    """
    return {"status": "healthy", "timestamp": datetime.now()}


if __name__ == "__main__":
    # Windows 이벤트 루프 정책 최종 확인
    import platform
    if platform.system() == 'Windows':
        try:
            current_policy = asyncio.get_event_loop_policy()
            if not isinstance(current_policy, asyncio.WindowsProactorEventLoopPolicy):
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                logger.info("서버 시작 시 Windows 이벤트 루프 정책 최종 설정 완료")
            else:
                logger.info("Windows 이벤트 루프 정책이 올바르게 설정되어 있습니다.")
        except Exception as e:
            logger.warning(f"서버 시작 시 Windows 이벤트 루프 정책 설정 실패: {e}")
    
    # Playwright 브라우저 설치 확인
    try:
        check_and_install_playwright_browsers()
    except Exception as e:
        logger.error(f"Playwright 브라우저 설정 실패: {e}")
        logger.error("API 서버를 시작할 수 없습니다.")
        exit(1)
    
    logger.info("Excel to Image Converter API 서버를 시작합니다...")
    logger.info(f"서버 주소: http://0.0.0.0:8000")
    logger.info(f"API 문서: http://0.0.0.0:8000/docs")
    
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=True, log_level="info")
