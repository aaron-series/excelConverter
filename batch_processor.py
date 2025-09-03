"""
Batch Processor Module

배치 처리 최적화 및 진행률 표시 기능
여러 Excel 파일을 효율적으로 처리하고 실시간 진행률을 제공합니다.
"""

import asyncio
import logging
import time
import platform
from typing import List, Dict, Any, Optional, Callable
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import threading

from excel_parser import ExcelParser
from html_renderer import HTMLRenderer
from image_converter import ImageConverter

logger = logging.getLogger(__name__)

class ProgressCallback:
    """진행률 콜백 클래스"""
    
    def __init__(self, total_files: int):
        self.total_files = total_files
        self.completed_files = 0
        self.failed_files = 0
        self.current_file = ""
        self.start_time = time.time()
        self.lock = threading.Lock()
    
    def update(self, file_name: str, status: str, progress: int = 0):
        """진행률 업데이트"""
        with self.lock:
            self.current_file = file_name
            if status == "completed":
                self.completed_files += 1
            elif status == "failed":
                self.failed_files += 1
            
            elapsed_time = time.time() - self.start_time
            if self.completed_files + self.failed_files > 0:
                avg_time_per_file = elapsed_time / (self.completed_files + self.failed_files)
                remaining_files = self.total_files - (self.completed_files + self.failed_files)
                estimated_remaining_time = remaining_files * avg_time_per_file
            else:
                estimated_remaining_time = 0
            
            logger.info(
                f"진행률: {self.completed_files + self.failed_files}/{self.total_files} "
                f"({((self.completed_files + self.failed_files) / self.total_files * 100):.1f}%) "
                f"| 현재: {file_name} | 상태: {status} | "
                f"예상 남은 시간: {estimated_remaining_time:.1f}초"
            )
    
    def get_summary(self) -> Dict[str, Any]:
        """진행률 요약 반환"""
        with self.lock:
            total_processed = self.completed_files + self.failed_files
            success_rate = (self.completed_files / self.total_files * 100) if self.total_files > 0 else 0
            elapsed_time = time.time() - self.start_time
            
            return {
                "total_files": self.total_files,
                "completed_files": self.completed_files,
                "failed_files": self.failed_files,
                "success_rate": success_rate,
                "elapsed_time": elapsed_time,
                "current_file": self.current_file
            }

class BatchProcessor:
    """배치 처리 클래스"""
    
    def __init__(self, max_workers: int = 3, output_dir: str = "outputs"):
        """
        BatchProcessor 초기화
        
        Args:
            max_workers (int): 최대 동시 처리 작업 수
            output_dir (str): 출력 디렉토리
        """
        self.max_workers = max_workers
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # 공유 리소스
        self.parser = ExcelParser()
        self.renderer = HTMLRenderer()
        self.converter = None
        
        # 결과 저장
        self.results: List[Dict[str, Any]] = []
        self.lock = threading.Lock()
    
    async def initialize(self):
        """초기화"""
        # Windows에서 asyncio 이벤트 루프 문제 해결
        if platform.system() == 'Windows':
            try:
                # Windows에서 ProactorEventLoop 사용
                if isinstance(asyncio.get_event_loop_policy(), asyncio.WindowsProactorEventLoopPolicy):
                    pass  # 이미 ProactorEventLoop 사용 중
                else:
                    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
                logger.info("Windows 이벤트 루프 정책 설정 완료")
            except Exception as e:
                logger.warning(f"Windows 이벤트 루프 정책 설정 실패: {e}")
        
        self.converter = ImageConverter()
        await self.converter.initialize()
    
    async def close(self):
        """리소스 정리"""
        if self.converter:
            await self.converter.close()
    
    async def process_single_file(
        self, 
        file_path: Path, 
        output_format: str = "png",
        quality: int = 100,
        sheet_name: Optional[str] = None,
        range_start: Optional[str] = None,
        range_end: Optional[str] = None,
        type: str = "image",
        progress_callback: Optional[ProgressCallback] = None
    ) -> Dict[str, Any]:
        """단일 파일 처리"""
        result = {
            "file_path": str(file_path),
            "file_name": file_path.name,
            "status": "processing",
            "start_time": datetime.now(),
            "end_time": None,
            "error": None,
            "output_file": None
        }
        
        try:
            if progress_callback:
                progress_callback.update(file_path.name, "started", 10)
            
            # 1. Excel 파싱
            sheet_data = self.parser.parse_sheet(
                str(file_path),
                sheet_name=sheet_name,
                range_start=range_start,
                range_end=range_end
            )
            
            if progress_callback:
                progress_callback.update(file_path.name, "parsing", 30)
            
            # 2. HTML 렌더링
            html_content = self.renderer.render_sheet(sheet_data)
            
            if progress_callback:
                progress_callback.update(file_path.name, "rendering", 50)
            
            # 3. type 파라미터에 따라 처리
            if type.lower() == "html":
                # HTML만 생성
                output_filename = f"{file_path.stem}.html"
                output_path = self.output_dir / output_filename
                
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
                
                # HTML 파일로 저장
                try:
                    with open(output_path, "w", encoding="utf-8") as f:
                        f.write(html_content)
                    success = True
                    logger.info(f"HTML 파일 생성 완료: {output_path}")
                except Exception as e:
                    logger.error(f"HTML 파일 생성 실패: {str(e)}")
                    success = False
            else:
                # 이미지 변환 (기본값)
                output_filename = f"{file_path.stem}.{output_format}"
                output_path = self.output_dir / output_filename
                
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
                
                success = await self.converter.convert_html_to_image(
                    html_content,
                    str(output_path),
                    image_format=output_format,
                    quality=quality
                )
            
            if success:
                result["status"] = "completed"
                result["output_file"] = output_filename
                if progress_callback:
                    progress_callback.update(file_path.name, "completed", 100)
            else:
                result["status"] = "failed"
                result["error"] = "이미지 변환 실패"
                if progress_callback:
                    progress_callback.update(file_path.name, "failed", 0)
            
        except Exception as e:
            result["status"] = "failed"
            result["error"] = str(e)
            if progress_callback:
                progress_callback.update(file_path.name, "failed", 0)
            logger.error(f"파일 처리 실패 {file_path}: {str(e)}")
        
        finally:
            result["end_time"] = datetime.now()
            result["duration"] = (result["end_time"] - result["start_time"]).total_seconds()
            
            with self.lock:
                self.results.append(result)
        
        return result
    
    async def process_batch(
        self,
        file_paths: List[Path],
        output_format: str = "png",
        quality: int = 100,
        sheet_name: Optional[str] = None,
        range_start: Optional[str] = None,
        range_end: Optional[str] = None,
        type: str = "image",
        progress_callback: Optional[Callable] = None
    ) -> List[Dict[str, Any]]:
        """
        배치 처리 실행
        
        Args:
            file_paths: 처리할 파일 경로 리스트
            output_format: 출력 형식
            quality: 이미지 품질
            sheet_name: 시트 이름
            range_start: 범위 시작
            range_end: 범위 끝
            progress_callback: 진행률 콜백 함수
            
        Returns:
            처리 결과 리스트
        """
        if not file_paths:
            return []
        
        # 진행률 콜백 초기화
        if progress_callback:
            progress_callback = ProgressCallback(len(file_paths))
        
        logger.info(f"배치 처리 시작: {len(file_paths)}개 파일")
        
        try:
            await self.initialize()
            
            # 세마포어로 동시 처리 수 제한
            semaphore = asyncio.Semaphore(self.max_workers)
            
            async def process_with_semaphore(file_path):
                async with semaphore:
                    return await self.process_single_file(
                        file_path,
                        output_format,
                        quality,
                        sheet_name,
                        range_start,
                        range_end,
                        type,
                        progress_callback
                    )
            
            # 모든 작업 생성
            tasks = [
                process_with_semaphore(file_path) 
                for file_path in file_paths
            ]
            
            # 동시 실행
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            # 예외 처리
            processed_results = []
            for i, result in enumerate(results):
                if isinstance(result, Exception):
                    error_result = {
                        "file_path": str(file_paths[i]),
                        "file_name": file_paths[i].name,
                        "status": "failed",
                        "error": str(result),
                        "start_time": datetime.now(),
                        "end_time": datetime.now(),
                        "duration": 0
                    }
                    processed_results.append(error_result)
                else:
                    processed_results.append(result)
            
            return processed_results
            
        finally:
            await self.close()
    
    def get_statistics(self) -> Dict[str, Any]:
        """처리 통계 반환"""
        if not self.results:
            return {
                "total_files": 0,
                "completed_files": 0,
                "failed_files": 0,
                "success_rate": 0,
                "total_duration": 0,
                "average_duration": 0
            }
        
        completed = [r for r in self.results if r["status"] == "completed"]
        failed = [r for r in self.results if r["status"] == "failed"]
        
        total_duration = sum(r.get("duration", 0) for r in self.results)
        average_duration = total_duration / len(self.results) if self.results else 0
        
        return {
            "total_files": len(self.results),
            "completed_files": len(completed),
            "failed_files": len(failed),
            "success_rate": (len(completed) / len(self.results) * 100) if self.results else 0,
            "total_duration": total_duration,
            "average_duration": average_duration
        }

def find_excel_files(directory: str, recursive: bool = True) -> List[Path]:
    """
    디렉토리에서 Excel 파일 찾기
    
    Args:
        directory: 검색할 디렉토리
        recursive: 하위 디렉토리 포함 여부
        
    Returns:
        Excel 파일 경로 리스트
    """
    dir_path = Path(directory)
    if not dir_path.exists():
        return []
    
    pattern = "**/*" if recursive else "*"
    excel_files = []
    
    for file_path in dir_path.glob(pattern):
        if file_path.is_file() and file_path.suffix.lower() in ['.xlsx', '.xls']:
            excel_files.append(file_path)
    
    return sorted(excel_files)

async def batch_convert_excel_files(
    input_paths: List[str],
    output_dir: str = "outputs",
    output_format: str = "png",
    quality: int = 100,
    max_workers: int = 3,
    sheet_name: Optional[str] = None,
    range_start: Optional[str] = None,
    range_end: Optional[str] = None,
    recursive: bool = True,
    type: str = "image"
) -> Dict[str, Any]:
    # Windows에서 asyncio 이벤트 루프 문제 해결
    if platform.system() == 'Windows':
        try:
            # Windows에서 ProactorEventLoop 사용
            if isinstance(asyncio.get_event_loop_policy(), asyncio.WindowsProactorEventLoopPolicy):
                pass  # 이미 ProactorEventLoop 사용 중
            else:
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            logger.info("Windows 이벤트 루프 정책 설정 완료")
        except Exception as e:
            logger.warning(f"Windows 이벤트 루프 정책 설정 실패: {e}")
    """
    Excel 파일 배치 변환
    
    Args:
        input_paths: 입력 파일/디렉토리 경로 리스트
        output_dir: 출력 디렉토리
        output_format: 출력 형식
        quality: 이미지 품질
        max_workers: 최대 동시 처리 수
        sheet_name: 시트 이름
        range_start: 범위 시작
        range_end: 범위 끝
        recursive: 하위 디렉토리 포함 여부
        
    Returns:
        처리 결과 요약
    """
    # 파일 경로 수집
    all_files = []
    for input_path in input_paths:
        path = Path(input_path)
        if path.is_file():
            if path.suffix.lower() in ['.xlsx', '.xls']:
                all_files.append(path)
        elif path.is_dir():
            excel_files = find_excel_files(str(path), recursive)
            all_files.extend(excel_files)
    
    if not all_files:
        logger.warning("처리할 Excel 파일을 찾을 수 없습니다.")
        return {"error": "처리할 Excel 파일을 찾을 수 없습니다."}
    
    logger.info(f"총 {len(all_files)}개의 Excel 파일을 찾았습니다.")
    
    # 배치 처리 실행
    processor = BatchProcessor(max_workers=max_workers, output_dir=output_dir)
    
    try:
        results = await processor.process_batch(
            all_files,
            output_format,
            quality,
            sheet_name,
            range_start,
            range_end,
            type
        )
        
        # 통계 생성
        stats = processor.get_statistics()
        
        logger.info(f"배치 처리 완료: {stats['completed_files']}/{stats['total_files']} 성공")
        
        return {
            "success": True,
            "statistics": stats,
            "results": results
        }
        
    except Exception as e:
        logger.error(f"배치 처리 실패: {str(e)}")
        return {
            "success": False,
            "error": str(e)
        }

if __name__ == "__main__":
    # 테스트 실행
    async def test_batch_processing():
        test_files = ["test2.xlsx"]  # 테스트용 파일
        result = await batch_convert_excel_files(
            test_files,
            output_format="png",
            quality=100,
            max_workers=2
        )
        print("배치 처리 결과:", result)
    
    asyncio.run(test_batch_processing())
