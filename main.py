"""
Excel to Image Converter - Main Application

Excel 파일을 이미지로 변환하는 메인 애플리케이션
전체 변환 파이프라인을 관리하고 CLI 인터페이스를 제공합니다.
"""

import os
import sys
import argparse
import logging
import subprocess
from typing import Optional, Dict, Any
from pathlib import Path

# 프로젝트 모듈 import
from excel_parser import ExcelParser, parse_excel_file
from html_renderer import HTMLRenderer, render_excel_to_html
from image_converter import ImageConverter, convert_html_to_image_sync
from batch_processor import batch_convert_excel_files, find_excel_files

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('excel2image.log', encoding='utf-8')
    ]
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


class ExcelToImageConverter:
    """Excel을 이미지로 변환하는 메인 클래스"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        ExcelToImageConverter 초기화
        
        Args:
            config (Dict[str, Any], optional): 설정 딕셔너리
        """
        self.config = config or {}
        self.parser = None
        self.renderer = None
        self.converter = None
        
        # 기본 설정
        self.default_config = {
            'output_format': 'png',
            'image_quality': 100,  # 최고품질로 변경
            'output_dir': 'outputs',
            'template_dir': 'templates',
            'headless': True,
            'width': None,
            'height': None,
            'max_workers': 3  # 배치 처리용
        }
        
        # 사용자 설정으로 기본값 업데이트
        self.default_config.update(self.config)
        self.config = self.default_config
        
    def convert_excel_to_image(self, 
                              excel_file: str,
                              output_file: Optional[str] = None,
                              sheet_name: Optional[str] = None,
                              range_start: Optional[str] = None,
                              range_end: Optional[str] = None,
                              **kwargs) -> bool:
        """
        Excel 파일을 이미지로 변환합니다.
        
        Args:
            excel_file (str): Excel 파일 경로
            output_file (str, optional): 출력 이미지 파일 경로
            sheet_name (str, optional): 시트 이름 (None이면 모든 시트 변환)
            range_start (str, optional): 시작 범위 (예: 'A1')
            range_end (str, optional): 끝 범위 (예: 'D10')
            **kwargs: 추가 옵션들
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            logger.info(f"Excel 파일 변환 시작: {excel_file}")
            
            # 설정 업데이트
            config = self.config.copy()
            config.update(kwargs)
            
            # 시트 이름이 지정되지 않으면 모든 시트 변환
            if sheet_name is None:
                return self._convert_all_sheets(excel_file, output_file, range_start, range_end, config)
            else:
                # 단일 시트 변환
                return self._convert_single_sheet(excel_file, output_file, sheet_name, range_start, range_end, config)
                
        except Exception as e:
            logger.error(f"변환 중 오류 발생: {str(e)}")
            return False
    
    def _parse_excel(self, excel_file: str, sheet_name: Optional[str], 
                    range_start: Optional[str], range_end: Optional[str]) -> Dict[str, Any]:
        """
        Excel 파일을 파싱합니다.
        
        Args:
            excel_file (str): Excel 파일 경로
            sheet_name (str, optional): 시트 이름
            range_start (str, optional): 시작 범위
            range_end (str, optional): 끝 범위
            
        Returns:
            Dict[str, Any]: 파싱된 시트 데이터
        """
        try:
            return parse_excel_file(excel_file, sheet_name, range_start, range_end)
        except Exception as e:
            logger.error(f"Excel 파싱 오류: {str(e)}")
            return {}
    
    def _render_html(self, sheet_data: Dict[str, Any], template_dir: str) -> str:
        """
        시트 데이터를 HTML로 렌더링합니다.
        
        Args:
            sheet_data (Dict[str, Any]): 시트 데이터
            template_dir (str): 템플릿 디렉토리
            
        Returns:
            str: 렌더링된 HTML
        """
        try:
            return render_excel_to_html(sheet_data)
        except Exception as e:
            logger.error(f"HTML 렌더링 오류: {str(e)}")
            return ""
    
    def _convert_to_image(self, html_content: str, output_file: str, 
                         image_format: str, quality: int, 
                         width: Optional[int], height: Optional[int]) -> bool:
        """
        HTML을 이미지로 변환합니다.
        
        Args:
            html_content (str): HTML 내용
            output_file (str): 출력 파일 경로
            image_format (str): 이미지 형식
            quality (int): 이미지 품질
            width (int, optional): 강제 너비
            height (int, optional): 강제 높이
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            return convert_html_to_image_sync(
                html_content, output_file, image_format, quality, width, height
            )
        except Exception as e:
            logger.error(f"이미지 변환 오류: {str(e)}")
            return False
    
    def _generate_output_path(self, excel_file: str, sheet_name: Optional[str], 
                            output_format: str) -> str:
        """
        출력 파일 경로를 생성합니다.
        
        Args:
            excel_file (str): Excel 파일 경로
            sheet_name (str, optional): 시트 이름
            output_format (str): 출력 형식
            
        Returns:
            str: 출력 파일 경로
        """
        try:
            # Excel 파일명에서 확장자 제거
            excel_path = Path(excel_file)
            base_name = excel_path.stem
            
            # 시트 이름이 있으면 추가
            if sheet_name:
                output_name = f"{base_name}_{sheet_name}.{output_format}"
            else:
                output_name = f"{base_name}.{output_format}"
            
            # 출력 디렉토리 생성 및 검증
            output_dir = Path(self.config['output_dir'])
            
            # 출력 디렉토리 경로 검증
            if not str(output_dir) or str(output_dir).strip() == '':
                raise ValueError("출력 디렉토리 경로가 비어있습니다.")
            
            # 디렉토리 생성
            try:
                output_dir.mkdir(exist_ok=True)
                logger.info(f"출력 디렉토리 확인/생성: {output_dir}")
            except Exception as e:
                logger.error(f"출력 디렉토리 생성 실패: {output_dir}, 오류: {str(e)}")
                raise ValueError(f"출력 디렉토리를 생성할 수 없습니다: {output_dir}")
            
            # 최종 출력 경로 생성 및 검증
            output_path = output_dir / output_name
            output_path_str = str(output_path)
            
            if not output_path_str or output_path_str.strip() == '':
                raise ValueError("출력 파일 경로가 비어있습니다.")
            
            logger.info(f"출력 경로 생성: {output_path_str}")
            return output_path_str
            
        except Exception as e:
            logger.error(f"출력 경로 생성 실패: {str(e)}")
            raise ValueError(f"출력 경로 생성 실패: {str(e)}")
    
    def _convert_all_sheets(self, excel_file: str, output_file: Optional[str], 
                           range_start: Optional[str], range_end: Optional[str], 
                           config: Dict[str, Any]) -> bool:
        """
        모든 시트를 각각의 이미지로 변환합니다.
        
        Args:
            excel_file (str): Excel 파일 경로
            output_file (str, optional): 출력 이미지 파일 경로 (기본 파일명에 시트명 추가)
            range_start (str, optional): 시작 범위
            range_end (str, optional): 끝 범위
            config (Dict[str, Any]): 설정
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            # Excel 파일에서 모든 시트 이름 가져오기
            parser = ExcelParser(excel_file)
            if not parser.load_workbook():
                logger.error("Excel 파일을 로드할 수 없습니다.")
                return False
            
            sheet_names = parser.get_sheet_names()
            parser.close()
            
            if not sheet_names:
                logger.error("Excel 파일에 시트가 없습니다.")
                return False
            
            logger.info(f"총 {len(sheet_names)}개 시트를 변환합니다: {', '.join(sheet_names)}")
            
            success_count = 0
            total_count = len(sheet_names)
            
            for sheet_name in sheet_names:
                try:
                    logger.info(f"시트 변환 중: {sheet_name}")
                    
                    # 각 시트별 출력 파일명 생성
                    output_type = config.get('type', 'image')
                    
                    if output_file:
                        # 출력 파일명이 지정된 경우 시트명을 추가
                        base_name = os.path.splitext(output_file)[0]
                        if output_type.lower() == 'html':
                            ext = '.html'
                        else:
                            ext = os.path.splitext(output_file)[1]
                        sheet_output_file = f"{base_name}_{sheet_name}{ext}"
                    else:
                        # 출력 파일명이 지정되지 않은 경우 기본 파일명에 시트명 추가
                        if output_type.lower() == 'html':
                            sheet_output_file = self._generate_output_path(excel_file, sheet_name, 'html')
                        else:
                            sheet_output_file = self._generate_output_path(excel_file, sheet_name, config['output_format'])
                    
                    # 단일 시트 변환
                    success = self._convert_single_sheet(excel_file, sheet_output_file, sheet_name, range_start, range_end, config)
                    
                    if success:
                        success_count += 1
                        logger.info(f"시트 '{sheet_name}' 변환 완료: {sheet_output_file}")
                    else:
                        logger.error(f"시트 '{sheet_name}' 변환 실패")
                        
                except Exception as e:
                    logger.error(f"시트 '{sheet_name}' 변환 중 오류: {str(e)}")
            
            logger.info(f"전체 시트 변환 완료: {success_count}/{total_count} 성공")
            return success_count > 0
            
        except Exception as e:
            logger.error(f"모든 시트 변환 중 오류: {str(e)}")
            return False
    
    def _convert_single_sheet(self, excel_file: str, output_file: str, sheet_name: str,
                             range_start: Optional[str], range_end: Optional[str], 
                             config: Dict[str, Any]) -> bool:
        """
        단일 시트를 이미지로 변환합니다.
        
        Args:
            excel_file (str): Excel 파일 경로
            output_file (str): 출력 이미지 파일 경로
            sheet_name (str): 시트 이름
            range_start (str, optional): 시작 범위
            range_end (str, optional): 끝 범위
            config (Dict[str, Any]): 설정
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            # 1단계: Excel 파싱
            logger.info(f"1단계: Excel 파일 파싱 중... (시트: {sheet_name})")
            sheet_data = self._parse_excel(excel_file, sheet_name, range_start, range_end)
            if not sheet_data:
                logger.error("Excel 파싱 실패")
                return False
            
            # 2단계: HTML 변환
            logger.info("2단계: HTML 변환 중...")
            html_content = self._render_html(sheet_data, config.get('template_dir'))
            if not html_content:
                logger.error("HTML 변환 실패")
                return False
            
            # type 파라미터에 따라 처리
            output_type = config.get('type', 'image')
            
            if output_type.lower() == 'html':
                # HTML만 생성
                logger.info("3단계: HTML 파일 저장 중...")
                try:
                    # HTML 파일로 저장
                    html_output_file = output_file.replace(f".{config['output_format']}", ".html")
                    with open(html_output_file, "w", encoding="utf-8") as f:
                        f.write(html_content)
                    logger.info(f"HTML 파일 생성 완료: {html_output_file}")
                    return True
                except Exception as e:
                    logger.error(f"HTML 파일 생성 실패: {str(e)}")
                    return False
            else:
                # 이미지 변환 (기본값)
                logger.info("3단계: 이미지 변환 중...")
                success = self._convert_to_image(
                    html_content, 
                    output_file, 
                    config['output_format'],
                    config['image_quality'],
                    config.get('width'),
                    config.get('height')
                )
            
            if success:
                logger.info(f"변환 완료: {output_file}")
                return True
            else:
                logger.error("이미지 변환 실패")
                return False
                
        except Exception as e:
            logger.error(f"단일 시트 변환 중 오류: {str(e)}")
            return False
    
    def batch_convert(self, excel_files: list, **kwargs) -> Dict[str, bool]:
        """
        여러 Excel 파일을 배치로 변환합니다.
        
        Args:
            excel_files (list): Excel 파일 경로 리스트
            **kwargs: 변환 옵션들
            
        Returns:
            Dict[str, bool]: 파일별 변환 결과
        """
        results = {}
        
        logger.info(f"배치 변환 시작: {len(excel_files)}개 파일")
        
        for i, excel_file in enumerate(excel_files, 1):
            logger.info(f"처리 중 ({i}/{len(excel_files)}): {excel_file}")
            
            try:
                success = self.convert_excel_to_image(excel_file, **kwargs)
                results[excel_file] = success
                
                if success:
                    logger.info(f"✓ 성공: {excel_file}")
                else:
                    logger.error(f"✗ 실패: {excel_file}")
                    
            except Exception as e:
                logger.error(f"✗ 오류: {excel_file} - {str(e)}")
                results[excel_file] = False
        
        # 결과 요약
        success_count = sum(1 for success in results.values() if success)
        logger.info(f"배치 변환 완료: {success_count}/{len(excel_files)} 성공")
        
        return results


def main():
    """메인 함수 - CLI 인터페이스"""
    parser = argparse.ArgumentParser(
        description='Excel 파일을 이미지로 변환하는 도구',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python main.py input.xlsx
  python main.py input.xlsx -o output.png
  python main.py input.xlsx -s "Sheet1" -r "A1:D10"
  python main.py input.xlsx -f jpeg --quality 90
  python main.py *.xlsx --batch
        """
    )
    
    # 필수 인수
    parser.add_argument('input', help='Excel 파일 경로 (또는 와일드카드 패턴)')
    
    # 출력 옵션
    parser.add_argument('-o', '--output', help='출력 이미지 파일 경로')
    parser.add_argument('-d', '--output-dir', default='outputs', help='출력 디렉토리 (기본값: outputs)')
    parser.add_argument('-f', '--format', choices=['png', 'jpeg', 'jpg'], default='png', 
                       help='출력 이미지 형식 (기본값: png)')
    parser.add_argument('--quality', type=int, default=100, 
                       help='JPEG 품질 (1-100, 기본값: 95)')
    parser.add_argument('--type', choices=['image', 'html'], default='image',
                       help='출력 타입 (image 또는 html, 기본값: image)')
    
    # Excel 옵션
    parser.add_argument('-s', '--sheet', help='시트 이름')
    parser.add_argument('-i', '--sheet-index', type=int, help='시트 인덱스 (0부터 시작)')
    parser.add_argument('-r', '--range', help='셀 범위 (예: A1:D10)')
    
    # 이미지 옵션
    parser.add_argument('--width', type=int, help='강제 이미지 너비')
    parser.add_argument('--height', type=int, help='강제 이미지 높이')
    
    # 배치 처리
    parser.add_argument('--batch', action='store_true', help='배치 처리 모드')
    parser.add_argument('--recursive', action='store_true', help='하위 디렉토리 포함')
    
    # 기타 옵션
    parser.add_argument('--headless', action='store_true', default=True, 
                       help='헤드리스 모드 (기본값: True)')
    parser.add_argument('--verbose', '-v', action='store_true', help='상세 로그 출력')
    parser.add_argument('--quiet', '-q', action='store_true', help='로그 출력 최소화')
    
    args = parser.parse_args()
    
    # 로깅 레벨 설정
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    elif args.quiet:
        logging.getLogger().setLevel(logging.WARNING)
    
    # 범위 파싱
    range_start = None
    range_end = None
    if args.range:
        try:
            if ':' in args.range:
                range_start, range_end = args.range.split(':')
            else:
                range_start = range_end = args.range
        except ValueError:
            logger.error("잘못된 범위 형식입니다. 예: A1:D10")
            return 1
    
    # 설정 구성
    config = {
        'output_format': args.format,
        'image_quality': args.quality,
        'output_dir': args.output_dir,
        'headless': args.headless,
        'width': args.width,
        'height': args.height,
        'type': args.type
    }
    
    # Playwright 브라우저 설치 확인
    try:
        check_and_install_playwright_browsers()
    except Exception as e:
        logger.error(f"Playwright 브라우저 설정 실패: {e}")
        return 1
    
    # 변환기 초기화
    converter = ExcelToImageConverter(config)
    
    try:
        # 입력 파일 처리
        if args.batch or '*' in args.input or '?' in args.input:
            # 배치 처리 (최적화된 버전)
            import glob
            excel_files = glob.glob(args.input, recursive=args.recursive)
            excel_files = [f for f in excel_files if f.endswith(('.xlsx', '.xls'))]
            
            if not excel_files:
                logger.error("변환할 Excel 파일을 찾을 수 없습니다.")
                return 1
            
            logger.info(f"배치 변환 시작: {len(excel_files)}개 파일")
            
            # 비동기 배치 처리 실행
            import asyncio
            result = asyncio.run(batch_convert_excel_files(
                input_paths=excel_files,
                output_dir=args.output_dir,
                output_format=args.format,
                quality=args.quality,
                max_workers=config.get('max_workers', 3),
                sheet_name=args.sheet,
                range_start=range_start,
                range_end=range_end,
                recursive=args.recursive,
                type=args.type
            ))
            
            if result.get('success'):
                stats = result['statistics']
                print(f"\n=== 배치 변환 완료 ===")
                print(f"총 파일: {stats['total_files']}개")
                print(f"성공: {stats['completed_files']}개")
                print(f"실패: {stats['failed_files']}개")
                print(f"성공률: {stats['success_rate']:.1f}%")
                print(f"총 소요시간: {stats['total_duration']:.1f}초")
                print(f"평균 처리시간: {stats['average_duration']:.1f}초")
                
                if stats['failed_files'] > 0:
                    print("\n실패한 파일:")
                    for file_result in result['results']:
                        if file_result['status'] == 'failed':
                            print(f"  - {file_result['file_name']}: {file_result['error']}")
            else:
                print(f"배치 처리 실패: {result.get('error', '알 수 없는 오류')}")
                return 1
            
        else:
            # 단일 파일 처리
            if not os.path.exists(args.input):
                logger.error(f"파일을 찾을 수 없습니다: {args.input}")
                return 1
            
            success = converter.convert_excel_to_image(
                args.input,
                output_file=args.output,
                sheet_name=args.sheet,
                range_start=range_start,
                range_end=range_end
            )
            
            if success:
                print("변환 완료!")
                return 0
            else:
                print("변환 실패!")
                return 1
                
    except KeyboardInterrupt:
        logger.info("사용자에 의해 중단되었습니다.")
        return 1
    except Exception as e:
        logger.error(f"예상치 못한 오류: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
