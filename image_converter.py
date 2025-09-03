"""
Image Converter Module

HTML을 이미지로 변환하는 모듈
Playwright를 사용하여 HTML을 렌더링하고 스크린샷을 캡처합니다.
"""

import asyncio
import io
import os
import tempfile
from typing import Optional, Tuple, Dict, Any
from playwright.sync_api import sync_playwright, Browser, Page
from PIL import Image
import logging

logger = logging.getLogger(__name__)


class ImageConverter:
    """HTML을 이미지로 변환하는 클래스"""
    
    def __init__(self, headless: bool = True):
        """
        ImageConverter 초기화
        
        Args:
            headless (bool): 헤드리스 모드 여부
        """
        self.headless = headless
        self.browser = None
        self.page = None
        self.playwright = None
        
    def initialize(self):
        """Playwright 브라우저를 초기화합니다."""
        try:
            # 이미 초기화되어 있으면 스킵
            if self.page and self.browser and self.playwright:
                logger.info("Playwright 브라우저가 이미 초기화되어 있습니다.")
                return
            
            # Windows에서 subprocess 문제 해결을 위한 환경 변수 설정
            import platform
            if platform.system() == 'Windows':
                import os
                os.environ['PYTHONPATH'] = os.pathsep.join([
                    os.environ.get('PYTHONPATH', ''),
                    os.path.dirname(os.path.abspath(__file__))
                ])
                # Windows에서 subprocess 실행을 위한 추가 환경 변수
                os.environ['PYTHONUNBUFFERED'] = '1'
                os.environ['PYTHONIOENCODING'] = 'utf-8'
                
            self.playwright = sync_playwright().start()
            
            # Windows와 Linux 호환성을 위한 브라우저 옵션 (간소화)
            browser_args = [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--no-first-run',
                '--disable-gpu',
                '--disable-background-timer-throttling',
                '--disable-backgrounding-occluded-windows',
                '--disable-renderer-backgrounding',
                '--disable-features=TranslateUI',
                '--disable-ipc-flooding-protection',
                '--disable-default-apps',
                '--disable-extensions',
                '--disable-plugins',
                '--disable-sync',
                '--disable-translate',
                '--hide-scrollbars',
                '--mute-audio',
                '--no-default-browser-check',
                '--no-pings',
                '--disable-web-security',
                '--allow-running-insecure-content'
            ]
            
            # Windows 특정 옵션 (간소화)
            if platform.system() == 'Windows':
                browser_args.extend([
                    '--disable-gpu-sandbox',
                    '--disable-software-rasterizer',
                    '--disable-gpu-process'
                ])
            
            # Windows와 Linux 호환성을 위한 브라우저 실행 옵션
            browser_options = {
                'headless': self.headless,
                'args': browser_args
            }
            
            # Windows에서 추가 옵션 (이미지 품질 향상)
            if platform.system() == 'Windows':
                browser_options.update({
                    'executable_path': None,  # 시스템 기본 경로 사용
                    'ignore_default_args': ['--disable-extensions'],
                    'chromium_sandbox': False  # 샌드박스 비활성화로 안정성 향상
                })
            
            # 브라우저 실행
            self.browser = self.playwright.chromium.launch(**browser_options)
            
            # 컨텍스트 생성 (이미지 품질 향상)
            context_options = {
                'viewport': {
                    'width': 1920,
                    'height': 1080
                },
                'device_scale_factor': 2,  # 고해상도 렌더링
                'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
            
            # Windows에서 추가 컨텍스트 옵션
            if platform.system() == 'Windows':
                context_options.update({
                    'locale': 'ko-KR',
                    'timezone_id': 'Asia/Seoul'
                })
            
            context = self.browser.new_context(**context_options)
            
            # 페이지 생성
            self.page = context.new_page()
            
            # 페이지 설정 (이미지 품질 향상)
            self.page.set_extra_http_headers({
                'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8'
            })
            
            # 이미지 품질 향상을 위한 추가 설정
            self.page.add_init_script("""
                // 폰트 렌더링 품질 향상
                document.documentElement.style.webkitFontSmoothing = 'antialiased';
                document.documentElement.style.mozOsxFontSmoothing = 'grayscale';
                document.documentElement.style.textRendering = 'optimizeLegibility';
                
                // 이미지 렌더링 품질 향상
                const style = document.createElement('style');
                style.textContent = `
                    * {
                        image-rendering: -webkit-optimize-contrast;
                        image-rendering: -webkit-crisp-edges;
                        image-rendering: -moz-crisp-edges;
                        image-rendering: crisp-edges;
                        text-rendering: optimizeLegibility;
                        -webkit-font-smoothing: antialiased;
                        -moz-osx-font-smoothing: grayscale;
                    }
                `;
                document.head.appendChild(style);
            """)
            
            logger.info("Playwright 브라우저 초기화 완료")
            
        except Exception as e:
            logger.error(f"Playwright 초기화 실패: {str(e)}")
            # 초기화 실패 시 리소스 정리
            self._cleanup_resources()
            raise
    
    def convert_html_to_image(self, html_content: str, output_path: str,
                            image_format: str = 'png', quality: int = 95,
                            width: Optional[int] = None, height: Optional[int] = None) -> bool:
        """
        HTML을 이미지로 변환합니다.
        
        Args:
            html_content (str): HTML 내용
            output_path (str): 출력 이미지 경로
            image_format (str): 이미지 형식 ('png' 또는 'jpeg')
            quality (int): 이미지 품질 (1-100, JPEG에만 적용)
            width (int, optional): 강제 너비
            height (int, optional): 강제 높이
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            # 입력 매개변수 검증
            if not html_content or html_content.strip() == '':
                logger.error("HTML 내용이 비어있습니다.")
                return False
            
            if not output_path or output_path.strip() == '':
                logger.error("출력 경로가 비어있습니다.")
                return False
            
            # 이미지 형식 검증
            if image_format.lower() not in ['png', 'jpeg', 'jpg']:
                logger.error(f"지원하지 않는 이미지 형식: {image_format}")
                return False
            
            # 품질 값 검증
            if not isinstance(quality, int) or quality < 1 or quality > 100:
                logger.error(f"유효하지 않은 품질 값: {quality} (1-100 범위여야 함)")
                return False
            
            # 브라우저가 초기화되지 않았으면 초기화
            if not self.page or not self.browser:
                self.initialize()
            
            # 임시 HTML 파일 생성
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            try:
                # HTML 파일 로드 (Windows 호환성 개선 - 간소화)
                try:
                    self.page.goto(f'file://{temp_html_path}', wait_until='domcontentloaded', timeout=30000)
                except Exception as e:
                    logger.warning(f"domcontentloaded 대기 실패, 기본 대기로 대체: {e}")
                    self.page.goto(f'file://{temp_html_path}', timeout=30000)
                
                # 페이지 로드 후 잠시 대기 (Windows 안정성)
                import time
                time.sleep(0.5)
                
                # 페이지 크기 조정
                self._adjust_page_size(width, height)
                
                # 스크린샷 캡처 (이미지 품질 향상)
                screenshot_options = {
                    'type': image_format,
                    'full_page': True,
                    'omit_background': False,  # 배경 포함으로 선명도 향상
                    'timeout': 30000,
                    'scale': 'css',  # CSS 스케일 사용으로 선명도 향상
                    'quality': 100 if image_format == 'jpeg' else None  # 최고 품질
                }
                
                # JPEG인 경우에만 quality 옵션 추가
                if image_format == 'jpeg':
                    screenshot_options['quality'] = quality
                
                # Windows에서 추가 안정성을 위한 재시도 로직 (이미지 품질 향상)
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        # 페이지가 완전히 로드될 때까지 대기
                        self.page.wait_for_load_state('networkidle', timeout=10000)
                        
                        # 추가 렌더링 대기 (이미지 품질 향상)
                        import time
                        time.sleep(1.0)
                        
                        screenshot_bytes = self.page.screenshot(**screenshot_options)
                        break
                    except Exception as e:
                        if attempt == max_retries - 1:
                            raise e
                        logger.warning(f"스크린샷 캡처 재시도 {attempt + 1}/{max_retries}: {e}")
                        time.sleep(1.0)
                
                # 이미지 후처리 및 저장
                self._process_and_save_image(screenshot_bytes, output_path, image_format, quality)
                
                logger.info(f"이미지 변환 완료: {output_path}")
                return True
                
            finally:
                # 임시 파일 삭제
                os.unlink(temp_html_path)
                
        except Exception as e:
            logger.error(f"이미지 변환 실패: {str(e)}")
            # 변환 실패 시 리소스 정리
            self._cleanup_resources()
            return False
    
    def _adjust_page_size(self, width: Optional[int] = None, height: Optional[int] = None):
        """
        페이지 크기를 조정합니다.
        
        Args:
            width (int, optional): 강제 너비
            height (int, optional): 강제 높이
        """
        try:
            # 페이지 내용의 실제 크기 측정 (Windows 호환성 개선 - 간소화)
            try:
                content_size = self.page.evaluate("""
                () => {
                    const body = document.body;
                    const html = document.documentElement;
                    
                    return {
                        width: Math.max(
                            body.scrollWidth,
                            body.offsetWidth,
                            html.clientWidth,
                            html.scrollWidth,
                            html.offsetWidth
                        ),
                        height: Math.max(
                            body.scrollHeight,
                            body.offsetHeight,
                            html.clientHeight,
                            html.scrollHeight,
                            html.offsetHeight
                        )
                    };
                }
                """)
            except Exception as e:
                logger.warning(f"페이지 크기 측정 실패, 기본값 사용: {e}")
                content_size = {'width': 800, 'height': 600}

            logger.info(f"페이지 크기 측정: {content_size}")
            
            # 사용자 지정 크기 또는 측정된 크기 사용
            final_width = width or content_size['width']
            final_height = height or content_size['height']
            
            # 여백 추가 (Windows 호환성을 위해 조정 - 간소화)
            margin = 40  # 여백을 더 줄여서 균등하게
            final_width += margin * 2
            final_height += margin * 2
            
            # 최소 크기 보장 (Windows 호환성을 위해 조정 - 간소화)
            final_width = max(final_width, 800)  # 최소 너비 조정
            final_height = max(final_height, 600)  # 최소 높이 조정
            
            # 높이에 따른 단계별 글자 크기 조정
            if final_height >= 20000:
                # 20000px 이상: 가장 큰 글자 크기
                try:
                    self.page.evaluate("""
                    () => {
                        const style = document.createElement('style');
                        style.textContent = `
                            .excel-table {
                                font-size: 68px !important;
                                line-height: 1.6 !important;
                            }
                            .excel-table td, .excel-table th {
                                font-size: 28px !important;
                                line-height: 1.6 !important;
                                padding: 12px 14px !important;
                                min-height: 45px !important;
                            }
                            .excel-table .long-text {
                                font-size: 26px !important;
                                line-height: 1.5 !important;
                            }
                            .excel-table .text-cell {
                                font-size: 27px !important;
                                line-height: 1.5 !important;
                            }
                            .excel-table .formula-cell {
                                font-size: 25px !important;
                                line-height: 1.5 !important;
                            }
                        `;
                        document.head.appendChild(style);
                    }
                    """)
                    logger.info(f"높이가 {final_height}px로 20000 이상이므로 글자 크기를 28px로 조정")
                except Exception as e:
                    logger.warning(f"글자 크기 조정 실패: {e}")
            elif final_height >= 15000:
                # 15000px 이상: 큰 글자 크기
                try:
                    self.page.evaluate("""
                    () => {
                        const style = document.createElement('style');
                        style.textContent = `
                            .excel-table {
                                font-size: 24px !important;
                                line-height: 1.55 !important;
                            }
                            .excel-table td, .excel-table th {
                                font-size: 24px !important;
                                line-height: 1.55 !important;
                                padding: 10px 12px !important;
                                min-height: 40px !important;
                            }
                            .excel-table .long-text {
                                font-size: 22px !important;
                                line-height: 1.45 !important;
                            }
                            .excel-table .text-cell {
                                font-size: 23px !important;
                                line-height: 1.45 !important;
                            }
                            .excel-table .formula-cell {
                                font-size: 21px !important;
                                line-height: 1.45 !important;
                            }
                        `;
                        document.head.appendChild(style);
                    }
                    """)
                    logger.info(f"높이가 {final_height}px로 15000 이상이므로 글자 크기를 24px로 조정")
                except Exception as e:
                    logger.warning(f"글자 크기 조정 실패: {e}")
            elif final_height >= 10000:
                # 10000px 이상: 중간 글자 크기
                try:
                    self.page.evaluate("""
                    () => {
                        const style = document.createElement('style');
                        style.textContent = `
                            .excel-table {
                                font-size: 22px !important;
                                line-height: 1.5 !important;
                            }
                            .excel-table td, .excel-table th {
                                font-size: 22px !important;
                                line-height: 1.5 !important;
                                padding: 9px 11px !important;
                                min-height: 38px !important;
                            }
                            .excel-table .long-text {
                                font-size: 20px !important;
                                line-height: 1.4 !important;
                            }
                            .excel-table .text-cell {
                                font-size: 21px !important;
                                line-height: 1.4 !important;
                            }
                            .excel-table .formula-cell {
                                font-size: 19px !important;
                                line-height: 1.4 !important;
                            }
                        `;
                        document.head.appendChild(style);
                    }
                    """)
                    logger.info(f"높이가 {final_height}px로 10000 이상이므로 글자 크기를 22px로 조정")
                except Exception as e:
                    logger.warning(f"글자 크기 조정 실패: {e}")
            elif final_height >= 5000:
                # 5000px 이상: 기본 큰 글자 크기
                try:
                    self.page.evaluate("""
                    () => {
                        const style = document.createElement('style');
                        style.textContent = `
                            .excel-table {
                                font-size: 20px !important;
                                line-height: 1.5 !important;
                            }
                            .excel-table td, .excel-table th {
                                font-size: 20px !important;
                                line-height: 1.5 !important;
                                padding: 8px 10px !important;
                                min-height: 35px !important;
                            }
                            .excel-table .long-text {
                                font-size: 18px !important;
                                line-height: 1.4 !important;
                            }
                            .excel-table .text-cell {
                                font-size: 19px !important;
                                line-height: 1.4 !important;
                            }
                            .excel-table .formula-cell {
                                font-size: 18px !important;
                                line-height: 1.4 !important;
                            }
                        `;
                        document.head.appendChild(style);
                    }
                    """)
                    logger.info(f"높이가 {final_height}px로 5000 이상이므로 글자 크기를 20px로 조정")
                except Exception as e:
                    logger.warning(f"글자 크기 조정 실패: {e}")
            
            # viewport 크기 설정 (Windows 호환성 개선 - 간소화)
            try:
                self.page.set_viewport_size({
                    "width": final_width,
                    "height": final_height
                })
            except Exception as e:
                logger.warning(f"viewport 크기 설정 실패, 기본값 사용: {e}")
                # Windows에서 viewport 설정이 실패하면 기본값 사용
                self.page.set_viewport_size({
                    "width": 1200,
                    "height": 800
                })
            
            logger.info(f"페이지 크기 조정: {final_width}x{final_height}")
            
        except Exception as e:
            logger.error(f"페이지 크기 조정 실패: {str(e)}")
    
    def _process_and_save_image(self, screenshot_bytes: bytes, output_path: str,
                              image_format: str, quality: int):
        """
        스크린샷을 후처리하고 저장합니다.
        
        Args:
            screenshot_bytes (bytes): 스크린샷 바이트 데이터
            output_path (str): 출력 경로
            image_format (str): 이미지 형식
            quality (int): 이미지 품질
        """
        try:
            # 경로 검증 및 정규화
            if not output_path or output_path.strip() == '':
                raise ValueError("출력 경로가 비어있습니다.")
            
            # 경로를 절대 경로로 정규화
            output_path = os.path.abspath(output_path.strip())
            
            # 디렉토리 경로 추출 및 검증
            output_dir = os.path.dirname(output_path)
            if not output_dir:
                raise ValueError(f"유효하지 않은 출력 경로: {output_path}")
            
            # 디렉토리가 존재하지 않으면 생성
            try:
                os.makedirs(output_dir, exist_ok=True)
                logger.info(f"출력 디렉토리 생성/확인: {output_dir}")
            except Exception as e:
                logger.error(f"출력 디렉토리 생성 실패: {output_dir}, 오류: {str(e)}")
                raise ValueError(f"출력 디렉토리를 생성할 수 없습니다: {output_dir}")
            
            # 파일명 검증
            filename = os.path.basename(output_path)
            if not filename or filename.strip() == '':
                raise ValueError(f"유효하지 않은 파일명: {output_path}")
            
            # PIL Image로 변환
            image = Image.open(io.BytesIO(screenshot_bytes))
            
            # 이미지 품질 향상을 위한 후처리
            if image.mode in ('RGBA', 'LA', 'P'):
                # 투명도가 있는 이미지는 RGB로 변환
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            
            # 이미지 저장 (최고품질 설정)
            try:
                if image_format.lower() == 'jpeg':
                    # JPEG 저장 (최고품질)
                    image.save(output_path, 'JPEG', 
                              quality=quality, 
                              optimize=True, 
                              progressive=True,  # 프로그레시브 JPEG
                              subsampling=0)  # 서브샘플링 비활성화 (최고품질)
                else:
                    # PNG 저장 (최고품질)
                    image.save(output_path, 'PNG', 
                              optimize=True, 
                              compress_level=0)  # 압축 레벨 0 (최고품질)
                
                logger.info(f"이미지 저장 완료: {output_path}")
                
            except PermissionError as e:
                logger.error(f"파일 저장 권한 오류: {output_path}, 오류: {str(e)}")
                raise ValueError(f"파일 저장 권한이 없습니다: {output_path}")
            except OSError as e:
                logger.error(f"파일 시스템 오류: {output_path}, 오류: {str(e)}")
                raise ValueError(f"파일 시스템 오류: {str(e)}")
            
        except ValueError as e:
            logger.error(f"경로 검증 실패: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"이미지 저장 실패: {str(e)}")
            raise
    
    def convert_html_file_to_image(self, html_file_path: str, output_path: str,
                                 image_format: str = 'png', quality: int = 95,
                                 width: Optional[int] = None, height: Optional[int] = None) -> bool:
        """
        HTML 파일을 이미지로 변환합니다.
        
        Args:
            html_file_path (str): HTML 파일 경로
            output_path (str): 출력 이미지 경로
            image_format (str): 이미지 형식
            quality (int): 이미지 품질
            width (int, optional): 강제 너비
            height (int, optional): 강제 높이
            
        Returns:
            bool: 변환 성공 여부
        """
        try:
            # 입력 매개변수 검증
            if not html_file_path or html_file_path.strip() == '':
                logger.error("HTML 파일 경로가 비어있습니다.")
                return False
            
            if not output_path or output_path.strip() == '':
                logger.error("출력 경로가 비어있습니다.")
                return False
            
            # HTML 파일 존재 여부 확인
            if not os.path.exists(html_file_path):
                logger.error(f"HTML 파일이 존재하지 않습니다: {html_file_path}")
                return False
            
            # 이미지 형식 검증
            if image_format.lower() not in ['png', 'jpeg', 'jpg']:
                logger.error(f"지원하지 않는 이미지 형식: {image_format}")
                return False
            
            # 품질 값 검증
            if not isinstance(quality, int) or quality < 1 or quality > 100:
                logger.error(f"유효하지 않은 품질 값: {quality} (1-100 범위여야 함)")
                return False
            
            if not self.page:
                self.initialize()
            
            # HTML 파일 로드
            self.page.goto(f'file://{os.path.abspath(html_file_path)}', wait_until='networkidle')
            
            # 페이지 크기 조정
            self._adjust_page_size(width, height)
            
            # 스크린샷 캡처
            screenshot_bytes = self.page.screenshot(
                type=image_format,
                quality=quality if image_format == 'jpeg' else None,
                full_page=True,
                omit_background=True
            )
            
            # 이미지 후처리 및 저장
            self._process_and_save_image(screenshot_bytes, output_path, image_format, quality)
            
            logger.info(f"HTML 파일 이미지 변환 완료: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"HTML 파일 이미지 변환 실패: {str(e)}")
            return False
    
    def get_page_dimensions(self, html_content: str) -> Tuple[int, int]:
        """
        HTML 페이지의 차원을 측정합니다.
        
        Args:
            html_content (str): HTML 내용
            
        Returns:
            Tuple[int, int]: (width, height)
        """
        try:
            if not self.page:
                self.initialize()
            
            # 임시 HTML 파일 생성
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            try:
                # HTML 파일 로드
                self.page.goto(f'file://{temp_html_path}', wait_until='networkidle')
                
                # 페이지 크기 측정
                dimensions = self.page.evaluate("""
                    () => {
                        const body = document.body;
                        const html = document.documentElement;
                        return {
                            width: Math.max(
                                body.scrollWidth,
                                body.offsetWidth,
                                html.clientWidth,
                                html.scrollWidth,
                                html.offsetWidth
                            ),
                            height: Math.max(
                                body.scrollHeight,
                                body.offsetHeight,
                                html.clientHeight,
                                html.scrollHeight,
                                html.offsetHeight
                            )
                        };
                    }
                """)
                
                return dimensions['width'], dimensions['height']
                
            finally:
                # 임시 파일 삭제
                os.unlink(temp_html_path)
                
        except Exception as e:
            logger.error(f"페이지 차원 측정 실패: {str(e)}")
            return 800, 600  # 기본값
    
    def _cleanup_resources(self):
        """리소스를 정리합니다."""
        try:
            if self.page:
                try:
                    self.page.close()
                except Exception as e:
                    logger.warning(f"페이지 닫기 실패: {e}")
                finally:
                    self.page = None
                    
            if self.browser:
                try:
                    self.browser.close()
                except Exception as e:
                    logger.warning(f"브라우저 닫기 실패: {e}")
                finally:
                    self.browser = None
                    
            if self.playwright:
                try:
                    self.playwright.stop()
                except Exception as e:
                    logger.warning(f"Playwright 정지 실패: {e}")
                finally:
                    self.playwright = None
                    
        except Exception as e:
            logger.error(f"리소스 정리 실패: {str(e)}")
    
    def close(self):
        """브라우저를 닫습니다."""
        try:
            self._cleanup_resources()
            logger.info("Playwright 브라우저를 닫았습니다.")
        except Exception as e:
            logger.error(f"브라우저 종료 중 오류: {e}")


# 비동기 래퍼 함수들 (API 서버용)
async def convert_html_to_image_async(html_content: str, output_path: str,
                                    image_format: str = 'png', quality: int = 95,
                                    width: Optional[int] = None, height: Optional[int] = None) -> bool:
    """
    HTML을 이미지로 변환하는 비동기 함수 (API 서버용)
    
    Args:
        html_content (str): HTML 내용
        output_path (str): 출력 이미지 경로
        image_format (str): 이미지 형식
        quality (int): 이미지 품질
        width (int, optional): 강제 너비
        height (int, optional): 강제 높이
        
    Returns:
        bool: 변환 성공 여부
    """
    def _convert():
        converter = None
        try:
            converter = ImageConverter()
            return converter.convert_html_to_image(
                html_content, output_path, image_format, quality, width, height
            )
        except Exception as e:
            logger.error(f"동기 변환 실패: {e}")
            return False
        finally:
            if converter:
                try:
                    converter.close()
                except Exception as e:
                    logger.warning(f"동기 변환 후 정리 실패: {e}")
    
    try:
        # ThreadPoolExecutor를 사용하여 동기 함수를 비동기로 실행
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, _convert)
    except Exception as e:
        logger.error(f"비동기 변환 실행 실패: {e}")
        return False


async def convert_html_file_to_image_async(html_file_path: str, output_path: str,
                                         image_format: str = 'png', quality: int = 95,
                                         width: Optional[int] = None, height: Optional[int] = None) -> bool:
    """
    HTML 파일을 이미지로 변환하는 비동기 함수 (API 서버용)
    
    Args:
        html_file_path (str): HTML 파일 경로
        output_path (str): 출력 이미지 경로
        image_format (str): 이미지 형식
        quality (int): 이미지 품질
        width (int, optional): 강제 너비
        height (int, optional): 강제 높이
        
    Returns:
        bool: 변환 성공 여부
    """
    def _convert():
        converter = None
        try:
            converter = ImageConverter()
            return converter.convert_html_file_to_image(
                html_file_path, output_path, image_format, quality, width, height
            )
        except Exception as e:
            logger.error(f"동기 파일 변환 실패: {e}")
            return False
        finally:
            if converter:
                try:
                    converter.close()
                except Exception as e:
                    logger.warning(f"동기 파일 변환 후 정리 실패: {e}")
    
    try:
        # ThreadPoolExecutor를 사용하여 동기 함수를 비동기로 실행
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(None, _convert)
    except Exception as e:
        logger.error(f"비동기 파일 변환 실행 실패: {e}")
        return False


# 동기 래퍼 함수들 (CLI 모드용 - 기존 호환성 유지)
def convert_html_to_image_sync(html_content: str, output_path: str,
                              image_format: str = 'png', quality: int = 95,
                              width: Optional[int] = None, height: Optional[int] = None) -> bool:
    """
    HTML을 이미지로 변환하는 동기 함수 (CLI 모드용)
    
    Args:
        html_content (str): HTML 내용
        output_path (str): 출력 이미지 경로
        image_format (str): 이미지 형식
        quality (int): 이미지 품질
        width (int, optional): 강제 너비
        height (int, optional): 강제 높이
        
    Returns:
        bool: 변환 성공 여부
    """
    converter = None
    try:
        converter = ImageConverter()
        return converter.convert_html_to_image(
            html_content, output_path, image_format, quality, width, height
        )
    except Exception as e:
        logger.error(f"동기 변환 실패: {e}")
        return False
    finally:
        if converter:
            try:
                converter.close()
            except Exception as e:
                logger.warning(f"동기 변환 후 정리 실패: {e}")


def convert_html_file_to_image_sync(html_file_path: str, output_path: str,
                                   image_format: str = 'png', quality: int = 95,
                                   width: Optional[int] = None, height: Optional[int] = None) -> bool:
    """
    HTML 파일을 이미지로 변환하는 동기 함수 (CLI 모드용)
    
    Args:
        html_file_path (str): HTML 파일 경로
        output_path (str): 출력 이미지 경로
        image_format (str): 이미지 형식
        quality (int): 이미지 품질
        width (int, optional): 강제 너비
        height (int, optional): 강제 높이
        
    Returns:
        bool: 변환 성공 여부
    """
    converter = None
    try:
        converter = ImageConverter()
        return converter.convert_html_file_to_image(
            html_file_path, output_path, image_format, quality, width, height
        )
    except Exception as e:
        logger.error(f"동기 파일 변환 실패: {e}")
        return False
    finally:
        if converter:
            try:
                converter.close()
            except Exception as e:
                logger.warning(f"동기 파일 변환 후 정리 실패: {e}")