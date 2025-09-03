"""
HTML Renderer Module

Excel 데이터를 HTML 테이블로 변환하는 모듈
jinja2 템플릿 엔진을 사용하여 Excel의 스타일을 CSS로 매핑합니다.
"""

from typing import Dict, List, Any, Optional
from jinja2 import Environment, FileSystemLoader, Template
import os
import logging

logger = logging.getLogger(__name__)


class HTMLRenderer:
    """Excel 데이터를 HTML로 변환하는 클래스"""

    def __init__(self, template_dir: str = "templates"):
        """
        HTMLRenderer 초기화

        Args:
            template_dir (str): 템플릿 디렉토리 경로
        """
        self.template_dir = template_dir
        self.env = Environment(
            loader=FileSystemLoader(template_dir),
            autoescape=True,
            trim_blocks=True,
            lstrip_blocks=True,
        )

    def render_sheet(self, sheet_data: Dict[str, Any]) -> str:
        """
        시트 데이터를 HTML로 렌더링합니다.

        Args:
            sheet_data (Dict[str, Any]): ExcelParser에서 추출한 시트 데이터

        Returns:
            str: 렌더링된 HTML 문자열
        """
        try:
            template = self.env.get_template("sheet.html")

            # CSS 스타일 생성
            css_styles = self._generate_css_styles(sheet_data)

            # 렌더링 컨텍스트 준비
            context = {
                "sheet_data": sheet_data,
                "css_styles": css_styles,
                "table_html": self._generate_table_html(sheet_data),
            }

            return template.render(**context)

        except Exception as e:
            logger.error(f"HTML 렌더링 실패: {str(e)}")
            return self._generate_fallback_html(sheet_data)

    def _generate_css_styles(self, sheet_data: Dict[str, Any]) -> str:
        """
        CSS 스타일을 생성합니다.

        Args:
            sheet_data (Dict[str, Any]): 시트 데이터

        Returns:
            str: CSS 스타일 문자열
        """
        css_rules = []

        # 기본 테이블 스타일
        css_rules.append(
            """
        .excel-table {
            border-collapse: collapse;
            font-family: 'Calibri', 'Arial', sans-serif;
            font-size: 16px;
            width: auto;
            table-layout: auto;
        }
        
        .excel-table td, .excel-table th {
            border: 1px solid #d0d0d0;
            padding: 2px 4px;
            vertical-align: top;
            white-space: normal;
            word-wrap: break-word;
            word-break: break-word;
            min-width: 50px;
            max-width: 400px;
            height: auto;
            min-height: 20px;
        }
        """
        )

        # 셀별 스타일 생성
        for row_idx, row in enumerate(sheet_data.get("cells", [])):
            for col_idx, cell in enumerate(row):
                if not cell:
                    continue

                cell_selector = f".cell-{row_idx}-{col_idx}"
                cell_styles = []

                # 폰트 스타일
                font = cell.get("font", {})
                if font.get("bold"):
                    cell_styles.append("font-weight: bold;")
                if font.get("italic"):
                    cell_styles.append("font-style: italic;")
                if font.get("underline"):
                    cell_styles.append("text-decoration: underline;")
                if font.get("size"):
                    cell_styles.append(f"font-size: {font['size']}px;")
                if font.get("name"):
                    cell_styles.append(f"font-family: '{font['name']}', sans-serif;")
                if font.get("color") and font["color"] is not None:
                    cell_styles.append(f"color: {self._format_color(font['color'])};")

                # 배경색
                fill = cell.get("fill", {})
                if fill.get("color") and fill["color"] is not None:
                    cell_styles.append(
                        f"background-color: {self._format_color(fill['color'])};"
                    )

                # 정렬
                alignment = cell.get("alignment", {})
                if alignment.get("horizontal"):
                    cell_styles.append(f"text-align: {alignment['horizontal']};")
                if alignment.get("vertical"):
                    cell_styles.append(f"vertical-align: {alignment['vertical']};")
                if alignment.get("wrap_text"):
                    cell_styles.append("white-space: normal; word-wrap: break-word;")
                else:
                    # 기본적으로 텍스트 줄바꿈 허용
                    cell_styles.append("white-space: normal; word-wrap: break-word;")

                # 테두리
                border = cell.get("border", {})
                border_styles = self._generate_border_css(border)
                if border_styles:
                    cell_styles.extend(border_styles)

                # 스타일이 있으면 CSS 규칙 추가
                if cell_styles:
                    css_rules.append(f"{cell_selector} {{ {' '.join(cell_styles)} }}")

        # 병합 셀 스타일
        for merged_cell in sheet_data.get("merged_cells", []):
            start_row = merged_cell["start_row"] - sheet_data["dimensions"]["start_row"]
            start_col = merged_cell["start_col"] - sheet_data["dimensions"]["start_col"]
            end_row = merged_cell["end_row"] - sheet_data["dimensions"]["start_row"]
            end_col = merged_cell["end_col"] - sheet_data["dimensions"]["start_col"]

            css_rules.append(
                f"""
            .cell-{start_row}-{start_col} {{
                grid-column: {start_col + 1} / {end_col + 2};
                grid-row: {start_row + 1} / {end_row + 2};
            }}
            """
            )

        return "\n".join(css_rules)

    def _generate_border_css(self, border: Dict[str, Any]) -> List[str]:
        """
        테두리 CSS를 생성합니다.

        Args:
            border (Dict[str, Any]): 테두리 정보

        Returns:
            List[str]: CSS 스타일 리스트
        """
        border_styles = []

        for side, border_info in border.items():
            if border_info and border_info.get("style"):
                style = border_info["style"]
                color = (
                    self._format_color(border_info.get("color"))
                    if border_info.get("color")
                    else "#000000"
                )

                if side == "left":
                    border_styles.append(f"border-left: 1px {style} {color};")
                elif side == "right":
                    border_styles.append(f"border-right: 1px {style} {color};")
                elif side == "top":
                    border_styles.append(f"border-top: 1px {style} {color};")
                elif side == "bottom":
                    border_styles.append(f"border-bottom: 1px {style} {color};")

        return border_styles

    def _format_color(self, color) -> str:
        """
        색상을 CSS 형식으로 변환합니다.

        Args:
            color: 색상 값 (문자열, RGB 객체, 또는 None)

        Returns:
            str: CSS 색상 값
        """
        if not color:
            return "#000000"

        try:
            # 문자열로 변환
            color_str = str(color)

            # RGB 형식 (예: FF0000)
            if color_str.startswith("FF"):
                return f"#{color_str[2:]}"
            elif len(color_str) == 6 and color_str.isalnum():
                return f"#{color_str}"
            elif color_str.startswith("theme_"):
                # 테마 색상은 기본 색상으로 대체
                return "#000000"
            elif color_str.startswith("indexed_"):
                # 인덱스 색상은 기본 색상으로 대체
                return "#000000"
            elif color_str.startswith("RGB"):
                # RGB 객체인 경우 기본 색상으로 대체
                return "#000000"

            return "#000000"
        except Exception:
            # 예외 발생 시 기본 색상 반환
            return "#000000"


    def _estimate_text_width(self, text: str) -> int:
        """문자열 길이를 픽셀 단위로 추정"""
        if not text:
            return 80

        # 개행이 있는 경우 가장 긴 줄만을 기준으로 계산
        if "\n" in text:
            lines = text.split("\n")
            max_line_width = 0
            
            for line in lines:
                line = line.strip()  # 앞뒤 공백 제거
                if not line:  # 빈 줄은 건너뛰기
                    continue
                    
                # 각 줄의 너비 계산
                line_width = self._calculate_line_width(line)
                max_line_width = max(max_line_width, line_width)
            
            return max(max_line_width, 80)
        else:
            # 개행이 없는 경우 기존 방식으로 계산
            return self._calculate_line_width(text)
    
    def _calculate_line_width(self, line: str) -> int:
        """한 줄의 텍스트 너비를 계산"""
        if not line:
            return 80
            
        text_length = len(line)
        korean_chars = sum(1 for c in line if '\u3131' <= c <= '\u318E' or '\uAC00' <= c <= '\uD7A3')
        english_chars = sum(1 for c in line if c.isascii() and c.isalpha())
        number_chars = sum(1 for c in line if c.isdigit())
        space_chars = sum(1 for c in line if c.isspace())
        special_chars = text_length - korean_chars - english_chars - number_chars - space_chars

        estimated_width = (
            korean_chars * 12 +
            english_chars * 8 +
            number_chars * 8 +
            space_chars * 4 +
            special_chars * 6
        )
        estimated_width += 20  # padding

        # 세로 배치 텍스트(한 글자씩 줄바꿈) 최소 보정
        if text_length > 5 and all(len(w) == 1 for w in line):
            estimated_width = max(estimated_width, 120)

        return max(estimated_width, 120)
        

    def _compute_column_widths(self, sheet_data: Dict[str, Any]) -> Dict[int, int]:
        """모든 셀을 스캔해서 열별 최대 폭 계산"""
        column_widths = {}

        for row in sheet_data.get("cells", []):
            for col_idx, cell in enumerate(row):
                if not cell:
                    continue
                text = str(cell.get("value", "") or "")
                width = self._estimate_text_width(text)

                if col_idx not in column_widths:
                    column_widths[col_idx] = width
                else:
                    column_widths[col_idx] = max(column_widths[col_idx], width)

        # 최소 120px 이상 보장
        for k in column_widths:
            column_widths[k] = max(column_widths[k], 120)

        return column_widths


    def _generate_table_html(self, sheet_data: Dict[str, Any]) -> str:
        """
        테이블 HTML을 생성합니다.

        Args:
            sheet_data (Dict[str, Any]): 시트 데이터

        Returns:
            str: 테이블 HTML 문자열
        """
        html_parts = ['<table class="excel-table">']

        # 행 높이와 열 너비 설정
        row_heights = sheet_data.get("row_heights", {})
        column_widths = self._compute_column_widths(sheet_data)

        for row_idx, row in enumerate(sheet_data.get("cells", [])):
            # 행 높이 설정
            row_height = row_heights.get(
                row_idx + sheet_data["dimensions"]["start_row"], None
            )
            row_style = f' style="height: {row_height}px;"' if row_height else ""

            html_parts.append(f"<tr{row_style}>")

            for col_idx, cell in enumerate(row):
                if not cell:
                    continue

                # 셀 클래스
                cell_class = f"cell-{row_idx}-{col_idx}"

                # 병합 셀 처리
                if cell.get("is_merged"):
                    # 병합된 셀 중 첫 번째 셀만 내용 표시
                    merged_range = cell.get("merge_range", "")
                    if merged_range and ":" in merged_range:
                        try:
                            start_cell, end_cell = merged_range.split(":")
                            if cell["address"] == start_cell:
                                # 병합 범위 계산
                                start_row = (
                                    cell["row"] - sheet_data["dimensions"]["start_row"]
                                )
                                start_col = (
                                    cell["col"] - sheet_data["dimensions"]["start_col"]
                                )

                                # 끝 셀 주소 파싱 (간단한 방법)
                                end_col_letter = "".join(filter(str.isalpha, end_cell))
                                end_row_num = int(
                                    "".join(filter(str.isdigit, end_cell))
                                )
                                end_row = (
                                    end_row_num - sheet_data["dimensions"]["start_row"]
                                )

                                # 열 주소를 숫자로 변환 (A=1, B=2, ..., Z=26, AA=27, ...)
                                end_col = 0
                                for i, char in enumerate(end_col_letter):
                                    end_col += (ord(char) - ord("A") + 1) * (
                                        26 ** (len(end_col_letter) - i - 1)
                                    )
                                end_col = (
                                    end_col - sheet_data["dimensions"]["start_col"]
                                )

                                colspan = end_col - start_col + 1
                                rowspan = end_row - start_row + 1

                                # 병합된 셀의 너비 계산 (병합된 셀의 내용에 맞춤)
                                cell_value = cell.get('value', '')
                                if cell_value is not None:
                                    text = str(cell_value)
                                    estimated_width = self._estimate_text_width(text)
                                    # 병합된 셀은 내용에 맞는 적절한 너비 설정
                                    merged_cell_width = max(estimated_width, 120)
                                    cell_style = f' style="min-width: {merged_cell_width}px; max-width: {merged_cell_width * 2}px;"'
                                else:
                                    cell_style = ' style="min-width: 120px;"'

                                html_parts.append(
                                    f'<td class="{cell_class}"{cell_style} '
                                    f'colspan="{colspan}" rowspan="{rowspan}">'
                                    f"{self._format_cell_value(cell)}</td>"
                                )
                            else:
                                # 병합된 셀은 건너뛰기
                                continue
                        except Exception as e:
                            # 병합 셀 처리 실패 시 일반 셀로 처리
                            cell_value = cell.get('value', '')
                            if cell_value is not None:
                                text = str(cell_value)
                                estimated_width = self._estimate_text_width(text)
                                cell_style = f' style="min-width: {estimated_width}px; max-width: {estimated_width * 2}px;"'
                            else:
                                cell_style = ' style="min-width: 120px;"'
                            
                            html_parts.append(
                                f'<td class="{cell_class}"{cell_style}>'
                                f"{self._format_cell_value(cell)}</td>"
                            )
                    else:
                        # 병합 셀 정보가 없는 경우 일반 셀로 처리
                        cell_value = cell.get('value', '')
                        if cell_value is not None:
                            text = str(cell_value)
                            estimated_width = self._estimate_text_width(text)
                            cell_style = f' style="min-width: {estimated_width}px; max-width: {estimated_width * 2}px;"'
                        else:
                            cell_style = ' style="min-width: 120px;"'
                        
                        html_parts.append(
                            f'<td class="{cell_class}"{cell_style}>'
                            f"{self._format_cell_value(cell)}</td>"
                        )
                else:
                    # 일반 셀 처리 - 각 셀의 데이터 길이에 맞는 독립적인 너비 설정
                    cell_value = cell.get('value', '')
                    if cell_value is not None:
                        text = str(cell_value)
                        estimated_width = self._estimate_text_width(text)
                        cell_style = f' style="min-width: {estimated_width}px; max-width: {estimated_width * 2}px;"'
                    else:
                        cell_style = ' style="min-width: 120px;"'
                    
                    html_parts.append(
                        f'<td class="{cell_class}"{cell_style}>'
                        f"{self._format_cell_value(cell)}</td>"
                    )

            html_parts.append("</tr>")

        html_parts.append("</table>")
        table_html = "\n".join(html_parts)

        return table_html

    def _save_html_file(self, table_html: str, sheet_data: Dict[str, Any]):
        """
        생성된 HTML을 파일로 저장합니다.
        
        Args:
            table_html (str): 테이블 HTML 문자열
            sheet_data (Dict[str, Any]): 시트 데이터
        """
        try:
            import os
            from pathlib import Path
            from datetime import datetime
            
            # outputs 디렉토리 생성
            outputs_dir = Path("outputs")
            outputs_dir.mkdir(exist_ok=True)
            
            # 파일명 생성 (시트명과 타임스탬프 포함)
            sheet_name = sheet_data.get("sheet_name", "unknown")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{sheet_name}_{timestamp}.html"
            file_path = outputs_dir / filename
            
            # 테이블 너비 계산
            total_width = self._calculate_table_width(sheet_data)
            
            # 완전한 HTML 문서 생성
            full_html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Sheet - {sheet_name}</title>
    <style>
        body {{
            font-family: 'Calibri', 'Arial', sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
            width: auto;
            min-width: auto;
            max-width: none;
            overflow-x: auto;
        }}
        .excel-table {{
            border-collapse: collapse;
            font-family: 'Calibri', 'Arial', sans-serif;
            font-size: 16px;
            width: auto;
            table-layout: auto;
            background-color: #ffffff;
        }}
        
        .excel-table td, .excel-table th {{
            border: 1px solid #d0d0d0;
            padding: 8px 10px;
            vertical-align: top;
            white-space: pre-wrap;
            word-wrap: break-word;
            word-break: break-all;
            overflow-wrap: break-word;
            min-width: 80px;
            max-width: none;
            min-height: 30px;
            height: auto;
            box-sizing: border-box;
            font-size: 16px;
            line-height: 1.4;
            text-overflow: ellipsis;
            overflow: visible;
        }}
        
        .excel-table th {{
            background-color: #f8f9fa;
            font-weight: bold;
            text-align: center;
        }}
        
        .excel-table tr:nth-child(even) {{
            background-color: #fafafa;
        }}
        
        .excel-table tr:hover {{
            background-color: #f0f8ff;
        }}
        
        /* 긴 텍스트 셀 스타일 */
        .excel-table .long-text {{
            max-height: none;
            overflow: visible;
            white-space: pre-wrap;
            word-wrap: break-word;
            word-break: break-all;
            overflow-wrap: break-word;
            min-height: 50px;
            height: auto;
        }}
        
        .excel-table .text-cell {{
            text-align: left;
            white-space: pre-wrap;
            word-wrap: break-word;
            word-break: break-all;
            overflow-wrap: break-word;
            min-height: 45px;
            height: auto;
            overflow: visible;
        }}
        
        /* 긴 텍스트 내용 스타일 */
        .long-text-content {{
            display: inline-block;
            word-wrap: break-word;
            word-break: break-word;
            overflow-wrap: break-word;
            white-space: pre-wrap;
            max-width: 100%;
        }}
        
        /* 수식 셀 스타일 */
        .excel-table .formula-cell {{
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 14px;
            background-color: #f8f9fa;
            border-left: 3px solid #007acc;
        }}
        
        /* 스크롤바 스타일 */
        .excel-table td::-webkit-scrollbar,
        .excel-table th::-webkit-scrollbar {{
            width: 6px;
            height: 6px;
        }}
        
        .excel-table td::-webkit-scrollbar-track,
        .excel-table th::-webkit-scrollbar-track {{
            background: #f1f1f1;
            border-radius: 3px;
        }}
        
        .excel-table td::-webkit-scrollbar-thumb,
        .excel-table th::-webkit-scrollbar-thumb {{
            background: #c1c1c1;
            border-radius: 3px;
        }}
        
        .excel-table td::-webkit-scrollbar-thumb:hover,
        .excel-table th::-webkit-scrollbar-thumb:hover {{
            background: #a8a8a8;
        }}
        
        /* 고해상도 디스플레이 대응 */
        @media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {{
            .excel-table td, .excel-table th {{
                padding: 8px 10px;
                min-height: 35px;
            }}
        }}
        
        /* 인쇄 스타일 */
        @media print {{
            body {{
                background-color: white;
                margin: 0;
                width: auto;
                min-width: auto;
                max-width: none;
            }}
            .excel-table {{
                page-break-inside: auto;
                width: auto;
            }}
            .excel-table tr {{
                page-break-inside: avoid;
                page-break-after: auto;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        {table_html}
    </div>
</body>
</html>"""

            # 파일 저장
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(full_html)

            logger.info(f"HTML 파일 저장 완료: {file_path}")

        except Exception as e:
            logger.warning(f"HTML 파일 저장 실패: {str(e)}")
            # HTML 파일 저장 실패는 전체 변환에 영향을 주지 않도록 경고만 출력
    
    def _calculate_table_width(self, sheet_data: Dict[str, Any]) -> int:
        """
        테이블의 총 너비를 계산합니다.
        
        Args:
            sheet_data (Dict[str, Any]): 시트 데이터
            
        Returns:
            int: 테이블 총 너비 (픽셀)
        """
        try:
            total_width = 0
            column_widths = sheet_data.get("column_widths", {})
            dimensions = sheet_data.get("dimensions", {})
            
            # 각 열의 너비를 합산
            for col_idx in range(dimensions.get("columns", 0)):
                col_num = col_idx + dimensions.get("start_col", 1)
                col_width = column_widths.get(col_num, 100)  # 기본 너비 100px
                
                # 열 너비가 픽셀 단위가 아닌 경우 변환
                if col_width and isinstance(col_width, (int, float)):
                    # Excel 열 너비를 픽셀로 변환 (대략적인 변환)
                    pixel_width = int(col_width * 7)  # Excel 열 너비 1 = 약 7픽셀
                    pixel_width = max(pixel_width, 80)  # 최소 너비 80px
                    total_width += pixel_width
                else:
                    total_width += 100  # 기본 너비
            
            # 테이블이 비어있는 경우 기본 너비 설정
            if total_width == 0:
                total_width = 800
            
            # 여백과 패딩 추가
            total_width += 40  # 좌우 여백
            
            logger.info(f"테이블 총 너비 계산: {total_width}px")
            return total_width
            
        except Exception as e:
            logger.warning(f"테이블 너비 계산 실패: {str(e)}, 기본값 800px 사용")
            return 800

    def _format_cell_value(self, cell: Dict[str, Any]) -> str:
        """
        셀 값을 HTML 형식으로 변환합니다.

        Args:
            cell (Dict[str, Any]): 셀 데이터

        Returns:
            str: 포맷된 셀 값
        """
        value = cell.get("value", "")

        if value is None:
            return ""
        elif isinstance(value, (int, float)):
            # 숫자 포맷 처리
            number_format = cell.get("number_format", "")
            if number_format:
                try:
                    # 간단한 포맷 처리 (실제로는 더 복잡한 로직 필요)
                    if "0.00" in number_format:
                        return f"{value:.2f}"
                    elif "0%" in number_format:
                        return f"{value:.0%}"
                except:
                    pass
            return str(value)
        else:
            return str(value)

    def _generate_fallback_html(self, sheet_data: Dict[str, Any]) -> str:
        """
        오류 발생 시 기본 HTML을 생성합니다.

        Args:
            sheet_data (Dict[str, Any]): 시트 데이터

        Returns:
            str: 기본 HTML 문자열
        """
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Excel Sheet - {sheet_data.get('sheet_name', 'Unknown')}</title>
            <style>
                table {{ border-collapse: collapse; width: 100%; }}
                td, th {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
            </style>
        </head>
        <body>
            <h2>Excel Sheet: {sheet_data.get('sheet_name', 'Unknown')}</h2>
            <p>Error occurred during rendering. Showing basic table.</p>
            <table>
                {self._generate_table_html(sheet_data)}
            </table>
        </body>
        </html>
        """


def render_excel_to_html(sheet_data: Dict[str, Any]) -> str:
    """
    Excel 데이터를 HTML로 변환하는 편의 함수

    Args:
        sheet_data (Dict[str, Any]): ExcelParser에서 추출한 시트 데이터

    Returns:
        str: 렌더링된 HTML 문자열
    """
    renderer = HTMLRenderer()
    return renderer.render_sheet(sheet_data)
