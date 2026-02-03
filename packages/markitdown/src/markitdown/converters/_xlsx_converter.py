import sys
from typing import BinaryIO, Any
from ._html_converter import HtmlConverter
from .._base_converter import DocumentConverter, DocumentConverterResult
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE
from .._stream_info import StreamInfo

# == Add by Masahiro (2026/01/28) ==
from openpyxl import load_workbook
from pathlib import Path
import os
from spire.xls import Workbook, XlsShape
import re

def safe_name(name: str) -> str:
    # Windowsで使えない文字を置換
    return re.sub(r'[\\/:*?"<>|]', "_", name)

def _export_charts_via_excel_com(xlsx_path: Path, sheet_name: str, media_dir: Path) -> list[str]:

    xlsx_path = Path(xlsx_path).resolve()
    # print('xlsx_path:', xlsx_path)
    media_dir = Path(media_dir).resolve()
    media_dir.mkdir(exist_ok=True)

    md_lines: list[str] = []

    try:
        workbook = Workbook()
        workbook.LoadFromFile(str(xlsx_path))
        for i in range(workbook.Worksheets.Count):
            worksheet = workbook.Worksheets.get_Item(i)
            ws_name = worksheet.Name  # ← いま処理中のシート名

            if ws_name != sheet_name:
                continue

            for j in range(worksheet.Charts.Count):
                print('チャート保存中')
                fname = f"{xlsx_path.stem}__{safe_name(ws_name)}__chart{j}.png"
                out_path = media_dir / fname

                chartImage = workbook.SaveChartAsImage(worksheet, j)
                chartImage.Save(str(out_path))

                md_lines.append(f"![]({media_dir.name}/{fname})")
    except Exception as e:
        print(e)

    return md_lines



# チャートシートPNG化関数
def _export_chart_sheets_via_excel_com(xlsx_path: Path, sheet_name: str, media_dir: Path) -> list[str]:
    
    xlsx_path = Path(xlsx_path).resolve()
    media_dir = Path(media_dir).resolve()
    media_dir.mkdir(exist_ok=True)

    md_lines: list[str] = []

    try:
        workbook = Workbook()
        workbook.LoadFromFile(str(xlsx_path))
        for i in range(workbook.Worksheets.Count):
            # ワークシートを取得
            worksheet = workbook.Worksheets.get_Item(i)
            ws_name = worksheet.Name

            if ws_name != sheet_name:
                continue

            # ワークシート内の図形を繰り返し処理
            for j in range(worksheet.PrstGeomShapes.Count):
                print('図形保存中')
                # 図形を取得
                shape = worksheet.PrstGeomShapes.get_Item(j)
                fname = f"{xlsx_path.stem}__{safe_name(ws_name)}__Shape{j}.png"
                out_path = media_dir / fname

                # 図形をXlsShapeオブジェクトに変換
                xlsShape = XlsShape(shape)
                # 図形を画像ストリームとして保存
                image = xlsShape.SaveToImage()
                # 画像ストリームをファイルに保存
                image.Save(str(out_path))

                md_lines.append(f"![]({media_dir.name}/{fname})")
    except Exception as e:
        print(e)

    return md_lines
# == Add by Masahiro (2026/01/28) ==


# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_xlsx_dependency_exc_info = None
try:
    import pandas as pd
    import openpyxl  # noqa: F401
except ImportError:
    _xlsx_dependency_exc_info = sys.exc_info()

_xls_dependency_exc_info = None
try:
    import pandas as pd  # noqa: F811
    import xlrd  # noqa: F401
except ImportError:
    _xls_dependency_exc_info = sys.exc_info()

ACCEPTED_XLSX_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
]
ACCEPTED_XLSX_FILE_EXTENSIONS = [".xlsx"]

ACCEPTED_XLS_MIME_TYPE_PREFIXES = [
    "application/vnd.ms-excel",
    "application/excel",
]
ACCEPTED_XLS_FILE_EXTENSIONS = [".xls"]


class XlsxConverter(DocumentConverter):
    """
    Converts XLSX files to Markdown, with each sheet presented as a separate Markdown table.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_XLSX_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_XLSX_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        

        import os

        print("CWD:", os.getcwd())
        print("stream_info.local_path:", stream_info.local_path)
        print("resolved:", str(Path(stream_info.local_path).resolve()))

        p = Path(stream_info.local_path).resolve()
        print("exists:", p.exists())

        # Check the dependencies
        if _xlsx_dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".xlsx",
                    feature="xlsx",
                )
            ) from _xlsx_dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _xlsx_dependency_exc_info[2]
            )

        sheets = pd.read_excel(file_stream, sheet_name=None, engine="openpyxl")
        print('file_stream:', file_stream)
        print('stream_info;', stream_info)


        # == Add by Masahiro (2026/01/28) ==

        print('stream_info.local_path:', stream_info.local_path)
        xlsx_path = Path(stream_info.local_path).resolve()
        media_dir = xlsx_path.parent / "media"
        media_dir.mkdir(exist_ok=True)

        file_stream.seek(0)
        wb_img = load_workbook(file_stream, data_only=True)
        
        all_md_content = ""

        def proc_md_content(md_content):
            md_content += f"## {s}\n"
            # html_content = sheets[s].to_html(index=False)
            df = sheets[s].replace(r"^\s*$", pd.NA, regex=True)  # 空文字も空扱いにする（不要なら削除OK）
            df = df.dropna(how="all")                            # NaNのみの行を削除
            html_content = df.to_html(index=False)

            md_content += (
                self._html_converter.convert_string(
                    html_content, **kwargs
                ).markdown.strip()
                + "\n\n"
            )

            ws = wb_img[s]

            # シート内画像の保存
            images = getattr(ws, "_images", [])
            if images:
                md_content += "### Images\n"
                for i, img in enumerate(images, start=1):
                    try:
                        data = img._data()
                    except Exception:
                        continue

                    fname = f"{xlsx_path.stem}__{s}__img{i}.png"
                    (media_dir / fname).write_bytes(data)
                    md_content += f"![]({media_dir.name}/{fname})\n\n"

            chart_md = _export_charts_via_excel_com(xlsx_path, s, media_dir)
            if chart_md:
                md_content += "### Charts\n" + "\n\n".join(chart_md) + "\n\n"


            chart_sheet_md = _export_chart_sheets_via_excel_com(xlsx_path, s, media_dir)
            if chart_sheet_md:
                md_content += "## Chart Sheets\n\n" + "\n".join(chart_sheet_md) + "\n"
            
            return md_content

        for s in sheets:
            md_content = ""
            md_content = proc_md_content(md_content)
            all_md_content = proc_md_content(all_md_content)

            md_path = xlsx_path.with_name(f"{xlsx_path.stem}_{safe_name(s)}.md")
            md_path.write_text(md_content, encoding="utf-8")
        
        # == Add by Masahiro (2026/01/28) ==

        return DocumentConverterResult(markdown=all_md_content.strip())


class XlsConverter(DocumentConverter):
    """
    Converts XLS files to Markdown, with each sheet presented as a separate Markdown table.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_XLS_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_XLS_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Load the dependencies
        if _xls_dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".xls",
                    feature="xls",
                )
            ) from _xls_dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _xls_dependency_exc_info[2]
            )

        sheets = pd.read_excel(file_stream, sheet_name=None, engine="xlrd")
        # md_content = ""

        # == Add by Masahiro (2026/01/28) ==

        print('stream_info.local_path:', stream_info.local_path)
        xlsx_path = Path(stream_info.local_path).resolve()
        media_dir = xlsx_path.parent / "media"
        media_dir.mkdir(exist_ok=True)

        file_stream.seek(0)
        wb_img = load_workbook(file_stream, data_only=True)

        # == Add by Masahiro (2026/01/28) ==

        all_md_content = ""

        def proc_md_content(md_content):
            md_content += f"## {s}\n"
            # html_content = sheets[s].to_html(index=False)
            df = sheets[s].replace(r"^\s*$", pd.NA, regex=True)  # 空文字も空扱いにする（不要なら削除OK）
            df = df.dropna(how="all")                            # NaNのみの行を削除
            html_content = df.to_html(index=False)

            md_content += (
                self._html_converter.convert_string(
                    html_content, **kwargs
                ).markdown.strip()
                + "\n\n"
            )

            ws = wb_img[s]

            # シート内画像の保存
            images = getattr(ws, "_images", [])
            if images:
                md_content += "### Images\n"
                for i, img in enumerate(images, start=1):
                    try:
                        data = img._data()
                    except Exception:
                        continue

                    fname = f"{xlsx_path.stem}__{s}__img{i}.png"
                    (media_dir / fname).write_bytes(data)
                    md_content += f"![]({media_dir.name}/{fname})\n\n"

            chart_md = _export_charts_via_excel_com(xlsx_path, s, media_dir)
            if chart_md:
                md_content += "### Charts\n" + "\n\n".join(chart_md) + "\n\n"


            chart_sheet_md = _export_chart_sheets_via_excel_com(xlsx_path, s, media_dir)
            if chart_sheet_md:
                md_content += "## Chart Sheets\n\n" + "\n".join(chart_sheet_md) + "\n"
            
            return md_content

        for s in sheets:
            md_content = ""
            md_content = proc_md_content(md_content)
            all_md_content = proc_md_content(all_md_content)

            md_path = xlsx_path.with_name(f"{xlsx_path.stem}_{safe_name(s)}.md")
            md_path.write_text(md_content, encoding="utf-8")
            # == Add by Masahiro (2026/01/28) ==

        return DocumentConverterResult(markdown=md_content.strip())
