"""A small Windows GUI for converting WPS Spreadsheets (.et) to Excel."""

from __future__ import annotations

import shutil
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EXCEL_FORMATS = {".xlsx": 51, ".xls": 56}


class ConversionError(RuntimeError):
    pass


def convert_with_wps(source: Path, destination: Path) -> None:
    """Use WPS's COM API, which preserves ET workbook features best."""
    try:
        import win32com.client  # type: ignore[import-not-found]
        import pythoncom  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ConversionError("未安装 pywin32，无法使用 WPS 自动化。") from exc

    # This function runs in a worker thread. COM must be initialized per thread.
    pythoncom.CoInitialize()
    app = None
    book = None
    try:
        app = win32com.client.DispatchEx("et.Application")
        app.Visible = False
        app.DisplayAlerts = False
        book = app.Workbooks.Open(str(source.resolve()))
        book.SaveAs(str(destination.resolve()), FileFormat=EXCEL_FORMATS[destination.suffix.lower()])
    except Exception as exc:  # COM errors vary by installed WPS version
        raise ConversionError(f"WPS 转换失败：{exc}") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def find_soffice() -> str | None:
    """Find LibreOffice without requiring it to be on PATH."""
    candidates = [
        shutil.which("soffice"),
        shutil.which("soffice.exe"),
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
    return next((item for item in candidates if item and Path(item).exists()), None)


def convert_with_libreoffice(source: Path, destination: Path) -> None:
    soffice = find_soffice()
    if not soffice:
        raise ConversionError("找不到 WPS 或 LibreOffice。请安装 WPS 表格，或安装 LibreOffice 并加入 PATH。")

    # LibreOffice writes to an output directory rather than an exact filename.
    output_dir = destination.parent
    # These are LibreOffice Calc's canonical filter names (not its UI labels).
    filter_name = "Calc MS Excel 2007 XML" if destination.suffix.lower() == ".xlsx" else "MS Excel 97"
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", f"{destination.suffix[1:]}:{filter_name}",
         "--outdir", str(output_dir), str(source)],
        capture_output=True,
        text=True,
        timeout=120,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    generated = output_dir / f"{source.stem}{destination.suffix}"
    if result.returncode != 0 or not generated.exists():
        detail = (result.stderr or result.stdout).strip()
        raise ConversionError(f"LibreOffice 转换失败。{detail}")
    if generated.resolve() != destination.resolve():
        if destination.exists():
            destination.unlink()
        generated.replace(destination)


def convert(source: Path, destination: Path) -> str:
    try:
        convert_with_wps(source, destination)
        return "已使用 WPS 完成转换"
    except ConversionError as wps_error:
        try:
            convert_with_libreoffice(source, destination)
            return "已使用 LibreOffice 完成转换"
        except ConversionError as libre_error:
            raise ConversionError(f"{wps_error}\n\n备用方案也失败：{libre_error}") from libre_error


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ET 转 Excel")
        self.geometry("620x255")
        self.resizable(False, False)
        self.source = tk.StringVar()
        self.format = tk.StringVar(value=".xlsx")
        self.status = tk.StringVar(value="请选择要转换的 .et 文件")
        self._build()

    def _build(self) -> None:
        root = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="ET 文件：").grid(row=0, column=0, sticky="w", pady=(0, 16))
        ttk.Entry(root, textvariable=self.source, state="readonly").grid(row=0, column=1, sticky="ew", pady=(0, 16))
        ttk.Button(root, text="选择文件", command=self.pick_source).grid(row=0, column=2, padx=(12, 0), pady=(0, 16))

        ttk.Label(root, text="导出格式：").grid(row=1, column=0, sticky="w")
        formats = ttk.Frame(root)
        formats.grid(row=1, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(formats, text="Excel 工作簿 (.xlsx)", variable=self.format, value=".xlsx").pack(side="left")
        ttk.Radiobutton(formats, text="Excel 97-2003 (.xls)", variable=self.format, value=".xls").pack(side="left", padx=(20, 0))

        self.convert_button = ttk.Button(root, text="转换并保存", command=self.start_conversion)
        self.convert_button.grid(row=2, column=1, sticky="w", pady=(26, 12))
        ttk.Label(root, textvariable=self.status, foreground="#444").grid(row=3, column=0, columnspan=3, sticky="w")

    def pick_source(self) -> None:
        filename = filedialog.askopenfilename(title="选择 ET 文件", filetypes=[("WPS 表格", "*.et"), ("所有文件", "*.*")])
        if filename:
            self.source.set(filename)
            self.status.set("已选择文件，点击“转换并保存”继续")

    def start_conversion(self) -> None:
        source_text = self.source.get()
        if not source_text:
            messagebox.showwarning("尚未选择文件", "请先选择一个 .et 文件。")
            return
        source = Path(source_text)
        if not source.is_file():
            messagebox.showerror("文件不存在", "所选文件已不存在或无法访问。")
            return
        destination_text = filedialog.asksaveasfilename(
            title="保存 Excel 文件", defaultextension=self.format.get(),
            initialfile=f"{source.stem}{self.format.get()}",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("Excel 97-2003", "*.xls")],
        )
        if not destination_text:
            return
        destination = Path(destination_text).with_suffix(self.format.get())
        self.convert_button.config(state="disabled")
        self.status.set("正在转换，请稍候…")
        threading.Thread(target=self._convert_worker, args=(source, destination), daemon=True).start()

    def _convert_worker(self, source: Path, destination: Path) -> None:
        try:
            result = convert(source, destination)
            self.after(0, lambda: self._complete(True, f"{result}：{destination}"))
        except Exception as exc:
            # Exception variables are cleared when an except block ends. Bind the
            # message now, otherwise this deferred callback raises NameError and
            # the user never sees the real conversion error.
            detail = str(exc)
            self.after(0, lambda message=detail: self._complete(False, message))

    def _complete(self, success: bool, detail: str) -> None:
        self.convert_button.config(state="normal")
        self.status.set(detail)
        if success:
            messagebox.showinfo("转换完成", detail)
        else:
            messagebox.showerror("无法转换", detail)


if __name__ == "__main__":
    if sys.platform != "win32":
        raise SystemExit("此工具目前仅支持 Windows。")
    ConverterApp().mainloop()
