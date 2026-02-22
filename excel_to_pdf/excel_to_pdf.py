# -*- coding: utf-8 -*-
"""
Excel to PDF Converter - Kurumsal Masaüstü Uygulaması
Microsoft Excel dosyalarını (.xls, .xlsx) yüksek kaliteli PDF'e dönüştürür.
Tek dosya veya klasör toplu işlem destekli.
"""

import os
import sys
import time
import threading
import traceback
from pathlib import Path
from tkinter import (
    Tk, Frame, Label, Entry, Button, Text, Scrollbar,
    StringVar, messagebox, ttk, font as tkfont
)
from tkinter import filedialog

# Windows için drag-drop (ctypes)
if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

# Excel COM için pywin32
try:
    import win32com.client
except ImportError:
    win32com = None


# --- Sabitler ---
APP_TITLE = "Excel → PDF Dönüştürücü"
CORPORATE_HEADER = "Kurumsal Excel - PDF Dönüştürme Sistemi"
CORPORATE_FOOTER = "© Tüm hakları saklıdır. Yüksek kaliteli çıktı."
EXCEL_EXTENSIONS = (".xls", ".xlsx")
PDF_QUALITY = 0  # xlQualityStandard (0); xlQualityMinimum (1) kalite düşük


def get_excel_files_from_folder(folder_path: str) -> list:
    """Klasördeki tüm .xls ve .xlsx dosyalarının listesini döndürür."""
    folder = Path(folder_path)
    if not folder.is_dir():
        return []
    files = []
    for ext in EXCEL_EXTENSIONS:
        files.extend(folder.glob(f"*{ext}"))
    return sorted([str(f) for f in files])


def get_pdf_path(excel_path: str) -> str:
    """Excel dosya yolundan aynı isimde PDF yolu üretir."""
    p = Path(excel_path)
    return str(p.parent / f"{p.stem}.pdf")


def excel_to_pdf_single(excel_app, excel_path: str, pdf_path: str, log_callback=None) -> tuple:
    """
    Tek bir Excel dosyasını PDF'e çevirir.
    excel_app: win32com Excel Application instance
    Returns: (başarılı: bool, hata_mesajı: str veya None)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    if not os.path.isfile(excel_path):
        return False, "Dosya bulunamadı."

    try:
        size_mb = os.path.getsize(excel_path) / (1024 * 1024)
        log(f"  Açılıyor: {Path(excel_path).name} ({size_mb:.2f} MB)")
        start = time.perf_counter()

        # ReadOnly=True, açık dosyada sorun azaltır
        wb = excel_app.Workbooks.Open(
            os.path.abspath(excel_path),
            ReadOnly=True,
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True
        )
        try:
            # ExportAsFixedFormat: Type=0 (PDF), Quality=0 (Standard)
            wb.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=os.path.abspath(pdf_path),
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )
            elapsed = time.perf_counter() - start
            out_size = os.path.getsize(pdf_path) / (1024 * 1024) if os.path.isfile(pdf_path) else 0
            log(f"  ✓ PDF oluşturuldu: {Path(pdf_path).name} ({out_size:.2f} MB) - {elapsed:.2f} sn")
            return True, None
        finally:
            wb.Close(SaveChanges=False)
    except Exception as e:
        err_msg = str(e)
        log(f"  ✗ Hata: {err_msg}")
        return False, err_msg


def run_conversion(path: str, is_folder: bool, progress_callback, log_callback, done_callback):
    """
    Arka planda Excel → PDF dönüşümünü çalıştırır.
    progress_callback(current, total, message)
    log_callback(message)
    done_callback(success_count, fail_count, error_summary)
    """
    success_count = 0
    fail_count = 0
    errors = []

    def log(msg):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    def progress(cur, total, msg=""):
        if progress_callback:
            try:
                progress_callback(cur, total, msg)
            except Exception:
                pass

    excel_app = None
    try:
        if win32com is None:
            log("HATA: pywin32 yüklü değil. 'pip install pywin32' ile yükleyin.")
            done_callback(0, 0, "pywin32 bulunamadı")
            return

        if is_folder:
            if not os.path.isdir(path):
                log("HATA: Klasör bulunamadı.")
                done_callback(0, 0, "Klasör bulunamadı")
                return
            files = get_excel_files_from_folder(path)
            if not files:
                log("Bu klasörde .xls veya .xlsx dosyası yok.")
                done_callback(0, 0, "Excel dosyası bulunamadı")
                return
        else:
            if not os.path.isfile(path):
                log("HATA: Dosya bulunamadı.")
                done_callback(0, 0, "Dosya bulunamadı")
                return
            path_lower = path.lower()
            if not any(path_lower.endswith(ext) for ext in EXCEL_EXTENSIONS):
                log("HATA: Geçerli bir Excel dosyası seçin (.xls veya .xlsx).")
                done_callback(0, 0, "Geçersiz dosya türü")
                return
            files = [path]

        total = len(files)
        log(f"Toplam {total} dosya işlenecek. Excel başlatılıyor...")

        try:
            excel_app = win32com.client.DispatchEx("Excel.Application")
        except Exception as com_err:
            err_text = str(com_err)
            if "invalid class string" in err_text.lower() or "0x800401f3" in err_text or "coinitialize" in err_text.lower():
                log("HATA: Microsoft Excel yüklü değil veya COM kayıtlı değil. Lütfen Office/Excel kurulumunu kontrol edin.")
                done_callback(0, 0, "Excel bulunamadı veya COM erişim hatası")
            else:
                log(f"HATA: Excel başlatılamadı: {err_text}")
                done_callback(0, 0, err_text)
            return

        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False

        for i, excel_path in enumerate(files):
            progress(i, total, Path(excel_path).name)
            pdf_path = get_pdf_path(excel_path)
            ok, err = excel_to_pdf_single(excel_app, excel_path, pdf_path, log_callback=log)
            if ok:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{Path(excel_path).name}: {err}")

            progress(i + 1, total, Path(excel_path).name)

        summary = ""
        if errors:
            summary = "; ".join(errors[:3])
            if len(errors) > 3:
                summary += f" (+{len(errors)-3} daha)"
        done_callback(success_count, fail_count, summary)

    except Exception as e:
        log(f"Kritik hata: {e}")
        traceback.print_exc()
        try:
            done_callback(success_count, fail_count, str(e))
        except Exception:
            pass
    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
            excel_app = None
        log("Excel kapatıldı.")


# --- Windows Drag & Drop (ctypes) ---
if sys.platform == "win32":
    try:
        WM_DROPFILES = 0x0233
        GWL_WNDPROC = -4
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        # 64-bit Windows'ta GetWindowLongPtrW / SetWindowLongPtrW kullan
        if hasattr(user32, "GetWindowLongPtrW"):
            GetWindowLong = user32.GetWindowLongPtrW
            SetWindowLong = user32.SetWindowLongPtrW
        else:
            GetWindowLong = user32.GetWindowLongW
            SetWindowLong = user32.SetWindowLongW
    except Exception:
        WM_DROPFILES = None
        GetWindowLong = SetWindowLong = user32 = shell32 = None
else:
    WM_DROPFILES = None
    user32 = shell32 = GetWindowLong = SetWindowLong = None


class ExcelToPdfApp:
    """Ana uygulama penceresi."""

    def __init__(self):
        self.root = Tk()
        self.root.title(APP_TITLE)
        self.root.minsize(520, 420)
        self.root.geometry("600x480")
        self.root.resizable(True, True)

        # Kurumsal renkler
        self.bg_dark = "#1e3a5f"
        self.bg_medium = "#2c5282"
        self.bg_light = "#edf2f7"
        self.accent = "#3182ce"
        self.text_light = "#f7fafc"
        self.text_dark = "#2d3748"

        self.root.configure(bg=self.bg_dark)

        self.path_var = StringVar()
        self.is_folder_mode = False
        self.conversion_thread = None
        self._drop_callback = None
        self._old_wndproc = None

        self._build_ui()
        # Sürükle-bırak penceresi hazır olduktan sonra etkinleştirilir
        self.root.after(200, self._setup_drag_drop)

    def _build_ui(self):
        """Arayüz bileşenlerini oluşturur."""
        main = Frame(self.root, bg=self.bg_dark, padx=20, pady=20)
        main.pack(fill="both", expand=True)

        # --- Başlık ---
        title_font = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        header = Label(
            main, text=CORPORATE_HEADER,
            font=title_font, fg=self.text_light, bg=self.bg_dark
        )
        header.pack(pady=(0, 16))

        # --- Dosya/Klasör girişi ---
        input_frame = Frame(main, bg=self.bg_dark)
        input_frame.pack(fill="x", pady=(0, 8))

        Label(
            input_frame, text="Dosya veya klasör yolu:",
            font=("Segoe UI", 9), fg=self.text_light, bg=self.bg_dark
        ).pack(anchor="w")

        entry_frame = Frame(input_frame, bg=self.bg_dark)
        entry_frame.pack(fill="x", pady=4)

        self.entry = Entry(
            entry_frame,
            textvariable=self.path_var,
            font=("Segoe UI", 10),
            bg="white", fg=self.text_dark,
            insertbackground=self.text_dark,
            relief="flat", highlightthickness=1, highlightcolor=self.accent,
            highlightbackground="#4a5568"
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=6, ipadx=8, padx=(0, 8))

        btn_style = {"font": ("Segoe UI", 9), "cursor": "hand2", "relief": "flat"}
        self.btn_file = Button(
            entry_frame, text="Dosya Seç",
            command=self._on_select_file,
            bg=self.bg_medium, fg=self.text_light,
            activebackground=self.accent, activeforeground="white",
            **btn_style
        )
        self.btn_file.pack(side="left", padx=(0, 6), ipady=4, ipadx=10)

        self.btn_folder = Button(
            entry_frame, text="Klasör Seç",
            command=self._on_select_folder,
            bg=self.bg_medium, fg=self.text_light,
            activebackground=self.accent, activeforeground="white",
            **btn_style
        )
        self.btn_folder.pack(side="left", padx=(0, 6), ipady=4, ipadx=10)

        # --- PDF'e Çevir ---
        self.btn_convert = Button(
            main, text="PDF'e Çevir",
            command=self._on_convert,
            font=("Segoe UI", 11, "bold"),
            bg=self.accent, fg="white",
            activebackground="#2b6cb0", activeforeground="white",
            cursor="hand2", relief="flat",
            padx=24, pady=10
        )
        self.btn_convert.pack(pady=16)

        # --- Progress bar ---
        self.progress_var = StringVar(value="")
        self.progress_label = Label(
            main, textvariable=self.progress_var,
            font=("Segoe UI", 9), fg=self.text_light, bg=self.bg_dark
        )
        self.progress_label.pack(anchor="w", pady=(0, 4))

        self.progress = ttk.Progressbar(main, mode="determinate", length=400)
        self.progress.pack(fill="x", pady=(0, 12))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "TProgressbar",
            troughcolor=self.bg_medium,
            background=self.accent,
            thickness=8
        )

        # --- Log alanı ---
        log_label = Label(
            main, text="Durum / Log:",
            font=("Segoe UI", 9), fg=self.text_light, bg=self.bg_dark
        )
        log_label.pack(anchor="w", pady=(0, 4))

        log_frame = Frame(main, bg="#2d3748")
        log_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.log_text = Text(
            log_frame,
            font=("Consolas", 9),
            bg="#2d3748", fg="#e2e8f0",
            insertbackground="white",
            wrap="word", relief="flat",
            padx=8, pady=8, state="normal"
        )
        scroll = Scrollbar(log_frame, command=self.log_text.yview, bg="#4a5568")
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # --- Footer ---
        footer = Label(
            main, text=CORPORATE_FOOTER,
            font=("Segoe UI", 8), fg="#a0aec0", bg=self.bg_dark
        )
        footer.pack(pady=(8, 0))

        self._log("Uygulama hazır. Dosya veya klasör seçip 'PDF'e Çevir' ile işlemi başlatın.")
        if win32com is None:
            self._log("UYARI: pywin32 yüklü değil. Excel dönüşümü için 'pip install pywin32' gerekir.")

    def _log(self, message: str):
        """Log alanına mesaj ekler (thread-safe)."""
        def _append():
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.update_idletasks()

        self.root.after(0, _append)

    def _on_select_file(self):
        """Dosya seçme diyaloğu."""
        path = filedialog.askopenfilename(
            title="Excel dosyası seçin",
            filetypes=[
                ("Excel dosyaları", "*.xls *.xlsx"),
                ("Tüm dosyalar", "*.*")
            ]
        )
        if path:
            self.path_var.set(path)
            self.is_folder_mode = False
            self._log(f"Dosya seçildi: {path}")

    def _on_select_folder(self):
        """Klasör seçme diyaloğu."""
        path = filedialog.askdirectory(title="Klasör seçin (içindeki tüm Excel dosyaları işlenecek)")
        if path:
            self.path_var.set(path)
            self.is_folder_mode = True
            self._log(f"Klasör seçildi: {path}")

    def _on_convert(self):
        """PDF dönüşümünü başlatır (thread'de)."""
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("Uyarı", "Lütfen bir dosya veya klasör seçin.")
            return
        if self.conversion_thread and self.conversion_thread.is_alive():
            messagebox.showinfo("Bilgi", "Dönüşüm zaten devam ediyor.")
            return

        self.btn_convert.config(state="disabled")
        self.btn_file.config(state="disabled")
        self.btn_folder.config(state="disabled")
        self.progress["value"] = 0
        self.progress_var.set("")
        self._log("-" * 40)
        self._log("Dönüşüm başlatıldı...")

        def progress_update(current, total, message):
            def _update():
                if total > 0:
                    pct = 100 * current / total
                    self.progress["value"] = pct
                    self.progress_var.set(f"{current}/{total} - {message}")
                self.root.update_idletasks()

            self.root.after(0, _update)

        def done_update(success_count, fail_count, error_summary):
            def _done():
                self.btn_convert.config(state="normal")
                self.btn_file.config(state="normal")
                self.btn_folder.config(state="normal")
                self.progress["value"] = 100
                self.progress_var.set("Tamamlandı.")
                self._log(f"Sonuç: {success_count} başarılı, {fail_count} hata.")
                if error_summary:
                    self._log(f"Hata özeti: {error_summary}")
                if success_count > 0:
                    messagebox.showinfo(
                        "İşlem tamamlandı",
                        f"{success_count} dosya PDF'e dönüştürüldü.\n"
                        + (f"{fail_count} dosyada hata oluştu." if fail_count else "")
                    )
                elif fail_count > 0:
                    messagebox.showerror("Hata", f"Dönüşüm başarısız.\n{error_summary}")
                self.root.update_idletasks()

            self.root.after(0, _done)

        self.conversion_thread = threading.Thread(
            target=run_conversion,
            args=(
                path,
                self.is_folder_mode,
                progress_update,
                self._log,
                done_update,
            ),
            daemon=True
        )
        self.conversion_thread.start()

    def _setup_drag_drop(self):
        """Windows'ta sürükle-bırak desteği (ctypes ile WM_DROPFILES)."""
        if sys.platform != "win32" or WM_DROPFILES is None or user32 is None or shell32 is None:
            return
        try:
            self.root.update_idletasks()
            self.root.update()
            hwnd = self.root.winfo_id()
            if not hwnd:
                return
            # Tk'de bazen alt pencere id'si döner; toplevel HWND için GetParent
            for _ in range(3):
                ph = user32.GetParent(hwnd)
                if not ph:
                    break
                hwnd = ph
            shell32.DragAcceptFiles(hwnd, True)
            LRESULT = ctypes.c_ssize_t  # 32/64 bit uyumlu
            WNDPROCTYPE = ctypes.WINFUNCTYPE(
                LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
            )

            def wndproc(hwnd, msg, wparam, lparam):
                if msg == WM_DROPFILES:
                    self._handle_wm_dropfiles(hwnd, wparam)
                    return 0
                return user32.CallWindowProcW(self._old_wndproc, hwnd, msg, wparam, lparam)

            self._wndproc = WNDPROCTYPE(wndproc)
            self._old_wndproc = GetWindowLong(hwnd, GWL_WNDPROC)
            if self._old_wndproc:
                SetWindowLong(hwnd, GWL_WNDPROC, ctypes.cast(self._wndproc, ctypes.c_void_p).value)
                self._log("Sürükle-bırak etkin.")
        except Exception as e:
            self._log(f"Sürükle-bırak başlatılamadı: {e}")

    def _handle_wm_dropfiles(self, hwnd, wparam):
        """WM_DROPFILES işleyicisi: bırakılan dosya/klasör yolunu alır."""
        try:
            shell32 = ctypes.windll.shell32
            max_path = 260
            buf = ctypes.create_unicode_buffer(max_path)
            num = shell32.DragQueryFileW(wintypes.HDROP(wparam), 0xFFFFFFFF, None, 0)
            if num >= 1:
                shell32.DragQueryFileW(wintypes.HDROP(wparam), 0, buf, max_path)
                path = buf.value
                shell32.DragFinish(wparam)
                self.root.after(0, lambda: self._on_drop(path))
        except Exception as e:
            self._log(f"Sürükle-bırak hatası: {e}")

    def _on_drop(self, path: str):
        """Dosya veya klasör bırakıldığında çağrılır."""
        if not path or not os.path.exists(path):
            return
        self.path_var.set(path)
        self.is_folder_mode = os.path.isdir(path)
        self._log(f"Sürükle-bırak: {'Klasör' if self.is_folder_mode else 'Dosya'} seçildi - {path}")

    def run(self):
        """Uygulama döngüsünü başlatır."""
        self.root.mainloop()


def main():
    """Giriş noktası."""
    app = ExcelToPdfApp()
    app.run()


if __name__ == "__main__":
    main()


# PyInstaller ile exe oluşturmak için (tek dosya, konsolsuz, özel ikon):
# pyinstaller --onefile --noconsole --icon=excel.ico excel_to_pdf.py
