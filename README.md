# Excel-to-PDF

**GitHub:** [github.com/tekbasserkan/excel-to-pdf](https://github.com/tekbasserkan/excel-to-pdf)

Tek veya toplu Excel (.xls, .xlsx) dosyalarını yüksek kaliteli PDF'e çeviren Windows masaüstü uygulaması (Python / Tkinter).

---

## Gereksinimler

- Windows, Microsoft Excel (COM)
- Python 3.11+ ve `pywin32`

## Kurulum ve çalıştırma

```bash
cd excel_to_pdf
pip install -r requirements.txt
python excel_to_pdf.py
```

## EXE üretimi (PyInstaller)

```bash
cd excel_to_pdf
pip install pyinstaller
python -m PyInstaller --noconfirm excel_to_pdf.spec
```

Çıktı: `excel_to_pdf\dist\excel_to_pdf\excel_to_pdf.exe` (klasörü birlikte dağıtın).

---

## CI/CD (GitHub Actions)

| Workflow   | Tetikleyici        | Açıklama |
|-----------|--------------------|----------|
| **CI**    | Push/PR → main     | Python lint (ruff), bağımlılık kontrolü |
| **Build** | Push → main / Manuel | Excel-to-PDF Windows zip → Artifacts |
| **Release** | Tag `v*` push   | Build + GitHub Release (zip asset) |

Release için: `git tag v1.0.0 && git push origin v1.0.0`

---

## Lisans

Proje sahibi: [tekbasserkan](https://github.com/tekbasserkan). Ticari kullanım için iletişime geçin.
