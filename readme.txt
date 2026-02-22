================================================================================
  SPOR ISTANBUL CLI & EXCEL-TO-PDF
  GitHub: https://github.com/tekbasserkan
================================================================================

Bu depoda iki ana uygulama vardir:

  1. SPOR ISTANBUL CLI (Go)
     - Salon rezervasyonu otomasyonu
     - Calistirma: go build -o spor-istanbul-cli ./cmd/main

  2. EXCEL-TO-PDF (Python / Windows)
     - Excel dosyalarini PDF'e cevirir (tek dosya veya klasor)
     - Calistirma: cd excel_to_pdf && pip install -r requirements.txt && python excel_to_pdf.py
     - EXE: cd excel_to_pdf && python -m PyInstaller --noconfirm excel_to_pdf.spec
     - Cikti: excel_to_pdf\dist\excel_to_pdf\ (excel_to_pdf.exe + _internal)

--------------------------------------------------------------------------------
CI/CD (GitHub Actions)
--------------------------------------------------------------------------------
  - CI:      Her push/PR'da lint ve build kontrolu
  - Build:   main branch'ta Excel-to-PDF zip + Go binary'leri (Artifacts)
  - Release: v1.0.0 gibi tag push'landiginda GitHub Release + dosyalar

--------------------------------------------------------------------------------
GITHUB'DA REPO OLUSTURMA (Repo yoksa)
--------------------------------------------------------------------------------
  1. https://github.com/new adresine gidin
  2. Repository name girin (ornegin: spor-istanbul-cli)
  3. Create repository (README eklemeyin)
  4. Asagidaki komutlari proje klasorunde calistirin:

     git init
     git add .
     git commit -m "Initial commit"
     git branch -M main
     git remote add origin https://github.com/tekbasserkan/REPO_ADI.git
     git push -u origin main

     REPO_ADI yerine olusturdugunuz repo adini yazin.

--------------------------------------------------------------------------------
Iletisim: https://github.com/tekbasserkan
================================================================================
