# Parti Takip Sistemi

## Kullanım

1. **Excel'i güncelle:** `veri.xlsx` dosyasını repo köküne koyun
2. **GitHub'a push yapın:** Otomatik olarak GitHub Actions çalışır
3. **Netlify'a deploy:** HTML + JSON otomatik deploy olur
4. **Mobilde açın:** Netlify linki her zaman güncel

## Manuel Çalıştırma

```bash
pip install pandas numpy openpyxl
python parti_takip.py
```

## Yerel Geliştirme

HTML şablonunu notebook'tan güncellemek için:
```bash
python extract_html.py
```