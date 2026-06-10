Windows'ta bir Excel VBA uygulaması kuracağız. Senin görevin: kurulumu çalıştırmak,
hata çıkarsa teşhis edip çözmek, sonucu doğrulamak. Adım adım ilerle, her komutun
çıktısını bana göster. Türkçe konuş.

## DURUM
- Her şey şu klasörde: C:\Users\ykahm\Downloads\icratakip
- Excel dosyası: C:\Users\ykahm\Downloads\icratakip\Personel_Icra_Takip_Profesyonel.xlsm
- Aynı klasörde kurulum paketi olmalı: Kur.ps1, BASLAT.bat, arayuz.json, KURULUM.md ve src\ klasörü.
  EKSİKSE: GitHub "medsel06/weld-cert-manager", branch "claude/sharp-bohr-mhy6w5", "kurulum/"
  klasöründen indir; ya da kullanıcıdaki IcraTakip_Kurulum.zip'i bu klasöre aç.
- Kur.ps1 ne yapar: Excel'i COM ile açar (makrolar kapalı, otomatik yedek alır),
  "İCRA ÖDEMELERİ" sayfasını + W/X formül kolonlarını ekler, eski form/modülleri söker,
  3 modül + 9 UserForm'u (arayuz.json + src\*.vba'dan) sıfırdan kurar, kaydeder.

## 1) Klasöre geç ve dosya engellerini kaldır
cd C:\Users\ykahm\Downloads\icratakip
Get-ChildItem -Recurse | Unblock-File
Get-ChildItem    # Kur.ps1, arayuz.json, src, xlsm hepsi burada mı kontrol et

## 2) VBA nesne modeline erişimi aç (manuel Excel ayarı yerine registry)
$ver = (Get-ChildItem 'HKCU:\Software\Microsoft\Office' |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object PSChildName -Descending | Select-Object -First 1).PSChildName
$key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
"Office $ver -> AccessVBOM=1"

## 3) Excel'i tamamen kapat
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force

## 4) Kurulumu çalıştır
powershell -NoProfile -ExecutionPolicy Bypass -File .\Kur.ps1 -Dosya ".\Personel_Icra_Takip_Profesyonel.xlsm"
# Beklenen: "1/4 ... 2/4 ... 3/4 ... 4/4 ... KURULUM TAMAMLANDI"
# Tam günlük kurulum_log.txt dosyasına yazılır; sorun olursa onu oku.

## 5) Hata çıkarsa teşhis et ve çöz
- "VBA projesine erişilemiyor" -> Adım 2 işe yaramadı; Office sürümünü ve registry yolunu
  doğrula, gerekirse Excel Güven Merkezi > Makro Ayarları'ndan "VBA proje nesne modeline
  erişime güven"i manuel aç ve tekrar dene.
- "Excel dosyası bulunamadı" -> xlsm adını/yolunu kontrol et.
- COM/Excel hatası -> Excel masaüstü sürümü kurulu mu (Office 365/2016+) doğrula.
  Hata satırını ve kurulum_log.txt'nin son 30 satırını bana göster.

## 6) Doğrula
- Çalışmayı bozmadan dosyayı incele: kurulum_log.txt'de "KURULUM TAMAMLANDI" var mı,
  ve oluşan _yedek_*.xlsm yedeği duruyor mu kontrol et.
- Sonra Excel'de Personel_Icra_Takip_Profesyonel.xlsm'i aç, "İçeriği Etkinleştir" de.
  Beklenen: arka plan gizlenir, "PERSONEL İCRA TAKİP" ana menüsü açılır ->
  PERSONELLER -> kişiye çift tık -> İcra Kartı (dosyalar + kesintiler + daireye ödemeler +
  "ÖDEME EKLE" butonu).
- Ekran görüntüsü al ve bana sonucu bildir.

Başla ve her adımın çıktısını paylaş.
