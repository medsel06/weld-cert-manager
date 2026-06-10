# Personel İcra Takip — Kurulum ve Yeni Kurgu

Bu paket, uygulamayı **personel-merkezli** yeni kurguyla sıfırdan kurar:
eski form ve modülleri kaldırır, yenilerini Excel COM otomasyonu ile
(`VBComponents.Add` + `Designer.Controls.Add` + `AddFromString`) oluşturur.

## Paket içeriği

| Dosya | Görev |
|---|---|
| `Kur.ps1` | Kurulum scripti (Excel'i açar, her şeyi kurar, kaydeder) |
| `arayuz.json` | 9 formun yerleşimi (171 kontrol; liste kolon başlıkları otomatik hizalı) |
| `src/*.vba` | Modüller ve form kodları (motor + UI + stil + 9 form) |

## Kurulum (5 dakika)

1. **Bir kez:** Excel > Dosya > Seçenekler > Güven Merkezi > Güven Merkezi
   Ayarları > Makro Ayarları > **[x] VBA proje nesne modeline erişime güven** → Tamam.
2. `Personel_Icra_Takip_Profesyonel.xlsm` dosyasını bu klasöre koyun
   (veya scripte `-Dosya "C:\yol\dosya.xlsm"` parametresi verin). Excel'de açıksa kapatın.
3. `Kur.ps1` → sağ tık → **PowerShell ile Çalıştır**.
   (Gerekirse: `powershell -ExecutionPolicy Bypass -File Kur.ps1`)
4. Script otomatik **yedek alır** (`*_yedek_tarih.xlsm`), kurar, kaydeder.
5. Dosyayı Excel'de açın → "İçeriği Etkinleştir" → program tam ekran başlar.

## Yeni kurgu

```
ANA MENÜ
 ├─ PERSONELLER            arama + hizalı liste + ekle/düzenle/sil
 │   └─ (çift tık) İCRA KARTI   <- uygulamanın kalbi
 │        • dosyalar öncelik sırasıyla (kalan, bekleyen, durum)
 │        • seçili dosyanın BORDRO KESİNTİLERİ ve DAİREYE ÖDEMELERİ yan yana
 │        • Yeni Dosya · Düzenle · Sil · Kesinti Ekle · ÖDEME EKLE
 ├─ AY SONU KESİNTİLERİ    dönem seç -> önizleme -> hesapla/kaydet/geri al
 ├─ PERSONEL ÖZETİ         kişi başı toplamlar (ödenen/bekleyen dahil)
 └─ ÇIKIŞ                  kaydeder, Excel görünümünü geri verir
```

## Bu kurulumda düzeltilen hatalar / eklenenler

- **İcra dairesine ödeme takibi** (yeni `İCRA ÖDEMELERİ` sayfası + dosya başına
  *Daireye Ödenen* / *Ödenmesi Bekleyen* formül kolonları W/X). Ödeme girişi
  icra kartındaki **ÖDEME EKLE** ile yapılır; bekleyen tutarı aşarsanız uyarır.
- Personele **çift tık → icra kartı** (eskiden yoktu).
- İcra dosyası **düzenleme ve silme** (bağlı kesinti/ödeme kayıtlarıyla birlikte,
  dosya no değişirse geçmiş kayıtlar otomatik taşınır).
- **Durum sözlüğü birleştirildi:** her yerde `ÖDEMEDE / SIRADA / KAPANDI`
  (eski `BEKLİYOR` karışıklığı ve filtre kaybolması giderildi).
- **Liste başlıkları kolonlarla birebir hizalı** (başlık etiketleri kolon
  genişliklerinden otomatik konumlanır — tek kaynak: `arayuz.json`).
- Ana menüdeki yanlış bağlı "Panel" butonu kaldırıldı; **zombi Excel** düzeltildi
  (menü X ile kapansa bile görünürlük geri gelir).
- Tüm formlar `FormStil` motorundan geçer: DPI ölçekleme + ekrana sığdırma +
  kurumsal tema (Kaydet yeşil, Kapat kırmızı, vb.).

## Sorun giderme

- *"VBA projesine erişilemiyor"* → 1. adımdaki Güven Merkezi ayarı açık değil.
- Script hata verirse dosyanız bozulmaz; aynı klasördeki `_yedek_` kopyası durur.
- Veri sayfalarını görmek için: Alt+F8 → `SAYFALARI_GOSTER`
  (program moduna dönüş: `SAYFALARI_GIZLE`).
