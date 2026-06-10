# Form Stil Motoru — Kurulum (modalları büyütme + güzelleştirme)

`FormStil.bas`, Codex'in eklediği 4 modalı (UserForm) açılırken **ekranın
DPI'ına göre orantılı büyütür**, yazıları Segoe UI + okunur boyuta çeker ve
kurumsal renk teması uygular. Böylece laptopta küçük görünen formlar her
ekranda net okunur; büyük monitördeki görünüm korunur.

## Neden gerekli
VBA UserForm'ları DPI'a duyarlı değildir: yüksek çözünürlüklü / %150 ölçekli
laptop ekranlarında küçük çizilir. Bu motor formu açılışta ölçekleyerek bu
Windows/Excel kısıtını çalışma anında telafi eder.

## Kurulum (tek seferlik, ~2 dakika)

1. Excel'de dosyayı açın, **Alt+F11** (VBA editörü).
2. **Dosya > Dosya Al...** (File > Import File) > `FormStil.bas` seçin.
3. Proje ağacında şu 4 formun her birine çift tıklayıp kod penceresini açın
   ve **`UserForm_Initialize`** olayının **ilk satırına** şu satırı ekleyin:

   ```vba
   FormStil.Uygula Me
   ```

   - **frmAnaMenu**
     ```vba
     Private Sub UserForm_Initialize()
         FormStil.Uygula Me          '<-- eklenen satır
         lblInfo.Caption = "Personel icra dosyalarını girin, sıraya alın ve ay sonu kesintilerini kaydedin."
     End Sub
     ```
   - **frmPersonelGiris**
     ```vba
     Private Sub UserForm_Initialize()
         FormStil.Uygula Me          '<-- eklenen satır
         txtIseGiris.Value = Format$(Date, "dd.mm.yyyy")
     End Sub
     ```
   - **frmIcraGiris**
     ```vba
     Private Sub UserForm_Initialize()
         FormStil.Uygula Me          '<-- eklenen satır
         UI_PersonelComboYukle cboPersonel
         ' ...mevcut kodun geri kalanı aynı kalsın...
     ```
   - **frmKesintiGiris**
     ```vba
     Private Sub UserForm_Initialize()
         FormStil.Uygula Me          '<-- eklenen satır
         UI_PersonelComboYukle cboPersonel
         ' ...mevcut kodun geri kalanı aynı kalsın...
     ```

4. **Alt+Q** ile editörden çıkın, dosyayı **.xlsm** olarak kaydedin.
   Programı açıp bir formu açın — büyük, renkli ve okunur gelmeli.

## İnce ayar (gerekirse)

`FormStil.bas` başındaki sabitlerle oynanır:

| Sabit | Varsayılan | Açıklama |
|-------|-----------|----------|
| `OLCEK_ZORLA` | `0` | `0` = ekran DPI'ından otomatik. Hâlâ küçük/büyük gelirse elle verin (örn. `1.3` = %30 büyük). |
| `OLCEK_TAVAN` | `1.8` | Üst büyütme sınırı. |
| `OLCEK_DIP`   | `0.6` | Ekrana sığdırırken en fazla bu kadar küçültür; altına inecekse kaydırma çubuğu açılır. |
| `MIN_PUNTO`   | `11`  | Giriş/buton/etiket için hedef minimum punto. |
| `EKRAN_PAY_X/Y` | `0.94 / 0.90` | Ekranın kullanılabilir oranı (görev çubuğu/başlık payı). |
| `YAZI_TIPI`   | `Segoe UI` | Tüm formlarda kullanılan yazı tipi. |

## Çalışma mantığı (özet)
1. **İçeriğe göre boyutlama:** tüm kontrolleri saran kutu hesaplanır, form iç
   ölçüsü buna + kenar boşluğuna ayarlanır → iç taşma, yarım kontrol veya
   gereksiz boşluk kalmaz (ölçekten bağımsız her zaman çalışır).
2. **Orantılı ölçek:** faktör = `max(DPI/96, MIN_PUNTO/en-küçük-font)`. Tek
   faktör tüm kontrollere uygulanır → oranlar korunur.
3. **Ekrana sığdırma:** form ekranın `%94 × %90`'ını aşacaksa ölçek otomatik
   küçültülür. Örn. geniş `frmIcraGiris` küçük bir 1366×768 laptopta `~0.82`
   ölçeğe inerek ekrana sığar; küçük formlar yüksek-DPI'da `~1.5` büyür.
4. **Güvenlik ağı:** her şeye rağmen sığmazsa kaydırma çubuğu açılır →
   hiçbir kontrol ekran dışında kalmaz / kesilmez.
5. **Ortalama + tema:** form ekranda ortalanır; butonlar role göre renklenir
   (Kaydet=yeşil, Kapat=kırmızı, Temizle/Geri Al=gri, Hesapla=turuncu,
   menü=lacivert), başlık (`lblTitle`) lacivert bant, bölüm başlıkları
   (`lblBolum*`) vurgulu, giriş kutuları beyaz/düz çerçeve, Segoe UI.

`frm.Tag` ile aynı form iki kez ölçeklenmez. 32/64-bit (PtrSafe) uyumludur.

## Doğrulama
Boyutlandırma mantığı 4 senaryoda (masaüstü %100, laptop %125/%150, küçük
1366×768 laptop) gerçek form ölçüleriyle simüle edildi; hiçbirinde ekran
taşması çıkmadı. En büyük form (`frmIcraGiris`, 701×406 pt) dar ekranlarda
otomatik küçülerek sığdı.
