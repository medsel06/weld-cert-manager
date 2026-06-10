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
| `OLCEK_TAVAN` | `2.0` | Üst büyütme sınırı. |
| `MIN_PUNTO`   | `11`  | Giriş/buton/etiket için hedef minimum punto. |
| `YAZI_TIPI`   | `Segoe UI` | Tüm formlarda kullanılan yazı tipi. |

## Çalışma mantığı (özet)
- Ölçek faktörü = `max(DPI/96, MIN_PUNTO/en-küçük-font, 1.0)`, `OLCEK_TAVAN` ile
  sınırlı. Tek faktör tüm kontrollere uygulanır → oranlar korunur, taşma olmaz.
- Tema: butonlar role göre renklenir (Kaydet=yeşil, Kapat=kırmızı,
  Temizle/Geri Al=gri, Hesapla=turuncu, menü=lacivert); başlık etiketi
  (`lblTitle`) lacivert bant, bölüm etiketleri (`lblBolum*`) vurgulu, giriş
  kutuları beyaz/düz çerçeve.
- Aynı form iki kez ölçeklenmesin diye `frm.Tag` ile işaretlenir.
