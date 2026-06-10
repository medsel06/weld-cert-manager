# -*- coding: utf-8 -*-
"""
arayuz.json üretici — Kur.ps1'in kuracağı tüm UserForm yerleşimleri.

Tek kaynak ilkesi: liste kolon genişlikleri burada tanımlanır; hem ListBox'ın
ColumnWidths özelliği hem de kolon BAŞLIK etiketlerinin x-konumları buradan
hesaplanır -> başlıklar kolonlarla birebir hizalı olur (eski sürümün
'başlıklar uymuyor' sorunu kökten biter).

Kullanım: python3 tools/arayuz_olustur.py   -> kurulum/arayuz.json
"""
import json
import os

F = []  # formlar


def ctl(n, t, x, y, w, h, cap=None, size=None, bold=None, italic=None,
        cols=None, locked=None, align=None):
    c = {"n": n, "t": t, "x": x, "y": y, "w": w, "h": h}
    if cap is not None: c["cap"] = cap
    if size: c["size"] = size
    if bold: c["bold"] = True
    if italic: c["italic"] = True
    if locked: c["locked"] = True
    if align is not None: c["align"] = align   # 1=sol 2=orta 3=sağ
    if cols:
        c["colCount"] = len(cols)
        c["colWidths"] = ";".join(f"{w} pt" for _, w in cols)
    return c


def headers(prefix, list_x, y, cols, size=8):
    """Liste kolonlarıyla birebir hizalı başlık etiketleri üret."""
    out, x = [], list_x + 2
    i = 0
    for ad, w in cols:
        if w > 0 and ad:
            out.append(ctl(f"lblH_{prefix}{i}", "Label", x, y, w - 2, 13,
                           cap=ad, size=size, bold=True))
        x += w
        i += 1
    return out


def form(name, caption, w, h, controls):
    F.append({"name": name, "caption": caption, "w": w, "h": h,
              "controls": controls})


# =====================================================================
# 1) ANA MENÜ
# =====================================================================
form("frmAnaMenu", "Personel İcra Takip", 380, 312, [
    ctl("lblTitle", "Label", 0, 0, 380, 34, cap="PERSONEL İCRA TAKİP", size=14),
    ctl("lblInfo", "Label", 20, 46, 340, 26,
        cap="Personel seçin, icra kartından dosyaları, kesintileri ve daireye ödemeleri yönetin."),
    ctl("cmdPersonel", "CommandButton", 90, 84, 200, 38, cap="PERSONELLER", size=11),
    ctl("cmdDonem", "CommandButton", 90, 130, 200, 38, cap="AY SONU KESİNTİLERİ", size=11),
    ctl("cmdOzet", "CommandButton", 90, 176, 200, 38, cap="PERSONEL ÖZETİ", size=11),
    ctl("cmdCikis", "CommandButton", 90, 222, 200, 32, cap="ÇIKIŞ"),
    ctl("lblAlt", "Label", 20, 266, 340, 30, italic=True, size=8,
        cap="Bakım: Excel sayfalarını görmek için Alt+F8 > SAYFALARI_GOSTER"),
])

# =====================================================================
# 2) PERSONELLER  (liste + düzenleme + karta geçiş)
# =====================================================================
PL_COLS = [("Ad Soyad", 220), ("T.C. Kimlik", 100), ("Net Maaş", 95),
           ("İşe Giriş", 85), ("", 0), ("", 0)]   # son iki: açıklama + satır (gizli)
form("frmPersonelListe", "Personeller", 640, 452, [
    ctl("lblTitle", "Label", 0, 0, 640, 32, cap="PERSONELLER", size=13),
    ctl("lblAra", "Label", 16, 46, 36, 16, cap="Ara"),
    ctl("txtAra", "TextBox", 56, 44, 220, 20),
    ctl("lblSayac", "Label", 480, 46, 144, 16, align=3),
    *headers("pl", 16, 70, PL_COLS),
    ctl("lstPersonel", "ListBox", 16, 86, 608, 168, cols=PL_COLS),
    ctl("lblBolum1", "Label", 16, 262, 608, 16, cap="  Seçili Personel (düzenle / yeni)"),
    ctl("lblAd", "Label", 16, 286, 70, 16, cap="Ad Soyad"),
    ctl("txtAd", "TextBox", 90, 284, 230, 20),
    ctl("lblTC", "Label", 336, 286, 28, 16, cap="TC"),
    ctl("txtTC", "TextBox", 368, 284, 120, 20),
    ctl("lblNet", "Label", 16, 312, 70, 16, cap="Net Maaş"),
    ctl("txtNetMaas", "TextBox", 90, 310, 120, 20),
    ctl("lblTarih", "Label", 226, 312, 56, 16, cap="İşe Giriş"),
    ctl("txtIseGiris", "TextBox", 286, 310, 90, 20),
    ctl("lblAciklama", "Label", 392, 312, 56, 16, cap="Açıklama"),
    ctl("txtAciklama", "TextBox", 452, 310, 172, 20),
    ctl("cmdYeni", "CommandButton", 16, 344, 80, 28, cap="Yeni"),
    ctl("cmdKaydet", "CommandButton", 102, 344, 90, 28, cap="Kaydet"),
    ctl("cmdSil", "CommandButton", 198, 344, 70, 28, cap="Sil"),
    ctl("cmdKart", "CommandButton", 282, 344, 170, 28, cap="İCRA KARTINI AÇ", bold=True),
    ctl("cmdKapat", "CommandButton", 544, 344, 80, 28, cap="Kapat"),
    ctl("lblNot", "Label", 16, 382, 608, 28, italic=True, size=8,
        cap="İpucu: listede personele ÇİFT TIKLAYIN -> icra kartı açılır (dosyalar, kesintiler, daireye ödemeler)."),
])

# =====================================================================
# 3) PERSONEL İCRA KARTI  (uygulamanın kalbi)
# =====================================================================
KD_COLS = [("Önc.", 34), ("Dosya No", 95), ("Tarih", 64), ("Alacaklı", 168),
           ("İcra Tutarı", 76), ("Kalan", 76), ("Bekleyen", 70), ("Durum", 58), ("", 0)]
KK_COLS = [("Yıl", 38), ("Ay", 64), ("Tutar", 76), ("Açıklama", 144)]
KO_COLS = [("Tarih", 64), ("Tutar", 76), ("Banka/Dekont", 86), ("Açıklama", 88), ("", 0)]
form("frmPersonelKart", "Personel İcra Kartı", 700, 532, [
    ctl("lblTitle", "Label", 0, 0, 700, 32, cap="İCRA KARTI", size=13),
    ctl("lblKisi", "Label", 16, 40, 430, 18, bold=True),
    ctl("lblToplamlar", "Label", 450, 40, 234, 18, align=3),
    *headers("kd", 16, 64, KD_COLS),
    ctl("lstDosyalar", "ListBox", 16, 80, 668, 128, cols=KD_COLS),
    ctl("lblDosyaBilgi", "Label", 16, 214, 668, 28, italic=True, size=8),
    ctl("lblKesilen", "Label", 16, 246, 200, 16),
    ctl("lblOdenen", "Label", 230, 246, 210, 16),
    ctl("lblBekleyen", "Label", 450, 246, 234, 16, bold=True),
    ctl("lblBolum1", "Label", 16, 270, 330, 16, cap="  BORDRO KESİNTİLERİ", bold=True),
    *headers("kk", 16, 290, KK_COLS),
    ctl("lstKesintiler", "ListBox", 16, 306, 330, 122, cols=KK_COLS),
    ctl("lblBolum2", "Label", 354, 270, 330, 16, cap="  İCRA DAİRESİNE ÖDEMELER", bold=True),
    *headers("ko", 354, 290, KO_COLS),
    ctl("lstOdemeler", "ListBox", 354, 306, 330, 122, cols=KO_COLS),
    ctl("cmdYeniDosya", "CommandButton", 16, 440, 96, 30, cap="Yeni Dosya"),
    ctl("cmdDuzenle", "CommandButton", 118, 440, 86, 30, cap="Düzenle"),
    ctl("cmdDosyaSil", "CommandButton", 210, 440, 70, 30, cap="Sil"),
    ctl("cmdKesintiEkle", "CommandButton", 296, 440, 100, 30, cap="Kesinti Ekle"),
    ctl("cmdOdemeEkle", "CommandButton", 402, 440, 120, 30, cap="ÖDEME EKLE", bold=True),
    ctl("cmdYenile", "CommandButton", 538, 440, 66, 30, cap="Yenile"),
    ctl("cmdKapat", "CommandButton", 610, 440, 74, 30, cap="Kapat"),
    ctl("lblNot", "Label", 16, 478, 668, 26, italic=True, size=8,
        cap="Dosya seçince kesinti ve ödeme geçmişi yüklenir. Ödeme satırına çift tık: silme onayı sorar."),
])

# =====================================================================
# 4) İCRA DOSYASI GİRİŞ / DÜZENLE
# =====================================================================
form("frmIcraGiris", "İcra Dosyası", 560, 440, [
    ctl("lblTitle", "Label", 0, 0, 560, 32, cap="İCRA DOSYASI", size=13),
    ctl("lblPersonel", "Label", 16, 46, 76, 16, cap="Personel"),
    ctl("cboPersonel", "ComboBox", 100, 44, 260, 20),
    ctl("lblPersonelBilgi", "Label", 16, 68, 528, 14, italic=True, size=8),
    ctl("lblBolum1", "Label", 16, 88, 528, 16, cap="  Dosya Bilgileri"),
    ctl("lblDosyaTarihi", "Label", 16, 112, 76, 16, cap="Dosya Tarihi"),
    ctl("txtDosyaTarihi", "TextBox", 100, 110, 90, 20),
    ctl("lblDosyaNo", "Label", 206, 112, 64, 16, cap="Dosya No"),
    ctl("txtDosyaNo", "TextBox", 274, 110, 120, 20),
    ctl("lblOncelik", "Label", 410, 112, 50, 16, cap="Öncelik"),
    ctl("txtOncelik", "TextBox", 464, 110, 40, 20),
    ctl("lblDaire", "Label", 16, 138, 76, 16, cap="İcra Dairesi"),
    ctl("txtDaire", "TextBox", 100, 136, 444, 20),
    ctl("lblAlacakli", "Label", 16, 164, 76, 16, cap="Alacaklı"),
    ctl("txtAlacakli", "TextBox", 100, 162, 444, 20),
    ctl("lblBolum2", "Label", 16, 190, 528, 16, cap="  Tutar ve Kesinti Kuralı"),
    ctl("lblIcraTutar", "Label", 16, 214, 76, 16, cap="İcra Tutarı"),
    ctl("txtIcraTutar", "TextBox", 100, 212, 110, 20),
    ctl("lblSekil", "Label", 226, 214, 78, 16, cap="Kesinti Şekli"),
    ctl("cboKesintiSekli", "ComboBox", 308, 212, 100, 20),
    ctl("lblPay", "Label", 420, 214, 28, 16, cap="Oran"),
    ctl("txtPay", "TextBox", 452, 212, 30, 20),
    ctl("lblBolme", "Label", 486, 214, 8, 16, cap="/"),
    ctl("txtPayda", "TextBox", 498, 212, 30, 20),
    ctl("lblSabit", "Label", 16, 240, 80, 16, cap="Sabit Tutar"),
    ctl("txtSabit", "TextBox", 100, 238, 110, 20),
    ctl("lblOnceki", "Label", 226, 240, 132, 16, cap="Önceki Firmada Tahsil"),
    ctl("txtOncekiTahsil", "TextBox", 362, 238, 110, 20),
    ctl("lblBolum3", "Label", 16, 266, 528, 16, cap="  Diğer"),
    ctl("lblIban", "Label", 16, 290, 76, 16, cap="Banka/IBAN"),
    ctl("txtIban", "TextBox", 100, 288, 300, 20),
    ctl("lblAciklama", "Label", 16, 316, 76, 16, cap="Açıklama"),
    ctl("txtAciklama", "TextBox", 100, 314, 444, 20),
    ctl("lblNot", "Label", 16, 342, 528, 14, italic=True, size=8,
        cap="Tarih kutusuna tıklayın -> takvim açılır. Öncelik boş bırakılırsa sıradaki numara verilir."),
    ctl("cmdKaydet", "CommandButton", 150, 368, 110, 30, cap="Kaydet"),
    ctl("cmdTemizle", "CommandButton", 268, 368, 90, 30, cap="Temizle"),
    ctl("cmdKapat", "CommandButton", 366, 368, 90, 30, cap="Kapat"),
])

# =====================================================================
# 5) MANUEL KESİNTİ
# =====================================================================
form("frmKesintiGiris", "Manuel Kesinti", 420, 312, [
    ctl("lblTitle", "Label", 0, 0, 420, 32, cap="MANUEL KESİNTİ", size=13),
    ctl("lblPersonel", "Label", 16, 48, 88, 16, cap="Personel"),
    ctl("cboPersonel", "ComboBox", 112, 46, 240, 20),
    ctl("lblDosya", "Label", 16, 76, 88, 16, cap="Dosya No"),
    ctl("cboDosya", "ComboBox", 112, 74, 240, 20),
    ctl("lblYil", "Label", 16, 104, 88, 16, cap="Kesinti Yılı"),
    ctl("txtYil", "TextBox", 112, 102, 80, 20),
    ctl("lblAy", "Label", 16, 132, 88, 16, cap="Kesinti Ayı"),
    ctl("cboAy", "ComboBox", 112, 130, 120, 20),
    ctl("lblTutar", "Label", 16, 160, 88, 16, cap="Tutar (TL)"),
    ctl("txtTutar", "TextBox", 112, 158, 120, 20),
    ctl("lblAciklama", "Label", 16, 188, 88, 16, cap="Açıklama"),
    ctl("txtAciklama", "TextBox", 112, 186, 240, 20),
    ctl("lblNot", "Label", 16, 214, 388, 26, italic=True, size=8,
        cap="Geçmiş aylar veya elle takip için. Ay sonu otomatik kesintiler AY SONU KESİNTİLERİ ekranından."),
    ctl("cmdKaydet", "CommandButton", 112, 248, 110, 30, cap="Kaydet"),
    ctl("cmdKapat", "CommandButton", 232, 248, 90, 30, cap="Kapat"),
])

# =====================================================================
# 6) İCRA DAİRESİNE ÖDEME
# =====================================================================
form("frmOdemeGiris", "İcra Dairesine Ödeme", 420, 344, [
    ctl("lblTitle", "Label", 0, 0, 420, 32, cap="İCRA DAİRESİNE ÖDEME", size=13),
    ctl("lblPersonel", "Label", 16, 48, 88, 16, cap="Personel"),
    ctl("cboPersonel", "ComboBox", 112, 46, 240, 20),
    ctl("lblDosya", "Label", 16, 76, 88, 16, cap="Dosya No"),
    ctl("cboDosya", "ComboBox", 112, 74, 240, 20),
    ctl("lblBekleyenBilgi", "Label", 112, 98, 250, 16, italic=True),
    ctl("lblTarih", "Label", 16, 122, 88, 16, cap="Ödeme Tarihi"),
    ctl("txtTarih", "TextBox", 112, 120, 90, 20),
    ctl("lblTutar", "Label", 16, 150, 88, 16, cap="Tutar (TL)"),
    ctl("txtTutar", "TextBox", 112, 148, 120, 20),
    ctl("lblDekont", "Label", 16, 178, 88, 16, cap="Banka/Dekont"),
    ctl("txtDekont", "TextBox", 112, 176, 240, 20),
    ctl("lblAciklama", "Label", 16, 206, 88, 16, cap="Açıklama"),
    ctl("txtAciklama", "TextBox", 112, 204, 240, 20),
    ctl("lblNot", "Label", 16, 232, 388, 26, italic=True, size=8,
        cap="Bordrodan kesilen tutarın icra dairesinin hesabına yatırılması burada kaydedilir."),
    ctl("cmdKaydet", "CommandButton", 112, 266, 110, 30, cap="Kaydet"),
    ctl("cmdKapat", "CommandButton", 232, 266, 90, 30, cap="Kapat"),
])

# =====================================================================
# 7) PERSONEL ÖZETİ
# =====================================================================
OZ_COLS = [("Ad Soyad", 140), ("Net Maaş", 70), ("Dosya", 38), ("Toplam İcra", 76),
           ("Tahsil", 76), ("Kalan", 76), ("Ödenen", 70), ("Bekleyen", 70)]
form("frmPersonelOzet", "Personel Özeti", 660, 408, [
    ctl("lblTitle", "Label", 0, 0, 660, 32, cap="PERSONEL ÖZETİ", size=13),
    ctl("lblSayac", "Label", 480, 42, 164, 16, align=3),
    *headers("oz", 16, 62, OZ_COLS),
    ctl("lstOzet", "ListBox", 16, 78, 628, 254, cols=OZ_COLS),
    ctl("lblNot", "Label", 16, 342, 400, 28, italic=True, size=8,
        cap="Tahsil = önceki firma + bordro kesintileri. Bekleyen = kesilen ama daireye henüz ödenmeyen."),
    ctl("cmdYenile", "CommandButton", 440, 342, 90, 28, cap="Yenile"),
    ctl("cmdKapat", "CommandButton", 540, 342, 90, 28, cap="Kapat"),
])

# =====================================================================
# 8) AY SONU KESİNTİLERİ (dönem)
# =====================================================================
form("frmDonem", "Ay Sonu Kesintileri", 400, 290, [
    ctl("lblTitle", "Label", 0, 0, 400, 32, cap="AY SONU KESİNTİLERİ", size=13),
    ctl("lblYil", "Label", 16, 50, 60, 16, cap="Yıl"),
    ctl("txtYil", "TextBox", 80, 48, 80, 20),
    ctl("lblAy", "Label", 180, 50, 30, 16, cap="Ay"),
    ctl("cboAy", "ComboBox", 214, 48, 120, 20),
    ctl("lblOnizleme", "Label", 16, 80, 368, 18, bold=True),
    ctl("lblNot", "Label", 16, 104, 368, 42, italic=True, size=8,
        cap="ÖDEMEDE durumundaki dosyalara net maaş x oran (veya sabit tutar) kadar kesinti işlenir; "
            "dosya kapanırsa artan tutar sıradakine taşar. Aynı aya ikinci kez basarsanız mükerrer kayıt oluşmaz."),
    ctl("cmdHesapla", "CommandButton", 16, 154, 180, 36, cap="HESAPLA ve KAYDET", bold=True),
    ctl("cmdGeriAl", "CommandButton", 204, 154, 180, 36, cap="Seçili Ayı GERİ AL"),
    ctl("cmdKapat", "CommandButton", 150, 200, 100, 28, cap="Kapat"),
])

# =====================================================================
# 9) TARİH SEÇ
# =====================================================================
form("frmTarihSec", "Tarih Seç", 240, 168, [
    ctl("lblTitle", "Label", 0, 0, 240, 26, cap="TARİH SEÇ", size=11),
    ctl("cboGun", "ComboBox", 16, 38, 56, 20),
    ctl("cboAy", "ComboBox", 80, 38, 84, 20),
    ctl("cboYil", "ComboBox", 172, 38, 52, 20),
    ctl("lblSecili", "Label", 16, 66, 208, 16, align=2, bold=True),
    ctl("cmdBugun", "CommandButton", 16, 92, 64, 26, cap="Bugün"),
    ctl("cmdSec", "CommandButton", 88, 92, 72, 26, cap="Seç", bold=True),
    ctl("cmdIptal", "CommandButton", 168, 92, 56, 26, cap="İptal"),
])

os.makedirs("kurulum", exist_ok=True)
with open("kurulum/arayuz.json", "w", encoding="utf-8") as f:
    json.dump({"forms": F}, f, ensure_ascii=False, indent=1)
print(f"kurulum/arayuz.json yazıldı: {len(F)} form, "
      f"{sum(len(x['controls']) for x in F)} kontrol")
