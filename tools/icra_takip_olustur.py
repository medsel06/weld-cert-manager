# -*- coding: utf-8 -*-
"""
Personel İcra Takip Excel şablonu üretici.

Zirve Müşavir "Personel İcra Bilgi Girişi" ekranındaki mantığı Excel'e taşır:
  - Bir personelin birden fazla icra dosyası öncelik sırasıyla (1, 2, 3...) izlenir.
  - Aylık bordro kesintileri ayrı sayfaya girilir, dosya bazında toplanır.
  - Tahsil Edilen / Kalan Borç / Durum (ÖDEMEDE - SIRADA - KAPANDI) otomatik hesaplanır:
    kalan borcu biten dosya KAPANDI olur, sıradaki en küçük öncelikli açık dosya
    kendiliğinden ÖDEMEDE konumuna geçer.

Kullanım:  python3 tools/icra_takip_olustur.py [çıktı.xlsx]
"""
import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

DOSYA_SON_SATIR = 200      # İCRA DOSYALARI sayfasında formül hazır son satır
KESINTI_SON_SATIR = 5000   # AYLIK KESİNTİLER sayfasında veri aralığı
OZET_SON_SATIR = 100

LACIVERT = "1F4E78"   # elle girilen kolon başlığı
TURUNCU = "C65911"    # otomatik (formüllü) kolon başlığı
ACIK_GRI = "F2F2F2"

INCE = Side(style="thin", color="BFBFBF")
KENARLIK = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)
PARA = "#,##0.00"
TARIH = "DD.MM.YYYY"


def baslik_yaz(ws, basliklar, otomatikler=()):
    for i, ad in enumerate(basliklar, start=1):
        c = ws.cell(row=1, column=i, value=ad)
        renk = TURUNCU if i in otomatikler else LACIVERT
        c.fill = PatternFill("solid", start_color=renk, end_color=renk)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = KENARLIK
    ws.row_dimensions[1].height = 32


def genislik(ws, genislikler):
    for i, w in enumerate(genislikler, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def durum_boyamasi(ws, aralik):
    kurallar = [
        ("ÖDEMEDE", "C6EFCE", "006100"),
        ("SIRADA", "FFEB9C", "9C6500"),
        ("KAPANDI", "D9D9D9", "595959"),
    ]
    for metin, zemin, yazi in kurallar:
        ws.conditional_formatting.add(
            aralik,
            CellIsRule(
                operator="equal",
                formula=[f'"{metin}"'],
                fill=PatternFill("solid", start_color=zemin, end_color=zemin),
                font=Font(bold=True, color=yazi),
            ),
        )


def main(yol):
    wb = Workbook()

    # ------------------------------------------------------------- KULLANIM
    ws = wb.active
    ws.title = "KULLANIM"
    ws.sheet_properties.tabColor = "808080"
    ws.column_dimensions["A"].width = 118

    satirlar = [
        ("PERSONEL İCRA TAKİP DOSYASI", "baslik"),
        ("Bir personelin birden fazla icra dosyasını öncelik sırasına göre izler; "
         "aylık kesintileri toplar, kalan borcu ve hangi dosyanın ödemede olduğunu kendisi hesaplar.", None),
        ("", None),
        ("1) İCRA DOSYALARI sayfasına her icra dosyasını AYRI SATIR olarak girin. "
         "Aynı personelin 5 icrası varsa 5 satır açın ve 'Öncelik Sırası' kolonuna 1, 2, 3, 4, 5 yazın.", None),
        ("2) Her ay bordrodan yapılan icra kesintisini AYLIK KESİNTİLER sayfasına bir satır olarak girin. "
         "Personel adı ve dosya numarası, İCRA DOSYALARI sayfasındakiyle BİREBİR AYNI yazılmalıdır "
         "(personel adı hücresinde hazır açılır liste vardır).", None),
        ("3) Hesaplamalar otomatiktir: 'Bordro Kesintileri Toplamı' aylık kesintilerden gelir; "
         "Kalan Borç = İcra Tutarı - (Önceki Firmada Tahsil + Bordro Kesintileri).", None),
        ("4) DURUM kolonu kendiliğinden değişir:  KAPANDI = kalan borç 0'a indi;  "
         "ÖDEMEDE = açık dosyalar içinde önceliği en küçük olan;  SIRADA = öndeki dosya kapanınca ödemeye girecek olan. "
         "Yani 1. sıradaki icra bitince 2. sıradaki kendiliğinden ÖDEMEDE konumuna geçer.", None),
        ("5) Personel önceki firmasında/şubesinde aynı dosyadan kesinti yaptırdıysa bu tutarı "
         "'Önceki Firmada Tahsil' kolonuna yazın.", None),
        ("6) PERSONEL ÖZET sayfası personel bazında toplamları ve şu an ödemede olan dosyayı gösterir. "
         "Yeni personel için adı A kolonuna yazmanız yeterlidir.", None),
        ("", None),
        ("RENK REHBERİ", "baslik2"),
        ("LACİVERT başlıklı kolonlar elle doldurulur.", "lacivert"),
        ("TURUNCU başlıklı kolonlar formüllüdür; bu kolonlara elle veri GİRMEYİN, silmeyin.", "turuncu"),
        ("Formüller 200 dosya satırı / 5.000 kesinti satırı için hazırdır. Daha fazlası gerekirse "
         "son formüllü satırı seçip aşağı kopyalayın.", None),
        ("", None),
        ("DURUM RENKLERİ:   YEŞİL = ÖDEMEDE     SARI = SIRADA     GRİ = KAPANDI", None),
    ]
    for r, (metin, stil) in enumerate(satirlar, start=2):
        c = ws.cell(row=r, column=1, value=metin)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if stil == "baslik":
            c.font = Font(bold=True, size=14, color=LACIVERT)
        elif stil == "baslik2":
            c.font = Font(bold=True, size=11, color=LACIVERT)
        elif stil == "lacivert":
            c.fill = PatternFill("solid", start_color=LACIVERT, end_color=LACIVERT)
            c.font = Font(bold=True, color="FFFFFF")
        elif stil == "turuncu":
            c.fill = PatternFill("solid", start_color=TURUNCU, end_color=TURUNCU)
            c.font = Font(bold=True, color="FFFFFF")
        ws.row_dimensions[r].height = 34 if len(metin) > 90 else 20

    # ------------------------------------------------------- İCRA DOSYALARI
    icra = wb.create_sheet("İCRA DOSYALARI")
    icra.sheet_properties.tabColor = LACIVERT
    basliklar = [
        "Personel Adı Soyadı", "T.C. Kimlik No", "Dosya Tarihi", "Dosya Numarası",
        "Öncelik Sırası", "İcra Dairesi", "Alacaklı", "İcra Tutarı (TL)",
        "Kesinti Şekli", "Kesinti Oranı (örn. 1/4)", "Aylık Kesinti Tutarı (TL)",
        "Önceki Firmada Tahsil (TL)", "Bordro Kesintileri Toplamı (TL)",
        "Toplam Tahsil Edilen (TL)", "Kalan Borç (TL)", "DURUM",
        "Ödeme Yapılacak Banka / IBAN", "Açıklama", "(otomatik - silmeyin)",
    ]
    baslik_yaz(icra, basliklar, otomatikler=(13, 14, 15, 16, 19))
    genislik(icra, [24, 14, 12, 16, 9, 30, 34, 14, 12, 12, 13, 13, 14, 14, 14, 12, 28, 24, 4])
    icra.freeze_panes = "C2"
    icra.auto_filter.ref = f"A1:R{DOSYA_SON_SATIR}"
    icra.column_dimensions["S"].hidden = True

    son = DOSYA_SON_SATIR
    for r in range(2, son + 1):
        icra.cell(row=r, column=13).value = (
            f'=IF($D{r}="","",SUMIFS(\'AYLIK KESİNTİLER\'!$E$2:$E${KESINTI_SON_SATIR},'
            f"'AYLIK KESİNTİLER'!$A$2:$A${KESINTI_SON_SATIR},$A{r},"
            f"'AYLIK KESİNTİLER'!$B$2:$B${KESINTI_SON_SATIR},$D{r}))"
        )
        icra.cell(row=r, column=14).value = f'=IF($D{r}="","",$L{r}+$M{r})'
        icra.cell(row=r, column=15).value = f'=IF($D{r}="","",$H{r}-$N{r})'
        icra.cell(row=r, column=16).value = (
            f'=IF($D{r}="","",IF($O{r}<=0,"KAPANDI",'
            f'IF(COUNTIFS($A$2:$A${son},$A{r},$E$2:$E${son},"<"&$E{r},$O$2:$O${son},">0")=0,'
            f'"ÖDEMEDE","SIRADA")))'
        )
        icra.cell(row=r, column=19).value = f'=IF($D{r}="","",$A{r}&"|"&$P{r})'
        for kol in (8, 11, 12, 13, 14, 15):
            icra.cell(row=r, column=kol).number_format = PARA
        icra.cell(row=r, column=2).number_format = "@"
        icra.cell(row=r, column=3).number_format = TARIH
        for kol in range(1, 19):
            icra.cell(row=r, column=kol).border = KENARLIK
        icra.cell(row=r, column=16).alignment = Alignment(horizontal="center")

    durum_boyamasi(icra, f"P2:P{son}")

    # Örnek kayıtlar (ekran görüntülerindeki MERT TÜZE verisiyle birebir)
    ornek_dosyalar = [
        ["MERT TÜZE", "38128425238", datetime(2026, 2, 23), "2025/203385", 1,
         "BANKA ALACAKLARI İCRA DAİRESİ", "AKBANK TÜRK ANONİM ŞİRKETİ",
         130656.65, "Oran", "1/4", None, 0, None, "ÖRNEK KAYIT - silebilirsiniz"],
        ["MERT TÜZE", "38128425238", datetime(2026, 6, 10), "deneme", 2,
         "", "", 10000.00, "Sabit Tutar", "", 10000.00, 0, None, "ÖRNEK KAYIT - silebilirsiniz"],
        ["AHMET BAŞOĞLU", "12856239602", datetime(2026, 1, 15), "2024/101010", 1,
         "ANKARA 5. İCRA DAİRESİ", "X FİNANS A.Ş.", 5000.00, "Sabit Tutar", "",
         2500.00, 5000.00, None, "ÖRNEK: borç bitti, otomatik KAPANDI"],
        ["AHMET BAŞOĞLU", "12856239602", datetime(2026, 2, 21), "2025/55555", 2,
         "ANKARA 12. İCRA DAİRESİ", "Y BANKASI A.Ş.", 24000.00, "Oran", "1/4",
         None, 0, None, "ÖRNEK: 1. dosya kapanınca ÖDEMEDE'ye geçti"],
        ["AHMET BAŞOĞLU", "12856239602", datetime(2026, 3, 1), "2026/77777", 3,
         "ANKARA 3. İCRA DAİRESİ", "Z TELEKOM A.Ş.", 8000.00, "Oran", "1/4",
         None, 0, None, "ÖRNEK: sırada bekliyor"],
    ]
    for r, satir in enumerate(ornek_dosyalar, start=2):
        (ad, tc, tarih, dosya, oncelik, daire, alacakli,
         tutar, sekil, oran, aylik, onceki, banka, aciklama) = satir
        degerler = {1: ad, 2: tc, 3: tarih, 4: dosya, 5: oncelik, 6: daire, 7: alacakli,
                    8: tutar, 9: sekil, 10: oran, 11: aylik, 12: onceki, 17: banka, 18: aciklama}
        for kol, deger in degerler.items():
            if deger is not None:
                icra.cell(row=r, column=kol, value=deger)

    # Doğrulamalar
    dv_sekil = DataValidation(type="list", formula1='"Oran,Sabit Tutar"', allow_blank=True)
    dv_sekil.prompt = "Kesinti maaşın oranı mı (örn. 1/4) yoksa sabit tutar mı?"
    icra.add_data_validation(dv_sekil)
    dv_sekil.add(f"I2:I{son}")

    dv_oncelik = DataValidation(type="whole", operator="greaterThanOrEqual",
                                formula1="1", allow_blank=True)
    dv_oncelik.errorTitle = "Öncelik Sırası"
    dv_oncelik.error = "1 veya daha büyük tam sayı girin (1 = ilk kesilecek dosya)."
    icra.add_data_validation(dv_oncelik)
    dv_oncelik.add(f"E2:E{son}")

    # ------------------------------------------------------ AYLIK KESİNTİLER
    kes = wb.create_sheet("AYLIK KESİNTİLER")
    kes.sheet_properties.tabColor = "2E75B6"
    baslik_yaz(kes, ["Personel Adı Soyadı", "Dosya Numarası", "Kesinti Yılı",
                     "Kesinti Dönemi (Ay)", "Kesinti Tutarı (TL)", "Açıklama"])
    genislik(kes, [24, 16, 11, 15, 15, 30])
    kes.freeze_panes = "A2"
    kes.auto_filter.ref = f"A1:F{KESINTI_SON_SATIR}"

    for r in range(2, KESINTI_SON_SATIR + 1):
        kes.cell(row=r, column=5).number_format = PARA
        if r <= 300:  # kenarlıkları ilk 300 satıra çiz, dosya boyutu şişmesin
            for kol in range(1, 7):
                kes.cell(row=r, column=kol).border = KENARLIK

    ornek_kesintiler = [
        ["MERT TÜZE", "2025/203385", 2026, "Şubat", 7018.88, "ÖRNEK KAYIT"],
        ["MERT TÜZE", "2025/203385", 2026, "Mart", 7018.88, "ÖRNEK KAYIT"],
        ["MERT TÜZE", "2025/203385", 2026, "Nisan", 7018.88, "ÖRNEK KAYIT"],
        ["AHMET BAŞOĞLU", "2025/55555", 2026, "Haziran", 6000.00, "ÖRNEK KAYIT"],
    ]
    for r, satir in enumerate(ornek_kesintiler, start=2):
        for kol, deger in enumerate(satir, start=1):
            kes.cell(row=r, column=kol, value=deger)

    wb.defined_names["PersonelListesi"] = DefinedName(
        "PersonelListesi", attr_text=f"'İCRA DOSYALARI'!$A$2:$A${DOSYA_SON_SATIR}")

    dv_personel = DataValidation(type="list", formula1="PersonelListesi", allow_blank=True)
    dv_personel.errorStyle = "warning"
    dv_personel.errorTitle = "Personel adı"
    dv_personel.error = "İCRA DOSYALARI sayfasındaki yazımla birebir aynı olmalı."
    kes.add_data_validation(dv_personel)
    dv_personel.add(f"A2:A{KESINTI_SON_SATIR}")

    aylar = "Ocak,Şubat,Mart,Nisan,Mayıs,Haziran,Temmuz,Ağustos,Eylül,Ekim,Kasım,Aralık"
    dv_ay = DataValidation(type="list", formula1=f'"{aylar}"', allow_blank=True)
    kes.add_data_validation(dv_ay)
    dv_ay.add(f"D2:D{KESINTI_SON_SATIR}")

    # --------------------------------------------------------- PERSONEL ÖZET
    ozet = wb.create_sheet("PERSONEL ÖZET")
    ozet.sheet_properties.tabColor = "548235"
    baslik_yaz(ozet, ["Personel Adı Soyadı", "Dosya Sayısı", "Toplam İcra Tutarı (TL)",
                      "Toplam Tahsil Edilen (TL)", "Toplam Kalan Borç (TL)",
                      "Ödemedeki Dosya No", "Ödemedeki Dosyanın Kalan Borcu (TL)"],
               otomatikler=(2, 3, 4, 5, 6, 7))
    genislik(ozet, [24, 11, 17, 18, 17, 17, 19])
    ozet.freeze_panes = "A2"

    d = f"'İCRA DOSYALARI'!$A$2:$A${DOSYA_SON_SATIR}"
    for r in range(2, OZET_SON_SATIR + 1):
        ozet.cell(row=r, column=2).value = f'=IF($A{r}="","",COUNTIF({d},$A{r}))'
        for kol, kaynak in ((3, "$H"), (4, "$N"), (5, "$O")):
            ozet.cell(row=r, column=kol).value = (
                f'=IF($A{r}="","",SUMIF({d},$A{r},'
                f"'İCRA DOSYALARI'!{kaynak}$2:{kaynak}${DOSYA_SON_SATIR}))"
            )
            ozet.cell(row=r, column=kol).number_format = PARA
        esle = (f'MATCH($A{r}&"|ÖDEMEDE",\'İCRA DOSYALARI\'!$S$2:$S${DOSYA_SON_SATIR},0)')
        ozet.cell(row=r, column=6).value = (
            f'=IF($A{r}="","",IFERROR(INDEX(\'İCRA DOSYALARI\'!$D$2:$D${DOSYA_SON_SATIR},{esle}),"-"))'
        )
        ozet.cell(row=r, column=7).value = (
            f'=IF($A{r}="","",IFERROR(INDEX(\'İCRA DOSYALARI\'!$O$2:$O${DOSYA_SON_SATIR},{esle}),"-"))'
        )
        ozet.cell(row=r, column=7).number_format = PARA
        for kol in range(1, 8):
            ozet.cell(row=r, column=kol).border = KENARLIK

    ozet.cell(row=2, column=1, value="MERT TÜZE")
    ozet.cell(row=3, column=1, value="AHMET BAŞOĞLU")

    wb.save(yol)
    print(f"Yazıldı: {yol}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Personel_Icra_Takip.xlsx")
