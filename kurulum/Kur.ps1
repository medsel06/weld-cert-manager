# =====================================================================
#  PERSONEL İCRA TAKİP - KURULUM SCRIPTI
#  ---------------------------------------------------------------
#  Excel'i COM ile açar; çalışma kitabını hazırlar (İCRA ÖDEMELERİ
#  sayfası, yeni formül kolonları, durum düzeltmesi), eski form ve
#  modülleri siler, yenilerini arayuz.json + src\*.vba dosyalarından
#  sıfırdan kurar ve dosyayı kaydeder.
#
#  GEREKSİNİM (bir kez): Excel > Dosya > Seçenekler > Güven Merkezi >
#  Güven Merkezi Ayarları > Makro Ayarları >
#  [x] VBA proje nesne modeline erişime güven
#
#  KULLANIM:  Sağ tık > PowerShell ile Çalıştır
#             veya:  powershell -ExecutionPolicy Bypass -File Kur.ps1
#             Farklı dosya:  ... -File Kur.ps1 -Dosya "C:\yol\dosya.xlsm"
# =====================================================================
param(
    [string]$Dosya = ""
)

$ErrorActionPreference = "Stop"
$kok = Split-Path -Parent $MyInvocation.MyCommand.Path

function Yaz([string]$m, [string]$renk = "Gray") { Write-Host $m -ForegroundColor $renk }

# ---------------------------------------------------------------- girdiler
if ([string]::IsNullOrWhiteSpace($Dosya)) {
    $Dosya = Join-Path $kok "Personel_Icra_Takip_Profesyonel.xlsm"
}
if (-not (Test-Path $Dosya)) {
    Yaz "HATA: Excel dosyası bulunamadı: $Dosya" "Red"
    Yaz "Dosyayı bu klasöre koyun veya -Dosya parametresiyle yol verin." "Yellow"
    Read-Host "Kapatmak için Enter"
    exit 1
}
$Dosya = (Resolve-Path $Dosya).Path

$jsonYol = Join-Path $kok "arayuz.json"
$srcKlasor = Join-Path $kok "src"
foreach ($p in @($jsonYol, $srcKlasor)) {
    if (-not (Test-Path $p)) {
        Yaz "HATA: Kurulum dosyası eksik: $p" "Red"
        Read-Host "Kapatmak için Enter"
        exit 1
    }
}
$arayuz = Get-Content -Raw -Encoding UTF8 $jsonYol | ConvertFrom-Json

function KodOku([string]$ad) {
    $yol = Join-Path $srcKlasor ($ad + ".vba")
    if (-not (Test-Path $yol)) { throw "Kaynak kod eksik: $yol" }
    return (Get-Content -Raw -Encoding UTF8 $yol)
}

# ---------------------------------------------------------------- yedek
$yedek = [System.IO.Path]::ChangeExtension($Dosya, $null).TrimEnd('.') + `
         "_yedek_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm"
Copy-Item $Dosya $yedek
Yaz "Yedek alındı: $yedek" "DarkGray"

# ---------------------------------------------------------------- excel
Yaz "Excel başlatılıyor..." "Cyan"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.EnableEvents = $false
# KRİTİK: makrolar çalışmadan aç (Workbook_Open formu açıp scripti kilitlemesin)
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

$wb = $null
try {
    $wb = $excel.Workbooks.Open($Dosya)

    # VBA proje erişimi kontrolü
    $project = $null
    try { $project = $wb.VBProject } catch { }
    if ($null -eq $project) {
        throw ("VBA projesine erişilemiyor. Excel'de şu ayarı açın ve scripti yeniden çalıştırın:`n" +
               "Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları >`n" +
               "[x] VBA proje nesne modeline erişime güven")
    }

    # ============================================================ 1) ÇALIŞMA KİTABI HAZIRLIĞI
    Yaz "1/4 Çalışma kitabı hazırlanıyor (ödeme sayfası + formüller)..." "Cyan"

    $LACIVERT = 7884319    # RGB(31,78,120)
    $TURUNCU = 1137094     # RGB(198,89,17)
    $BEYAZ = 16777215

    # İCRA ÖDEMELERİ sayfası
    $odemeVar = $false
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "İCRA ÖDEMELERİ") { $odemeVar = $true } }
    if (-not $odemeVar) {
        $wsOnce = $wb.Worksheets.Item("AYLIK KESİNTİLER")
        $wsO = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsOnce)
        $wsO.Name = "İCRA ÖDEMELERİ"
        $basliklar = @("Personel Adı Soyadı", "Dosya Numarası", "Ödeme Tarihi",
                       "Ödeme Tutarı (TL)", "Banka / Dekont No", "Açıklama")
        for ($i = 0; $i -lt $basliklar.Count; $i++) {
            $h = $wsO.Cells.Item(1, $i + 1)
            $h.Value2 = $basliklar[$i]
            $h.Interior.Color = $LACIVERT
            $h.Font.Color = $BEYAZ
            $h.Font.Bold = $true
        }
        $wsO.Rows.Item(1).RowHeight = 28
        $genislikler = @(24, 16, 12, 14, 18, 30)
        for ($i = 0; $i -lt $genislikler.Count; $i++) {
            $wsO.Columns.Item($i + 1).ColumnWidth = $genislikler[$i]
        }
        $wsO.Range("C2:C5000").NumberFormat = "dd.mm.yyyy"
        $wsO.Range("D2:D5000").NumberFormat = "#,##0.00"
        Yaz "   + İCRA ÖDEMELERİ sayfası eklendi" "Green"
    } else {
        Yaz "   = İCRA ÖDEMELERİ sayfası zaten var" "DarkGray"
    }

    # İCRA DOSYALARI: W/X kolonları + durum formülü düzeltmesi (SIRADA)
    $wsI = $wb.Worksheets.Item("İCRA DOSYALARI")
    $wsI.Cells.Item(1, 23).Value2 = "Daireye Ödenen (TL)"
    $wsI.Cells.Item(1, 24).Value2 = "Ödenmesi Bekleyen (TL)"
    foreach ($k in @(23, 24)) {
        $h = $wsI.Cells.Item(1, $k)
        $h.Interior.Color = $TURUNCU
        $h.Font.Color = $BEYAZ
        $h.Font.Bold = $true
        $wsI.Columns.Item($k).ColumnWidth = 14
    }
    $wsI.Range("W2:W300").FormulaR1C1 = '=IF(RC5="","",SUMIFS(''İCRA ÖDEMELERİ''!R2C4:R5000C4,''İCRA ÖDEMELERİ''!R2C1:R5000C1,RC1,''İCRA ÖDEMELERİ''!R2C2:R5000C2,RC5))'
    $wsI.Range("X2:X300").FormulaR1C1 = '=IF(RC5="","",RC16-RC23)'
    $wsI.Range("W2:X300").NumberFormat = "#,##0.00"
    $wsI.Range("S2:S300").FormulaR1C1 = '=IF(RC5="","",IF(RC18<=0,"KAPANDI",IF(COUNTIFS(R2C1:R300C1,RC1,R2C6:R300C6,"<"&RC6,R2C18:R300C18,">0")=0,"ÖDEMEDE","SIRADA")))'
    Yaz "   + W/X kolonları yazıldı, DURUM formülü SIRADA olarak düzeltildi" "Green"

    # ============================================================ 2) ESKİ BİLEŞENLERİ KALDIR
    Yaz "2/4 Eski form ve modüller kaldırılıyor..." "Cyan"
    $silinecek = @()
    foreach ($comp in $project.VBComponents) {
        # 1 = modül, 3 = UserForm; belge modüllerine (100) dokunma
        if ($comp.Type -eq 3) { $silinecek += $comp }
        elseif ($comp.Type -eq 1 -and @("IcraTakip", "IcraTakipUI", "FormStil") -contains $comp.Name) {
            $silinecek += $comp
        }
    }
    foreach ($comp in $silinecek) {
        $ad = $comp.Name
        $project.VBComponents.Remove($comp)
        Yaz "   - $ad kaldırıldı" "DarkGray"
    }

    # ============================================================ 3) MODÜLLER
    Yaz "3/4 Modüller yükleniyor..." "Cyan"
    foreach ($modAd in @("IcraTakip", "IcraTakipUI", "FormStil")) {
        $comp = $project.VBComponents.Add(1)   # vbext_ct_StdModule
        $comp.Name = $modAd
        $comp.CodeModule.AddFromString((KodOku $modAd))
        Yaz "   + $modAd" "Green"
    }
    # ThisWorkbook (BuÇalışmaKitabı) kodu
    $twb = $project.VBComponents.Item($wb.CodeName)
    if ($twb.CodeModule.CountOfLines -gt 0) {
        $twb.CodeModule.DeleteLines(1, $twb.CodeModule.CountOfLines)
    }
    $twb.CodeModule.AddFromString((KodOku "BuCalismaKitabi"))
    Yaz "   + ThisWorkbook olayları" "Green"

    # ============================================================ 4) FORMLAR
    Yaz "4/4 Formlar kuruluyor..." "Cyan"
    foreach ($f in $arayuz.forms) {
        $comp = $project.VBComponents.Add(3)   # vbext_ct_MSForm
        $comp.Name = $f.name
        $comp.Properties.Item("Caption").Value = $f.caption
        $comp.Properties.Item("Width").Value = [double]$f.w + 8
        $comp.Properties.Item("Height").Value = [double]$f.h + 30

        $d = $comp.Designer
        foreach ($ctl in $f.controls) {
            $c = $d.Controls.Add("Forms." + $ctl.t + ".1", $ctl.n, $true)
            $c.Left = [single]$ctl.x
            $c.Top = [single]$ctl.y
            $c.Width = [single]$ctl.w
            $c.Height = [single]$ctl.h
            if ($null -ne $ctl.PSObject.Properties["cap"]) { $c.Caption = $ctl.cap }
            if ($null -ne $ctl.PSObject.Properties["size"]) { $c.Font.Size = [single]$ctl.size }
            if ($null -ne $ctl.PSObject.Properties["bold"]) { $c.Font.Bold = $true }
            if ($null -ne $ctl.PSObject.Properties["italic"]) { $c.Font.Italic = $true }
            if ($null -ne $ctl.PSObject.Properties["align"]) { $c.TextAlign = [int]$ctl.align }
            if ($null -ne $ctl.PSObject.Properties["locked"]) { $c.Locked = $true }
            if ($null -ne $ctl.PSObject.Properties["colCount"]) {
                $c.ColumnCount = [int]$ctl.colCount
                $c.ColumnWidths = $ctl.colWidths
            }
        }
        $comp.CodeModule.AddFromString((KodOku $f.name))
        Yaz ("   + {0}  ({1} kontrol)" -f $f.name, $f.controls.Count) "Green"
    }

    # ============================================================ kaydet
    $wb.Save()
    $wb.Close($false)
    $excel.EnableEvents = $true
    Yaz ""
    Yaz "KURULUM TAMAMLANDI: $Dosya" "Green"
    Yaz "Dosyayı Excel'de açın; program otomatik başlar (İçeriği Etkinleştir gerekebilir)." "Yellow"
}
catch {
    Yaz ""
    Yaz ("HATA: " + $_.Exception.Message) "Red"
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    Yaz "Dosyanız değişmediyse de yedeğiniz hazır: $yedek" "Yellow"
}
finally {
    try { $excel.Quit() } catch { }
    foreach ($o in @($wb, $excel)) {
        if ($null -ne $o) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Read-Host "Kapatmak için Enter"
}
