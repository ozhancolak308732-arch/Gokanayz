'==============================================================================
' Mernis Chrome Otomasyon Modulu
' clsMernisKPS class module'unu kullanir
'==============================================================================

Option Explicit

'==============================================================================
' YARDIMCI: Sorgulanacak TC Kimlik No'yu belirle
'==============================================================================
Private Function SorgulanacakTC(textBoxDeger As String) As String
    If textBoxDeger <> "" Then
        SorgulanacakTC = textBoxDeger
    ElseIf Sheet1.Range("B4").Value <> "" Then
        SorgulanacakTC = CStr(Sheet1.Range("B4").Value)
    ElseIf Sheet1.Range("B19").Value <> "" Then
        SorgulanacakTC = CStr(Sheet1.Range("B19").Value)
    Else
        SorgulanacakTC = ""
    End If
End Function

'==============================================================================
' 1. MERNIS ADRES AL - Chrome
'==============================================================================
Sub MernisAdresAl_Chrome()
    Dim chrome As New clsBrowser
    Dim kps As New clsMernisKPS
    Dim hataMesaji As String
    Dim tcNo As String
    Dim kayitHata As String

    ' TC No belirle
    tcNo = SorgulanacakTC(UserForm1.TextBox1.Value)
    If tcNo = "" Then Exit Sub

    On Error Resume Next

    Set kps.Browser = chrome

1:
    ' Chrome baslat ve giris yap
    If Not kps.BaslatVeGirisYap( _
            UserForm1.txtkullanici_adi.Text, _
            Trim(UserForm1.txtSifre.Text)) Then
        GoTo 1
    End If

    ' TC sorgusu yap
    hataMesaji = kps.TCKNSorgula(tcNo)

    ' Hata varsa (hataMesaji <> "1" demek hata var)
    If Not hataMesaji = "1" Then
        kps.HataDialogKapat

        ' Yabanci uyruklu kontrolu (TC 99 ile basliyorsa)
        If Left(tcNo, 2) = "99" Then
            kps.YabanciKimlikSorgula tcNo
            kps.YabanciKisiBilgileriniOku False
            GoTo FormaYaz
        End If

        ' Normal hata - mesaji goster
        UserForm1.TextBox1.Text = hataMesaji
        GoTo HataGoster
    End If

    Application.Wait Now + TimeValue("0:00:01")
    kps.KisiBilgileriniOku False

FormaYaz:
    ' Sonuclari forma ve sayfaya yaz
    UserForm1.TextBox1.Text = kps.TC
    Sayfa8.Range("L6").Value = kps.TC
    UserForm1.TextBox2.Text = kps.Ad & Chr(32) & kps.Soyad
    Sayfa8.Range("L5").Value = kps.Ad & Chr(32) & kps.Soyad
    UserForm1.TextBox3.Text = kps.Adres
    Sayfa8.Range("L7").Value = kps.Adres

    ' Kayit bulunamadi kontrolu
    kayitHata = kps.KayitBulunamadiMi()
    If kayitHata <> "" Then
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            kps.YabanciNavigasyonTikla
        End If
        UserForm1.TextBox3.Text = kayitHata
    End If
    Exit Sub

HataGoster:
    MsgBox hataMesaji, vbInformation, "Coder By Ozhan COLAK"
    chrome.maximized
End Sub

'==============================================================================
' 2. MERNIS HASTA DOGUM/OLUM TARIHI AL - Chrome
'==============================================================================
Sub MernisHastaDogumTarihiOlumTarihiAl_Chrome()
    Dim chrome As New clsBrowser
    Dim kps As New clsMernisKPS
    Dim hataMesaji As String
    Dim tcNo As String
    Dim kayitHata As String
    Dim kullaniciAdi As String
    Dim sifre As String

    ' TC No belirle (sadece B4'ten)
    If Sheet1.Range("B4").Value = "" Then Exit Sub
    tcNo = CStr(Sheet1.Range("B4").Value)

    ' Kullanici bilgilerini belirle (form aciksa formdan, degilse hucrelerden)
    If UserForm1.visible = True Then
        kullaniciAdi = UserForm1.txtkullanici_adi.Text
        sifre = UserForm1.txtSifre.Text
    Else
        kullaniciAdi = CStr(Range("AX1").Value)
        sifre = CStr(Range("AY1").Value)
    End If

    On Error Resume Next

    Set kps.Browser = chrome

    ' Chrome baslat ve giris yap
    If Not kps.BaslatVeGirisYap(kullaniciAdi, Trim(sifre)) Then
        MsgBox "Chrome baslatilamadi!", vbExclamation, "Coder By Ozhan COLAK"
        Exit Sub
    End If

    ' TC sorgusu yap
    hataMesaji = kps.TCKNSorgula(tcNo)

    ' Hata varsa
    If Not hataMesaji = "1" Then
        kps.HataDialogKapat

        ' Yabanci uyruklu kontrolu (TC 99 ile basliyorsa)
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            kps.YabanciKimlikSorgula UserForm1.TextBox1.Value
            kps.YabanciKisiBilgileriniOku True
            GoTo FormaYaz
        End If

        ' Normal hata
        UserForm1.TextBox1.Text = hataMesaji
        GoTo HataGoster
    End If

    Application.Wait Now + TimeValue("0:00:01")
    kps.KisiBilgileriniOku True

FormaYaz:
    ' Formu temizle
    iadeFormTemizle

    ' Hasta bilgilerini forma yaz
    UserForm1.LabelHastaBilgi.Caption = "HASTA BILGILERI" & _
        Chr(10) & kps.TC & _
        Chr(10) & kps.Ad & Chr(32) & kps.Soyad & _
        Chr(10) & kps.DogumTarihi & _
        Chr(10) & kps.OlumTarihi

    ' Sayfa ve hucrelere yaz
    Sayfa8.Range("L4").Value = kps.TC
    Sayfa8.Range("L3").Value = kps.Ad & Chr(32) & kps.Soyad

    Application.EnableEvents = False
    Sheet1.Range("B22").Value = kps.DogumTarihi
    Sheet1.Range("B23").Value = kps.OlumTarihi
    Application.EnableEvents = True

    ' Kayit bulunamadi kontrolu
    kayitHata = kps.KayitBulunamadiMi()
    If kayitHata <> "" Then
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            kps.AdresCheckboxTikla
        End If
        UserForm1.LabelHastaBilgi.Caption = UserForm1.LabelHastaBilgi.Caption & Chr(10) & kayitHata
    End If
    Exit Sub

HataGoster:
    MsgBox hataMesaji, vbInformation, "Coder By Ozhan COLAK"
End Sub
