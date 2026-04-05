'==============================================================================
' Coder By Ozhan COLAK
' Mernis Chrome Otomasyon Modulu - Optimize Edilmis Versiyon
'
' Degisiklikler:
'   - Ortak kod helper fonksiyonlara cikarildi (ChromeBaslat, KPSGirisYap, vb.)
'   - Syntax hatasi duzeltildi (document.querySelector -> .jsEval)
'   - Degisken tanimlama hatalari duzeltildi (her biri As String)
'   - B19/B4 ters mantik duzeltildi
'   - GoTo spaghetti kaldirildi, yapisal akis saglandi
'   - Gereksiz yorumlar temizlendi
'   - On Error Resume Next daraltildi
'==============================================================================

Option Explicit

Private Const KPS_LOGIN_URL As String = "https://kps.sgk.intra/KPS/Login.aspx"
Private Const KPS_SORGU_URL As String = "https://kps.sgk.intra/KPS/TCKNodanSorgula.aspx"
Private Const WAIT_SHORT As String = "0:00:01"
Private Const WAIT_LONG As String = "0:00:03"

'==============================================================================
' YARDIMCI: Chrome baslatip KPS'ye giris yap
' Basarili ise True dondurur
'==============================================================================
Private Function ChromeBaslatVeGirisYap(chrome As clsBrowser, _
                                         kullaniciAdi As String, _
                                         sifre As String) As Boolean
    Dim jscd As String

    ChromeBaslatVeGirisYap = False

    ' Onceki Chrome'u kapat
    CreateObject("wscript.shell").Run "cmd /c """ & "taskkill /IM chrome.exe >nul""""", 0, True

    With chrome
        .minimized
        .start cleanActiveSession:=True, userProfile:="User G"

        ' Session kontrolu
        If .SessionID = "" Then Exit Function

        ' KPS Login sayfasina git
        .navigate KPS_LOGIN_URL
        .wait

        ' SSL sertifika uyarisini gec
        SSLUyarisiGec chrome

        ' Giris yap
        jscd = "document.getElementsByName('txtUser')[0].value = '" & kullaniciAdi & "';" & _
               "document.getElementsByName('txtPass')[0].value = '" & sifre & "';" & _
               "document.getElementsByName('btnLogin')[0].click();"
        .jsEval jscd
        .wait

        ' Sorgu sayfasina git
        .navigate KPS_SORGU_URL
        .wait

        ' Adres bilgisi checkbox'ini tikla
        .jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"
    End With

    ChromeBaslatVeGirisYap = True
End Function

'==============================================================================
' YARDIMCI: SSL sertifika uyarisini gec
'==============================================================================
Private Sub SSLUyarisiGec(chrome As clsBrowser)
    Dim basla As Single
    basla = Timer
    Do While (Timer - basla) < 1: Loop

    With chrome
        On Error Resume Next
        .jsEval "document.querySelector('#details-button').click()"
        .wait
        .jsEval "document.querySelector('#proceed-link').click()"
        .wait
        On Error GoTo 0
    End With
End Sub

'==============================================================================
' YARDIMCI: TC Kimlik No ile sorgu yap
' hataMesaji parametresine hata dondurur (bos ise basarili)
'==============================================================================
Private Function TCKNSorgula(chrome As clsBrowser, _
                              tcKimlikNo As String) As String
    Dim jscd As String

    With chrome
        jscd = "document.getElementsByName('ctl00$MainContent$txtbox_TCKNo')[0].value = '" & tcKimlikNo & "';" & _
               "document.getElementById('MainContent_ddlist_NkoTip')[0].value = '2';" & _
               "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
        .jsEval jscd

        Application.Wait Now + TimeValue(WAIT_LONG)

        TCKNSorgula = Trim(.jsEval("document.getElementById('uppnl_HataMesaj').innerText.trim()"))
    End With
End Function

'==============================================================================
' YARDIMCI: Yabanci kimlik ile sorgu yap (TC 99 ile basliyorsa)
'==============================================================================
Private Sub YabanciKimlikSorgula(chrome As clsBrowser, yabanciNo As String)
    Dim jscd As String

    With chrome
        .jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
        .wait

        jscd = "document.getElementsByName('ctl00$MainContent$txtbox_YabanciKNo')[0].value = '" & yabanciNo & "';" & _
               "document.getElementById('MainContent_chkbox_AdresBilgisi').click();" & _
               "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
        .jsEval jscd

        Application.Wait Now + TimeValue(WAIT_LONG)
    End With
End Sub

'==============================================================================
' YARDIMCI: Kisi bilgilerini sayfadan oku (ByRef ile dondurur)
'==============================================================================
Private Sub KisiBilgileriniOku(chrome As clsBrowser, _
                                dogumOlumAlsinMi As Boolean, _
                                ByRef outAd As String, _
                                ByRef outSoyad As String, _
                                ByRef outTC As String, _
                                ByRef outAdres As String, _
                                ByRef outDogumTarihi As String, _
                                ByRef outOlumTarihi As String)

    With chrome
        outAd = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")

        ' Bazi kayitlarda ad bos olabiliyor, index kaydiriliyor
        If outAd = "" Then
            outAd = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
            outSoyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[3].innerText")
        Else
            outSoyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        End If

        outTC = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
        outAdres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")

        outDogumTarihi = ""
        outOlumTarihi = ""
        If dogumOlumAlsinMi Then
            outOlumTarihi = .jsEval("document.getElementById('MainContent_lbl_OlumTar').parentElement.parentElement.innerText")
            outDogumTarihi = .jsEval("document.getElementById('MainContent_lbl_DogTar').parentElement.parentElement.innerText")
        End If
    End With
End Sub

'==============================================================================
' YARDIMCI: "Kayit Bulunamadi" kontrolu
'==============================================================================
Private Function KayitBulunamadiMi(chrome As clsBrowser) As String
    Dim jscd As String
    Dim sonuc As String

    With chrome
        On Error Resume Next
        jscd = "document.evaluate(""//span[contains(., 'Kayit Bulunamadi.')]"", document).iterateNext().innerText"
        sonuc = .jsEval(jscd)
        .wait
        On Error GoTo 0
    End With

    If sonuc Like "*Kayit Bulunamadi.*" Then
        KayitBulunamadiMi = sonuc
    Else
        KayitBulunamadiMi = ""
    End If
End Function

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
    Dim hataMesaji As String
    Dim tcNo As String
    Dim kayitHata As String
    Dim ad As String
    Dim soyad As String
    Dim tc As String
    Dim adres As String
    Dim dogTar As String
    Dim olumTar As String

    ' TC No belirle
    tcNo = SorgulanacakTC(UserForm1.TextBox1.Value)
    If tcNo = "" Then Exit Sub

    ' Chrome baslat ve giris yap (basarisizsa tekrar dene)
    If Not ChromeBaslatVeGirisYap(chrome, _
            Trim(UserForm1.txtkullanici_adi.Text), _
            Trim(UserForm1.txtSifre.Text)) Then
        ' Tekrar dene
        If Not ChromeBaslatVeGirisYap(chrome, _
                Trim(UserForm1.txtkullanici_adi.Text), _
                Trim(UserForm1.txtSifre.Text)) Then
            MsgBox "Chrome baslatilamadi!", vbExclamation, "Coder By Ozhan COLAK"
            Exit Sub
        End If
    End If

    ' TC sorgusu yap
    hataMesaji = TCKNSorgula(chrome, tcNo)

    ' Hata varsa (HATA <> "1" demek hata var)
    If hataMesaji <> "1" Then
        chrome.jsEval "document.getElementsByTagName('button')[0].click()"
        chrome.wait

        ' Yabanci uyruklu kontrolu (TC 99 ile basliyorsa)
        If Left(tcNo, 2) = "99" Then
            YabanciKimlikSorgula chrome, tcNo
            KisiBilgileriniOku chrome, False, ad, soyad, tc, adres, dogTar, olumTar
            GoTo FormaYaz
        End If

        ' Normal hata - mesaji goster
        UserForm1.TextBox1.Text = hataMesaji
        MsgBox hataMesaji, vbInformation, "Coder By Ozhan COLAK"
        chrome.maximized
        Exit Sub
    End If

    Application.Wait Now + TimeValue(WAIT_SHORT)
    KisiBilgileriniOku chrome, False, ad, soyad, tc, adres, dogTar, olumTar

FormaYaz:
    ' Sonuclari forma ve sayfaya yaz
    UserForm1.TextBox1.Text = tc
    Sayfa8.Range("L6").Value = tc
    UserForm1.TextBox2.Text = ad & Chr(32) & soyad
    Sayfa8.Range("L5").Value = ad & Chr(32) & soyad
    UserForm1.TextBox3.Text = adres
    Sayfa8.Range("L7").Value = adres

    ' Kayit bulunamadi kontrolu
    kayitHata = KayitBulunamadiMi(chrome)
    If kayitHata <> "" Then
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            ' Yabanci uyruklu icin ek navigasyon
            chrome.jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
        End If
        UserForm1.TextBox3.Text = kayitHata
    End If
End Sub

'==============================================================================
' 2. MERNIS HASTA DOGUM/OLUM TARIHI AL - Chrome
'==============================================================================
Sub MernisHastaDogumTarihiOlumTarihiAl_Chrome()
    Dim chrome As New clsBrowser
    Dim hataMesaji As String
    Dim tcNo As String
    Dim kayitHata As String
    Dim kullaniciAdi As String
    Dim sifre As String
    Dim ad As String
    Dim soyad As String
    Dim tc As String
    Dim adres As String
    Dim dogTar As String
    Dim olumTar As String

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

    ' Chrome baslat ve giris yap
    If Not ChromeBaslatVeGirisYap(chrome, kullaniciAdi, Trim(sifre)) Then
        MsgBox "Chrome baslatilamadi!", vbExclamation, "Coder By Ozhan COLAK"
        Exit Sub
    End If

    ' TC sorgusu yap
    hataMesaji = TCKNSorgula(chrome, tcNo)

    ' Hata varsa
    If hataMesaji <> "1" Then
        chrome.jsEval "document.getElementsByTagName('button')[0].click()"
        chrome.wait

        ' Yabanci uyruklu kontrolu
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            YabanciKimlikSorgula chrome, UserForm1.TextBox1.Value
            KisiBilgileriniOku chrome, True, ad, soyad, tc, adres, dogTar, olumTar
            GoTo FormaYaz
        End If

        ' Normal hata
        UserForm1.TextBox1.Text = hataMesaji
        MsgBox hataMesaji, vbInformation, "Coder By Ozhan COLAK"
        Exit Sub
    End If

    Application.Wait Now + TimeValue(WAIT_SHORT)
    KisiBilgileriniOku chrome, True, ad, soyad, tc, adres, dogTar, olumTar

FormaYaz:
    ' Formu temizle
    iadeFormTemizle

    ' Hasta bilgilerini forma yaz
    UserForm1.LabelHastaBilgi.Caption = "HASTA BILGILERI" & _
        Chr(10) & tc & _
        Chr(10) & ad & Chr(32) & soyad & _
        Chr(10) & dogTar & _
        Chr(10) & olumTar

    ' Sayfa ve hucrelere yaz
    Sayfa8.Range("L4").Value = tc
    Sayfa8.Range("L3").Value = ad & Chr(32) & soyad

    Application.EnableEvents = False
    Sheet1.Range("B22").Value = dogTar
    Sheet1.Range("B23").Value = olumTar
    Application.EnableEvents = True

    ' Kayit bulunamadi kontrolu
    kayitHata = KayitBulunamadiMi(chrome)
    If kayitHata <> "" Then
        If Left(UserForm1.TextBox1.Value, 2) = "99" Then
            chrome.jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"
            chrome.wait
        End If
        UserForm1.LabelHastaBilgi.Caption = UserForm1.LabelHastaBilgi.Caption & Chr(10) & kayitHata
    End If
End Sub
