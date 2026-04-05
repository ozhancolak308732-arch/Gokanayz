'==============================================================================
' Mernis Chrome Otomasyon Modulu - Optimize Edilmis Versiyon
'==============================================================================

Option Explicit

Private Const KPS_LOGIN_URL As String = "https://kps.sgk.intra/KPS/Login.aspx"
Private Const KPS_SORGU_URL As String = "https://kps.sgk.intra/KPS/TCKNodanSorgula.aspx"
Private Const WAIT_SHORT As String = "0:00:01"
Private Const WAIT_LONG As String = "0:00:03"

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
    Dim jscd As String
    Dim hataMesaji As String
    Dim tcNo As String
    Dim kayitHata As String
    Dim ad As String
    Dim soyad As String
    Dim tc As String
    Dim adres As String
    Dim basla As Single

    ' TC No belirle
    tcNo = SorgulanacakTC(UserForm1.TextBox1.Value)
    If tcNo = "" Then Exit Sub

On Error Resume Next
1:
    CreateObject("wscript.shell").Run "cmd /c """ & "taskkill /IM chrome.exe >nul""""", 0, True

    With chrome
        .minimized
        .start cleanActiveSession:=True, userProfile:="User G"

        If .SessionID = "" Then GoTo 1

        ' KPS Login sayfasina git
        .navigate KPS_LOGIN_URL
        .wait

        ' SSL sertifika uyarisini gec
        basla = Timer: Do While (Timer - basla) < 1: Loop
        .jsEval "document.querySelector('#details-button').click()"
        .wait
        .jsEval "document.querySelector('#proceed-link').click()"
        .wait

        ' Giris yap
        jscd = "document.getElementsByName('txtUser')[0].value = '" & UserForm1.txtkullanici_adi.Text & "';" & _
               "document.getElementsByName('txtPass')[0].value = '" & Trim(UserForm1.txtSifre.Text) & "';" & _
               "document.getElementsByName('btnLogin')[0].click();"
        .jsEval jscd
        .wait

        ' Sorgu sayfasina git
        .navigate KPS_SORGU_URL
        .wait

        ' Adres bilgisi checkbox'ini tikla
        .jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"

        ' TC sorgusu yap
        jscd = "document.getElementsByName('ctl00$MainContent$txtbox_TCKNo')[0].value = '" & tcNo & "';" & _
               "document.getElementById('MainContent_ddlist_NkoTip')[0].value = '2';" & _
               "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
        .jsEval jscd

        Application.Wait Now + TimeValue(WAIT_LONG)

        ' Hata kontrolu
        hataMesaji = .jsEval("document.getElementById('uppnl_HataMesaj').innerText.trim()")

        If Not hataMesaji = "1" Then
            .jsEval "document.getElementsByTagName('button')[0].click()"
            .wait

            ' Yabanci uyruklu kontrolu (TC 99 ile basliyorsa)
            If Left(tcNo, 2) = "99" Then
                .jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
                .wait
                jscd = "document.getElementsByName('ctl00$MainContent$txtbox_YabanciKNo')[0].value = '" & tcNo & "';" & _
                       "document.getElementById('MainContent_chkbox_AdresBilgisi').click();" & _
                       "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
                .jsEval jscd
                Application.Wait Now + TimeValue(WAIT_LONG)

                ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")
                If ad = "" Then
                    ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
                Else
                    soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
                End If
                soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[3].innerText")
                tc = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
                adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")
                GoTo FormaYaz
            End If

            ' Normal hata - mesaji goster
            UserForm1.TextBox1.Text = hataMesaji
            GoTo HataGoster
        End If

        Application.Wait Now + TimeValue(WAIT_SHORT)

        ' Kisi bilgilerini oku
        ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")
        soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        tc = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
        adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")

FormaYaz:
        ' Sonuclari forma ve sayfaya yaz
        UserForm1.TextBox1.Text = tc
        Sayfa8.Range("L6").Value = tc
        UserForm1.TextBox2.Text = ad & Chr(32) & soyad
        Sayfa8.Range("L5").Value = ad & Chr(32) & soyad
        UserForm1.TextBox3.Text = adres
        Sayfa8.Range("L7").Value = adres

        ' Kayit bulunamadi kontrolu
        jscd = "document.evaluate(""//span[contains(., 'Kayit Bulunamadi.')]"", document).iterateNext().innerText"
        kayitHata = .jsEval(jscd)
        .wait
        If kayitHata Like "*Kayit Bulunamadi.*" Then
            If Left(UserForm1.TextBox1.Value, 2) = "99" Then
                .jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
            End If
            UserForm1.TextBox3.Text = kayitHata
        End If

    End With
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
    Dim jscd As String
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
    Dim basla As Single

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
    CreateObject("wscript.shell").Run "cmd /c """ & "taskkill /IM chrome.exe >nul""""", 0, True

    With chrome
        .minimized
        .start cleanActiveSession:=True, userProfile:="User G"

        ' KPS Login sayfasina git
        .navigate KPS_LOGIN_URL
        .wait

        ' SSL sertifika uyarisini gec
        basla = Timer: Do While (Timer - basla) < 1: Loop
        .jsEval "document.querySelector('#details-button').click()"
        .wait
        .jsEval "document.querySelector('#proceed-link').click()"
        .wait

        ' Giris yap
        jscd = "document.getElementsByName('txtUser')[0].value = '" & kullaniciAdi & "';" & _
               "document.getElementsByName('txtPass')[0].value = '" & Trim(sifre) & "';" & _
               "document.getElementsByName('btnLogin')[0].click();"
        .jsEval jscd
        .wait

        ' Sorgu sayfasina git
        .navigate KPS_SORGU_URL
        .wait

        ' Adres bilgisi checkbox'ini tikla
        .jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"

        ' TC sorgusu yap
        jscd = "document.getElementsByName('ctl00$MainContent$txtbox_TCKNo')[0].value = '" & tcNo & "';" & _
               "document.getElementById('MainContent_ddlist_NkoTip')[0].value = '2';" & _
               "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
        .jsEval jscd

        Application.Wait Now + TimeValue(WAIT_LONG)

        ' Hata kontrolu
        hataMesaji = .jsEval("document.getElementById('uppnl_HataMesaj').innerText.trim()")

        If Not hataMesaji = "1" Then
            .jsEval "document.getElementsByTagName('button')[0].click()"
            .wait

            ' Yabanci uyruklu kontrolu (TC 99 ile basliyorsa)
            If Left(UserForm1.TextBox1.Value, 2) = "99" Then
                .jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
                .wait
                jscd = "document.getElementsByName('ctl00$MainContent$txtbox_YabanciKNo')[0].value = '" & UserForm1.TextBox1.Value & "';" & _
                       "document.getElementById('MainContent_chkbox_AdresBilgisi').click();" & _
                       "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
                .jsEval jscd
                Application.Wait Now + TimeValue(WAIT_LONG)

                ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")
                If ad = "" Then
                    ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
                Else
                    soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
                End If
                soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[3].innerText")
                tc = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
                adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")
                olumTar = .jsEval("document.getElementById('MainContent_lbl_EGMId').parentElement.parentElement.innerText")
                dogTar = .jsEval("document.getElementById('MainContent_lbl_DogTar').parentElement.parentElement.innerText")
                GoTo FormaYaz
            End If

            ' Normal hata
            UserForm1.TextBox1.Text = hataMesaji
            GoTo HataGoster
        End If

        Application.Wait Now + TimeValue(WAIT_SHORT)

        ' Kisi bilgilerini oku
        ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")
        soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        tc = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
        adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")
        olumTar = .jsEval("document.getElementById('MainContent_lbl_OlumTar').parentElement.parentElement.innerText")
        dogTar = .jsEval("document.getElementById('MainContent_lbl_DogTar').parentElement.parentElement.innerText")

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
        jscd = "document.evaluate(""//span[contains(., 'Kayit Bulunamadi.')]"", document).iterateNext().innerText"
        kayitHata = .jsEval(jscd)
        .wait
        If kayitHata Like "*Kayit Bulunamadi.*" Then
            If Left(UserForm1.TextBox1.Value, 2) = "99" Then
                .jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"
                .wait
            End If
            UserForm1.LabelHastaBilgi.Caption = UserForm1.LabelHastaBilgi.Caption & Chr(10) & kayitHata
        End If

    End With
    Exit Sub

HataGoster:
    MsgBox hataMesaji, vbInformation, "Coder By Ozhan COLAK"
End Sub
