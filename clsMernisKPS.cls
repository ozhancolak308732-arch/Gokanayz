VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMernisKPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsMernisKPS - Mernis KPS Chrome Islemleri Class Module
'
' Kullanim:
'   Dim chrome As New clsBrowser
'   Dim kps As New clsMernisKPS
'   Set kps.Browser = chrome
'   kps.BaslatVeGirisYap "kullanici", "sifre"
'   kps.TCKNSorgula "12345678901"
'==============================================================================

Option Explicit

Private Const KPS_LOGIN_URL As String = "https://kps.sgk.intra/KPS/Login.aspx"
Private Const KPS_SORGU_URL As String = "https://kps.sgk.intra/KPS/TCKNodanSorgula.aspx"
Private Const WAIT_SHORT As String = "0:00:01"
Private Const WAIT_LONG As String = "0:00:03"

' Chrome browser nesnesi
Private m_chrome As clsBrowser

' Okunan kisi bilgileri
Public Ad As String
Public Soyad As String
Public TC As String
Public Adres As String
Public DogumTarihi As String
Public OlumTarihi As String

'==============================================================================
' Browser nesnesini ata
'==============================================================================
Public Property Set Browser(ByRef br As clsBrowser)
    Set m_chrome = br
End Property

Public Property Get Browser() As clsBrowser
    Set Browser = m_chrome
End Property

'==============================================================================
' Chrome baslatip KPS'ye giris yap
' Basarili ise True dondurur
'==============================================================================
Public Function BaslatVeGirisYap(kullaniciAdi As String, _
                                  sifre As String) As Boolean
    Dim jscd As String
    Dim basla As Single

    BaslatVeGirisYap = False

    ' Onceki Chrome'u kapat
    CreateObject("wscript.shell").Run "cmd /c """ & "taskkill /IM chrome.exe >nul""""", 0, True

    With m_chrome
        .minimized
        .start cleanActiveSession:=True, userProfile:="User G"

        ' Session kontrolu
        If .SessionID = "" Then Exit Function

        ' KPS Login sayfasina git
        .navigate KPS_LOGIN_URL
        .wait

        ' SSL sertifika uyarisini gec
        basla = Timer: Do While (Timer - basla) < 1: Loop
        On Error Resume Next
        .jsEval "document.querySelector('#details-button').click()"
        .wait
        .jsEval "document.querySelector('#proceed-link').click()"
        .wait
        On Error GoTo 0

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

    BaslatVeGirisYap = True
End Function

'==============================================================================
' TC Kimlik No ile sorgu yap
' Donen deger: hata mesaji ("1" ise basarili)
'==============================================================================
Public Function TCKNSorgula(tcKimlikNo As String) As String
    Dim jscd As String

    With m_chrome
        jscd = "document.getElementsByName('ctl00$MainContent$txtbox_TCKNo')[0].value = '" & tcKimlikNo & "';" & _
               "document.getElementById('MainContent_ddlist_NkoTip')[0].value = '2';" & _
               "document.getElementsByName('ctl00$MainContent$btn_Sorgula')[0].click();"
        .jsEval jscd

        Application.Wait Now + TimeValue(WAIT_LONG)

        TCKNSorgula = .jsEval("document.getElementById('uppnl_HataMesaj').innerText.trim()")
    End With
End Function

'==============================================================================
' Yabanci kimlik ile sorgu yap (TC 99 ile basliyorsa)
'==============================================================================
Public Sub YabanciKimlikSorgula(yabanciNo As String)
    Dim jscd As String

    With m_chrome
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
' Kisi bilgilerini sayfadan oku
' dogumOlumAl = True ise dogum/olum tarihlerini de okur
'==============================================================================
Public Sub KisiBilgileriniOku(dogumOlumAl As Boolean)
    With m_chrome
        Ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")

        ' Bazi kayitlarda ad bos olabiliyor, index kaydiriliyor
        If Ad = "" Then
            Ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
            Soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[3].innerText")
        Else
            Soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        End If

        TC = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
        Adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")

        DogumTarihi = ""
        OlumTarihi = ""
        If dogumOlumAl Then
            OlumTarihi = .jsEval("document.getElementById('MainContent_lbl_OlumTar').parentElement.parentElement.innerText")
            DogumTarihi = .jsEval("document.getElementById('MainContent_lbl_DogTar').parentElement.parentElement.innerText")
        End If
    End With
End Sub

'==============================================================================
' Yabanci uyruklu icin kisi bilgilerini oku (farkli element ID'leri)
'==============================================================================
Public Sub YabanciKisiBilgileriniOku(dogumOlumAl As Boolean)
    With m_chrome
        Ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[1].innerText")
        If Ad = "" Then
            Ad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        Else
            Soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[2].innerText")
        End If
        Soyad = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[3].innerText")
        TC = .jsEval("document.getElementById('kisibilgileriSol').getElementsByTagName('dd')[0].innerText")
        Adres = .jsEval("document.getElementById('MainContent_lbl_AdresBilgi').getElementsByTagName('dl')[4].getElementsByTagName('dd')[0].innerText")

        DogumTarihi = ""
        OlumTarihi = ""
        If dogumOlumAl Then
            OlumTarihi = .jsEval("document.getElementById('MainContent_lbl_EGMId').parentElement.parentElement.innerText")
            DogumTarihi = .jsEval("document.getElementById('MainContent_lbl_DogTar').parentElement.parentElement.innerText")
        End If
    End With
End Sub

'==============================================================================
' "Kayit Bulunamadi" kontrolu
' Bos string donerse kayit bulunmus demektir
'==============================================================================
Public Function KayitBulunamadiMi() As String
    Dim jscd As String
    Dim sonuc As String

    On Error Resume Next
    jscd = "document.evaluate(""//span[contains(., 'Kayit Bulunamadi.')]"", document).iterateNext().innerText"
    sonuc = m_chrome.jsEval(jscd)
    m_chrome.wait
    On Error GoTo 0

    If sonuc Like "*Kayit Bulunamadi.*" Then
        KayitBulunamadiMi = sonuc
    Else
        KayitBulunamadiMi = ""
    End If
End Function

'==============================================================================
' Hata butonuna tikla (dialog kapat)
'==============================================================================
Public Sub HataDialogKapat()
    With m_chrome
        .jsEval "document.getElementsByTagName('button')[0].click()"
        .wait
    End With
End Sub

'==============================================================================
' AdresCheckbox tikla
'==============================================================================
Public Sub AdresCheckboxTikla()
    m_chrome.jsEval "document.querySelector('#MainContent_chkbox_AdresBilgisi').click()"
    m_chrome.wait
End Sub

'==============================================================================
' Yabanci navigasyon tikla
'==============================================================================
Public Sub YabanciNavigasyonTikla()
    m_chrome.jsEval "document.querySelector('#navigationSol > li:nth-child(4) > a').click()"
End Sub
