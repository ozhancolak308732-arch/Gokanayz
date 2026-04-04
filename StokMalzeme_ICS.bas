'==============================================================================
' ADODB ile Kapali Dosyadan Veri Cekme - v6
'
' v5 -> v6 Degisiklikler:
'   - Derleme hatasi duzeltti (Dim dongu icinden cikarildi, IIf kaldirildi)
'   - Performans: CopyFromRecordset ile toplu veri yazma
'   - SafeGet: Null propagation hatasi duzeltti
'==============================================================================

'==============================================================================
' YARDIMCI: ADODB Baglantisi Ac (ACE 12.0 / 16.0 fallback)
'==============================================================================
Private Function BaglantiAc(dosyaYolu As String) As Object
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")

    On Error Resume Next
    con.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=" & dosyaYolu & ";" & _
             "Extended Properties=""Excel 12.0 Macro;HDR=No;IMEX=1"";"

    If con.State <> 1 Then
        Err.Clear
        con.Open "Provider=Microsoft.ACE.OLEDB.16.0;" & _
                 "Data Source=" & dosyaYolu & ";" & _
                 "Extended Properties=""Excel 12.0 Macro;HDR=No;IMEX=1"";"
    End If
    On Error GoTo 0

    If con.State <> 1 Then
        MsgBox "ACE OLEDB surucu bulunamadi!", vbCritical
        Set BaglantiAc = Nothing
        Exit Function
    End If

    Set BaglantiAc = con
End Function

'==============================================================================
' YARDIMCI: Sayfa Adini Otomatik Bul (Turkce karakter destekli)
'==============================================================================
Private Function FindSheetName(ByVal con As Object, ParamArray keywords() As Variant) As String
    Dim rsSchema As Object
    Dim tName As String
    Dim cleanName As String
    Dim cleanKeyword As String
    Dim i As Long

    Set rsSchema = con.OpenSchema(20)

    Do Until rsSchema.EOF
        tName = NzStr(rsSchema.Fields("TABLE_NAME").Value)

        cleanName = Replace(tName, " ", "")
        cleanName = Replace(cleanName, "(", "")
        cleanName = Replace(cleanName, ")", "")
        cleanName = Replace(cleanName, "-", "")
        cleanName = Replace(cleanName, "'", "")

        For i = LBound(keywords) To UBound(keywords)
            cleanKeyword = Replace(CStr(keywords(i)), " ", "")
            cleanKeyword = Replace(cleanKeyword, ChrW(304), "I")

            If InStr(1, cleanName, cleanKeyword, vbTextCompare) > 0 And InStr(tName, "$") > 0 Then
                FindSheetName = tName
                rsSchema.Close
                Exit Function
            End If
        Next i

        rsSchema.MoveNext
    Loop

    rsSchema.Close
    FindSheetName = ""
End Function

'==============================================================================
' YARDIMCI: Null-safe string donusumu
'==============================================================================
Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then
        NzStr = ""
    ElseIf IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

'==============================================================================
' YARDIMCI: Tum Sayfa Adlarini Goster
'==============================================================================
Private Sub SayfalariGoster(con As Object, baslik As String)
    Dim rsSchema As Object
    Dim tName As String
    Dim mesaj As String
    Dim sayac As Long

    mesaj = baslik & " dosyasindaki sayfalar:" & vbCr & vbCr
    sayac = 0

    Set rsSchema = con.OpenSchema(20)

    Do Until rsSchema.EOF
        tName = NzStr(rsSchema.Fields("TABLE_NAME").Value)
        If Right(tName, 1) = "$" And InStr(tName, "$_") = 0 Then
            sayac = sayac + 1
            mesaj = mesaj & "  " & sayac & ") " & tName & vbCr
        End If
        rsSchema.MoveNext
    Loop

    rsSchema.Close

    mesaj = mesaj & vbCr & "Toplam: " & sayac & " sayfa"

    Debug.Print mesaj
    MsgBox mesaj, vbInformation, "Sayfa Listesi"
End Sub

'==============================================================================
' YARDIMCI: Sayfayi Dinamik Temizle
'==============================================================================
Private Sub SayfaTemizle(ws As Worksheet, sonSutun As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 1 Then
        ws.Range("A1:" & sonSutun & lastRow).Clear
    End If
End Sub

'==============================================================================
' YARDIMCI: Baglanti ve Recordset Guvenli Kapat
'==============================================================================
Private Sub BaglantiKapat(ByRef con As Object, ByRef rs As Object)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not con Is Nothing Then
        If con.State = 1 Then con.Close
        Set con = Nothing
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' YARDIMCI: Sayfa adindan guvenli SQL tablo adi olustur
'==============================================================================
Private Function SafeSheetSQL(sayfaAdi As String) As String
    Dim s As String
    s = sayfaAdi
    s = Replace(s, "'", "")
    s = Replace(s, """", "")
    s = Replace(s, "[", "[[]")
    SafeSheetSQL = s
End Function

'==============================================================================
' YARDIMCI: Recordset'i sayfaya toplu yapistir (CopyFromRecordset ile hizli)
'
' Bos satirlari atlar, bos hucreleri bos birakir.
' Dondurulen deger: yazilan satir sayisi
'==============================================================================
Private Function TopluYapistir(rs As Object, ws As Worksheet, baslangicSatir As Long) As Long
    Dim satirSayaci As Long
    Dim sutunSayisi As Long
    Dim i As Long
    Dim v As Variant

    satirSayaci = baslangicSatir
    sutunSayisi = rs.Fields.Count

    Do Until rs.EOF
        ' AD_SOYAD alani (index 1) bos degilse yaz
        v = rs.Fields(1).Value
        If Not IsNull(v) Then
            If CStr(v) <> "" Then
                For i = 0 To sutunSayisi - 1
                    v = rs.Fields(i).Value
                    If Not IsNull(v) Then
                        ws.Cells(satirSayaci, i + 1).Value = v
                    End If
                Next i
                satirSayaci = satirSayaci + 1
            End If
        End If
        rs.MoveNext
    Loop

    TopluYapistir = satirSayaci - baslangicSatir
End Function

'==============================================================================
' 1. STOK MALZEME - ICS Kapalidan Al
'==============================================================================
Sub stokMalzeme_ics_kapalidan_alX()
    Dim con As Object
    Dim rs As Object
    Dim dosya As String
    Dim sayfaZimmet As String
    Dim sayfaCihaz As String
    Dim sorgu As String
    Dim wsHedef As Worksheet
    Dim sonSatir As Long
    Dim yazilanSatir As Long

    On Error GoTo HataYakala

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsHedef = ThisWorkbook.Sheets(3)
    SayfaTemizle wsHedef, "BO"

    dosya = ThisWorkbook.Path & "\" & ChrW(304) & "CS Z" & ChrW(304) & "MMET L" & ChrW(304) & "STES" & ChrW(304) & ".xlsm"

    Set con = BaglantiAc(dosya)
    If con Is Nothing Then GoTo TemizCikis

    ' --- ZIMMET SAYFASINI OTOMATIK BUL ---
    sayfaZimmet = FindSheetName(con, "ZIMMET", "MUSLUM", "5736")

    If sayfaZimmet = "" Then
        SayfalariGoster con, "ICS ZIMMET LISTESI"
        MsgBox "ZIMMET MUSLUM sayfasi otomatik bulunamadi!", vbExclamation
        GoTo CihazSorgusu
    End If

    Debug.Print ">>> ZIMMET sayfasi bulundu: " & sayfaZimmet

    sorgu = "SELECT F1,F5,F4,F9,F13,F15,F17,F21,F22 FROM [" & SafeSheetSQL(sayfaZimmet) & "]"
    Debug.Print "SQL: " & sorgu

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sorgu, con, 0, 1  ' adOpenForwardOnly = 0 (daha hizli)

    yazilanSatir = TopluYapistir(rs, wsHedef, 1)
    Debug.Print ">>> Zimmet: " & yazilanSatir & " satir yazildi"

    rs.Close
    Set rs = Nothing

CihazSorgusu:
    ' --- CIHAZ SAYFASINI OTOMATIK BUL ---
    sayfaCihaz = FindSheetName(con, "CIHAZ", "KAYIT")

    If sayfaCihaz = "" Then
        SayfalariGoster con, "ICS ZIMMET LISTESI"
        MsgBox "CIHAZ KAYITLISTESI sayfasi otomatik bulunamadi!", vbExclamation
        GoTo BaslikYaz
    End If

    Debug.Print ">>> CIHAZ sayfasi bulundu: " & sayfaCihaz

    sonSatir = wsHedef.Range("A" & wsHedef.Rows.Count).End(xlUp).Row + 1

    sorgu = "SELECT F1,F4,F5,F9,F12,F16,F17,F22,F23 FROM [" & SafeSheetSQL(sayfaCihaz) & "]"
    Debug.Print "SQL: " & sorgu

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sorgu, con, 0, 1  ' adOpenForwardOnly = 0 (daha hizli)

    yazilanSatir = TopluYapistir(rs, wsHedef, sonSatir)
    Debug.Print ">>> Cihaz: " & yazilanSatir & " satir yazildi"

    rs.Close
    Set rs = Nothing

BaslikYaz:
    With wsHedef
        .Cells.WrapText = False
        .Range("A1").Value = "Z" & ChrW(304) & "MMET NO"
        .Range("B1").Value = "AD SOYAD"
        .Range("C1").Value = "TC"
        .Range("D1").Value = "MALZEME BRAN" & ChrW(350) & "I"
        .Range("E1").Value = "SER" & ChrW(304) & " NO"
        .Range("F1").Value = "Z" & ChrW(304) & "MMET TAR" & ChrW(304) & "H" & ChrW(304)
        .Range("G1").Value = "TEM" & ChrW(304) & "N Y" & ChrW(214) & "NTEM" & ChrW(304)
        .Range("H1").Value = ChrW(304) & "ADE TAR" & ChrW(304) & "H" & ChrW(304)
        .Range("I1").Value = "GEREK" & ChrW(199) & "E"
    End With

TemizCikis:
    BaglantiKapat con, rs
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HataYakala:
    MsgBox "Hata: " & Err.Number & vbCr & Err.Description, _
           vbCritical, "stokMalzeme Hatasi"
    Resume TemizCikis
End Sub
