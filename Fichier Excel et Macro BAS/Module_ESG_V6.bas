Attribute VB_Name = "M_ESG_DURABILITE_V5"
Option Explicit
'==================================================================================
'  OUTIL DE DURABILITE  -  FEC + DSN  ->  TABLEAU DE BORD ESG   (moteur V5)
'  - Import FEC et DSN (logique eprouvee, conservee)
'  - Mapping par NUMERO DE COMPTE pilote par la feuille 10 (prefixe le plus long)
'  - Les colonnes de calcul de la feuille 04 sont des FORMULES vivantes
'  - Boutons dessines par Installer_Outil (aucune manip XML)
'==================================================================================

Private Const SH_CMD   As String = "01_ACCUEIL"
Private Const SH_FEC   As String = "04_FEC_BRUT"
Private Const SH_DSN   As String = "05_DSN_BRUT"
Private Const SH_BG    As String = "06_BALANCE_ESG"
Private Const SH_DSN_L As String = "07_SOCIAL_DSN"
Private Const SH_CTRL  As String = "08_CONTROLES"
Private Const SH_DASH  As String = "03_TABLEAU_DE_BORD"
Private Const SH_MAN   As String = "02_DONNEES"
Private Const SH_PARAM As String = "09_PARAMETRES_SOURCES"

Private Const FEC_ROW0 As Long = 5   ' premiere ligne de donnees FEC
Private Const BG_ROW0  As Long = 5
Private Const DSN_ROW0 As Long = 5
Private Const DL_ROW0  As Long = 5

'================================ ETAPES (boutons) ================================

Public Sub Importer_FEC()
    On Error GoTo EH
    Dim p As String: p = ResolveFecPath()
    If Len(p) = 0 Then Exit Sub
    Application.ScreenUpdating = False
    ImportFecFile p
    Application.ScreenUpdating = True
    Dim n As Long: n = Worksheets(SH_FEC).Cells(Worksheets(SH_FEC).Rows.Count, 1).End(xlUp).Row - (FEC_ROW0 - 1)
    If n < 0 Then n = 0
    SetStatus "FEC importe : " & n & " lignes. Vous pouvez importer les DSN."
    MsgBox "Import FEC termine : " & n & " lignes.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Erreur import FEC : " & Err.Description, vbCritical
End Sub

Public Sub Importer_DSN()
    On Error GoTo EH
    Application.ScreenUpdating = False
    ImportDsnFolder
    Application.ScreenUpdating = True
    Dim n As Long: n = Worksheets(SH_DSN).Cells(Worksheets(SH_DSN).Rows.Count, 1).End(xlUp).Row - (DSN_ROW0 - 1)
    If n < 0 Then n = 0
    SetStatus "DSN importees : " & n & " lignes brutes. Lancez la generation de l'analyse."
    MsgBox "Import DSN termine : " & n & " lignes brutes.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Erreur import DSN : " & Err.Description, vbCritical
End Sub

Public Sub Generer_Analyse()
    On Error GoTo EH
    Application.ScreenUpdating = False
    GenererBG
    ConstruireDsnLisible
    Application.CalculateFull
    Application.ScreenUpdating = True
    Dim n As Long: n = Worksheets(SH_BG).Cells(Worksheets(SH_BG).Rows.Count, 1).End(xlUp).Row - (BG_ROW0 - 1)
    If n < 0 Then n = 0
    SetStatus "Analyse generee : " & n & " comptes traites. Completez la feuille 09 puis consultez le tableau de bord."
    MsgBox "Analyse ESG generee : " & n & " comptes." & vbCrLf & _
           "Pensez a completer la feuille 09 (donnees manuelles).", vbInformation
    Worksheets(SH_DASH).Activate
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Erreur generation analyse : " & Err.Description, vbCritical
End Sub

Public Sub Actualiser_Dashboard()
    On Error GoTo EH
    Application.ScreenUpdating = False
    If Worksheets(SH_FEC).Cells(Worksheets(SH_FEC).Rows.Count, 1).End(xlUp).Row >= FEC_ROW0 Then GenererBG
    If Worksheets(SH_DSN).Cells(Worksheets(SH_DSN).Rows.Count, 1).End(xlUp).Row >= DSN_ROW0 Then ConstruireDsnLisible
    Application.CalculateFull
    Application.ScreenUpdating = True
    SetStatus "Tableau de bord actualise le " & Format(Now, "dd/mm/yyyy hh:nn") & "."
    Worksheets(SH_DASH).Activate
    MsgBox "Tableau de bord actualise.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Erreur actualisation : " & Err.Description, vbCritical
End Sub

Public Sub Reinitialiser_Outil()
    On Error GoTo EH
    If MsgBox("Vider les donnees importees (FEC, DSN, analyse) et les reponses manuelles ?" & vbCrLf & _
              "Les parametres et facteurs de la feuille 10 sont conserves.", _
              vbYesNo + vbQuestion, "Reinitialiser") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    ClearFrom SH_FEC, FEC_ROW0
    ClearFrom SH_DSN, DSN_ROW0
    Worksheets(SH_BG).Range("A" & BG_ROW0 & ":W" & Worksheets(SH_BG).Rows.Count).ClearContents
    ClearFrom SH_DSN_L, DL_ROW0
    Worksheets(SH_MAN).Range("D5:D100").ClearContents
    Application.CalculateFull
    Application.ScreenUpdating = True
    SetStatus "Outil reinitialise. Renseignez les parametres puis relancez l'import."
    MsgBox "Outil remis a blanc.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Erreur reinitialisation : " & Err.Description, vbCritical
End Sub

'================================ IMPORT FEC ======================================

Private Sub ImportFecFile(ByVal path As String)
    Dim ws As Worksheet: Set ws = Worksheets(SH_FEC)
    ws.Range("A" & FEC_ROW0 & ":R" & ws.Rows.Count).ClearContents
    Dim txt As String: txt = ReadTextSmart(path)
    txt = Replace(txt, vbCrLf, vbLf): txt = Replace(txt, vbCr, vbLf)
    Dim lines() As String: lines = Split(txt, vbLf)
    Dim maxRows As Long: maxRows = UBound(lines) + 1
    Dim arr() As Variant: ReDim arr(1 To maxRows, 1 To 18)
    Dim r As Long, i As Long, f() As String, line As String, idx As Long
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        If Len(Trim(line)) > 0 Then
            f = Split(line, vbTab)
            If UBound(f) >= 17 Then
                If UCase(Replace(CleanField(f(0)), Chr(34), "")) <> "JOURNALCODE" Then
                    r = r + 1
                    For idx = 0 To 17
                        If idx = 11 Or idx = 12 Then
                            arr(r, idx + 1) = ToDoubleFR(f(idx))   ' Debit / Credit numeriques
                        Else
                            arr(r, idx + 1) = CleanField(f(idx))
                        End If
                    Next idx
                End If
            End If
        End If
    Next i
    If r > 0 Then ws.Range("A" & FEC_ROW0).Resize(r, 18).Value = arr
    ws.Columns("L:M").NumberFormat = "#,##0.00"
End Sub

'================================ IMPORT DSN ======================================

Private Sub ImportDsnFolder()
    Dim folder As String: folder = GetFolder()
    If Len(folder) = 0 Then Exit Sub
    Dim ext As String: ext = Trim(CStr(Worksheets(SH_CMD).Range("C8").Value))
    If Len(ext) = 0 Then ext = ".dsn"
    If Left(ext, 1) <> "." Then ext = "." & ext
    Dim ws As Worksheet: Set ws = Worksheets(SH_DSN)
    ws.Range("A" & DSN_ROW0 & ":E" & ws.Rows.Count).ClearContents
    Dim fn As String: fn = Dir(folder & "*" & ext)
    Dim r As Long: r = DSN_ROW0 - 1
    Dim txt As String, lines() As String, i As Long, rub As String, valeur As String, line As String, mois As Long
    Do While Len(fn) > 0
        txt = ReadTextSmart(folder & fn)
        txt = Replace(txt, vbCrLf, vbLf): txt = Replace(txt, vbCr, vbLf)
        lines = Split(txt, vbLf)
        mois = MonthFromFilename(fn)
        For i = LBound(lines) To UBound(lines)
            line = Trim(lines(i))
            If Len(line) > 0 And InStr(1, line, ",") > 0 Then
                rub = Split(line, ",")(0)
                valeur = CleanField(Mid(line, Len(rub) + 2))
                r = r + 1
                ws.Cells(r, 1).Value = mois
                ws.Cells(r, 2).Value = fn
                ws.Cells(r, 3).Value = rub
                ws.Cells(r, 4).Value = valeur
            End If
        Next i
        fn = Dir()
    Loop
End Sub

'================================ GENERER LA BG ESG ===============================

Private Sub GenererBG()
    Dim fec As Worksheet: Set fec = Worksheets(SH_FEC)
    Dim bg As Worksheet: Set bg = Worksheets(SH_BG)
    bg.Range("A" & BG_ROW0 & ":W" & bg.Rows.Count).ClearContents
    Dim last As Long: last = fec.Cells(fec.Rows.Count, 5).End(xlUp).Row
    If last < FEC_ROW0 Then Exit Sub

    ' --- agreger le FEC par numero de compte ---
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, cpt As String, lib As String, deb As Double, cre As Double, a As Variant
    For i = FEC_ROW0 To last
        cpt = Trim(CStr(fec.Cells(i, 5).Value))
        If Len(cpt) > 0 Then
            lib = CStr(fec.Cells(i, 6).Value)
            deb = ToDoubleFR(fec.Cells(i, 12).Value)
            cre = ToDoubleFR(fec.Cells(i, 13).Value)
            If Not d.Exists(cpt) Then d.Add cpt, Array(cpt, lib, 0#, 0#)
            a = d(cpt)
            a(2) = a(2) + deb
            a(3) = a(3) + cre
            If Len(lib) > 0 Then a(1) = lib
            d(cpt) = a
        End If
    Next i
    If d.Count = 0 Then Exit Sub

    ' --- charger la table de correspondance (feuille 10) ---
    Dim prefixes() As String, np As Long
    LoadPrefixes prefixes, np
    If np = 0 Then
        MsgBox "Aucune correspondance trouvee dans la feuille 10.", vbExclamation
        Exit Sub
    End If

    ' --- ecrire compte / libelle / debit / credit / prefixe mappe ---
    Dim keys As Variant: keys = d.Keys
    QuickSortKeys keys, LBound(keys), UBound(keys)
    Dim n As Long: n = d.Count
    Dim outV() As Variant: ReDim outV(1 To n, 1 To 4)   ' A,B,E,F
    Dim outI() As Variant: ReDim outI(1 To n, 1 To 1)   ' I
    Dim j As Long
    For j = LBound(keys) To UBound(keys)
        a = d(keys(j))
        outV(j + 1, 1) = CStr(a(0))
        outV(j + 1, 2) = a(1)
        outV(j + 1, 3) = a(2)
        outV(j + 1, 4) = a(3)
        outI(j + 1, 1) = MatchPrefix(CStr(a(0)), prefixes, np)
    Next j

    Dim lastBG As Long: lastBG = BG_ROW0 + n - 1
    bg.Range("A" & BG_ROW0 & ":A" & lastBG).NumberFormat = "@"
    bg.Range("I" & BG_ROW0 & ":I" & lastBG).NumberFormat = "@"
    bg.Range("A" & BG_ROW0).Resize(n, 1).Value = Application.Index(outV, 0, 1)
    bg.Range("B" & BG_ROW0).Resize(n, 1).Value = Application.Index(outV, 0, 2)
    bg.Range("E" & BG_ROW0).Resize(n, 1).Value = Application.Index(outV, 0, 3)
    bg.Range("F" & BG_ROW0).Resize(n, 1).Value = Application.Index(outV, 0, 4)
    bg.Range("I" & BG_ROW0).Resize(n, 1).Value = outI
    bg.Range("E" & BG_ROW0 & ":H" & lastBG).NumberFormat = "#,##0.00"

    PoseFormulesBG lastBG
End Sub

' Pose les formules vivantes sur la plage de donnees de la BG
Private Sub PoseFormulesBG(ByVal lastBG As Long)
    Dim bg As Worksheet: Set bg = Worksheets(SH_BG)
    Dim a As Long: a = BG_ROW0
    bg.Range("C" & a & ":C" & lastBG).Formula = "=IF($A" & a & "="""","""",LEFT($A" & a & ",1))"
    bg.Range("D" & a & ":D" & lastBG).Formula = "=IF($A" & a & "="""","""",LEFT($A" & a & ",2))"
    bg.Range("G" & a & ":G" & lastBG).Formula = "=IF($A" & a & "="""","""",$E" & a & "-$F" & a & ")"
    bg.Range("H" & a & ":H" & lastBG).Formula = "=IF($A" & a & "="""","""",ABS($G" & a & "))"
    bg.Range("J" & a & ":J" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_IND,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("K" & a & ":K" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_PIL,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("L" & a & ":L" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_THE,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("M" & a & ":M" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_ESRS,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("N" & a & ":N" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_SCOPE,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("O" & a & ":O" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_UNIT,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("P" & a & ":P" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_METH,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("Q" & a & ":Q" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_TYPE,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("R" & a & ":R" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_FEUR,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("S" & a & ":S" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_FEPH,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("T" & a & ":T" & lastBG).Formula = "=IF($I" & a & "="""","""",IFERROR(INDEX(P_FEMO,MATCH($I" & a & ",P_PREF,0)),""""))"
    bg.Range("U" & a & ":U" & lastBG).Formula = "=IF($P" & a & "=""PHYSIQUE"",IFERROR($H" & a & "*$R" & a & ",""""),"""")"
    bg.Range("V" & a & ":V" & lastBG).Formula = "=IF($P" & a & "=""PHYSIQUE"",IFERROR($U" & a & "*$S" & a & ",0),IF($P" & a & "=""MONETAIRE"",IFERROR($H" & a & "*$T" & a & ",0),0))"
    bg.Range("W" & a & ":W" & lastBG).Formula = "=IF($V" & a & "="""","""",IFERROR($V" & a & "/1000,0))"
End Sub

'================================ SYNTHESE SOCIALE (DSN) ==========================

Private Sub ConstruireDsnLisible()
    Dim raw As Worksheet: Set raw = Worksheets(SH_DSN)
    Dim out As Worksheet: Set out = Worksheets(SH_DSN_L)
    out.Range("A" & DL_ROW0 & ":N" & out.Rows.Count).ClearContents
    Dim last As Long: last = raw.Cells(raw.Rows.Count, 1).End(xlUp).Row
    If last < DSN_ROW0 Then Exit Sub

    ' SMIC mensuel de reference (feuille 02), proratise ensuite par la quotite
    Dim smic As Double
    On Error Resume Next
    smic = Worksheets(SH_MAN).Range("M_SMIC_MENSUEL").Value
    On Error GoTo 0
    If smic <= 0 Then smic = 1801.8

    ' emp : mois|nir -> Array(sexe, anNaiss, contrat, quotite, idcc, remu)
    Dim emp As Object: Set emp = CreateObject("Scripting.Dictionary")
    Dim cntAll As Object: Set cntAll = CreateObject("Scripting.Dictionary")
    Dim i As Long, m As Long, rub As String, val As String, curKey As String, a As Variant
    curKey = ""
    For i = DSN_ROW0 To last
        m = CLng(VBA.Val(raw.Cells(i, 1).Value))
        If m < 1 Or m > 12 Then m = 0
        rub = Trim(CStr(raw.Cells(i, 3).Value))
        val = CleanField(CStr(raw.Cells(i, 4).Value))
        If m > 0 Then cntAll(m) = NzL(cntAll, m) + 1
        Select Case rub
            Case "S21.G00.30.001"
                curKey = m & "|" & val
                If Not emp.Exists(curKey) Then emp.Add curKey, Array(0, 0, 0, 0#, "", 0#)
            Case "S21.G00.30.005"
                If Len(curKey) > 0 Then a = emp(curKey): a(0) = CLng(VBA.Val(val)): emp(curKey) = a
            Case "S21.G00.30.006"
                If Len(curKey) > 0 Then a = emp(curKey): a(1) = CLng(VBA.Val(Right(val, 4))): emp(curKey) = a
            Case "S21.G00.40.007"
                If Len(curKey) > 0 Then a = emp(curKey): a(2) = CLng(VBA.Val(val)): emp(curKey) = a
            Case "S21.G00.40.013"
                If Len(curKey) > 0 Then a = emp(curKey): a(3) = ToDoubleFR(val): emp(curKey) = a
            Case "S21.G00.40.017"
                If Len(curKey) > 0 Then a = emp(curKey): a(4) = val: emp(curKey) = a
            Case "S21.G00.51.013"
                If Len(curKey) > 0 Then a = emp(curKey): a(5) = a(5) + ToDoubleFR(val): emp(curKey) = a
        End Select
    Next i

    ' agreger par mois : Array(eff,h,f,etp,cdi,cdd,ageSum,ageN,couverts,remHsum,remHn,remFsum,remFn,nbSMIC)
    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim k As Variant, parts() As String, mm As Long
    For Each k In emp.Keys
        parts = Split(CStr(k), "|")
        mm = CLng(VBA.Val(parts(0)))
        If Not agg.Exists(mm) Then agg.Add mm, Array(0, 0, 0, 0#, 0, 0, 0#, 0, 0, 0#, 0, 0#, 0, 0)
        a = agg(mm)
        Dim e As Variant: e = emp(k)
        a(0) = a(0) + 1
        If e(0) = 1 Then
            a(1) = a(1) + 1
        ElseIf e(0) = 2 Then
            a(2) = a(2) + 1
        End If
        If e(3) > 0 Then a(3) = a(3) + e(3) / 151.67 Else a(3) = a(3) + 1
        If e(2) = 1 Then
            a(4) = a(4) + 1
        ElseIf e(2) = 2 Then
            a(5) = a(5) + 1
        End If
        If e(1) > 1900 Then a(6) = a(6) + (Year(Date) - e(1)): a(7) = a(7) + 1
        ' convention collective
        If Len(CStr(e(4))) > 0 And CStr(e(4)) <> "0000" And CStr(e(4)) <> "9999" Then a(8) = a(8) + 1
        ' remuneration par sexe
        If e(0) = 1 Then
            a(9) = a(9) + e(5): a(10) = a(10) + 1
        ElseIf e(0) = 2 Then
            a(11) = a(11) + e(5): a(12) = a(12) + 1
        End If
        ' au-dessus du SMIC (seuil proratise par la quotite)
        Dim seuil As Double
        If e(3) > 0 Then seuil = smic * e(3) / 151.67 Else seuil = smic
        If e(5) > 0 And e(5) >= seuil Then a(13) = a(13) + 1
        agg(mm) = a
    Next k

    ' trier les mois
    Dim mk() As Long, c As Long: c = 0
    ReDim mk(1 To agg.Count)
    For Each k In agg.Keys: c = c + 1: mk(c) = CLng(k): Next k
    Dim x As Long, y As Long, tmp As Long
    For x = 1 To c - 1
        For y = x + 1 To c
            If mk(y) < mk(x) Then tmp = mk(x): mk(x) = mk(y): mk(y) = tmp
        Next y
    Next x

    Dim rows() As Variant: ReDim rows(1 To c, 1 To 14)
    For x = 1 To c
        a = agg(mk(x))
        rows(x, 1) = MoisNom(mk(x))
        rows(x, 2) = a(0)
        rows(x, 3) = a(1)
        rows(x, 4) = a(2)
        rows(x, 5) = a(3)
        rows(x, 6) = a(4)
        rows(x, 7) = a(5)
        rows(x, 8) = IIf(a(7) > 0, a(6) / a(7), 0)
        rows(x, 9) = NzL(cntAll, mk(x))
        rows(x, 10) = "DSN"
        rows(x, 11) = IIf(a(0) > 0, a(8) / a(0), 0)
        rows(x, 12) = IIf(a(10) > 0, a(9) / a(10), 0)
        rows(x, 13) = IIf(a(12) > 0, a(11) / a(12), 0)
        rows(x, 14) = a(13)
    Next x
    out.Range("A" & DL_ROW0).Resize(c, 14).Value = rows
    out.Range("E" & DL_ROW0 & ":E" & DL_ROW0 + c - 1).NumberFormat = "#,##0.00"
    out.Range("H" & DL_ROW0 & ":H" & DL_ROW0 + c - 1).NumberFormat = "#,##0.0"
    out.Range("K" & DL_ROW0 & ":K" & DL_ROW0 + c - 1).NumberFormat = "0.0%"
    out.Range("L" & DL_ROW0 & ":M" & DL_ROW0 + c - 1).NumberFormat = "#,##0"
End Sub

'================================ TABLE DE CORRESPONDANCE =========================

' Charge la colonne A du bloc de mapping de la feuille 10 dans un tableau
Private Sub LoadPrefixes(ByRef prefixes() As String, ByRef np As Long)
    Dim ws As Worksheet: Set ws = Worksheets(SH_PARAM)
    Dim hdr As Long: hdr = FindRowText(ws, 1, "Prefixe")
    If hdr = 0 Then hdr = FindRowText(ws, 1, "Préfixe")
    np = 0
    If hdr = 0 Then Exit Sub
    Dim r As Long: r = hdr + 1
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim prefixes(1 To (lastRow - hdr + 1))
    Do While r <= lastRow
        Dim v As String: v = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(v) = 0 Then Exit Do
        np = np + 1
        prefixes(np) = v
        r = r + 1
    Loop
End Sub

' Renvoie le prefixe le plus long de la table qui prefixe le compte ; "" si aucun
Private Function MatchPrefix(ByVal cpt As String, ByRef prefixes() As String, ByVal np As Long) As String
    Dim i As Long, best As String, bestLen As Long
    best = "": bestLen = 0
    For i = 1 To np
        Dim p As String: p = prefixes(i)
        If Len(p) > 0 And Len(p) <= Len(cpt) Then
            If Left(cpt, Len(p)) = p And IsAllDigits(p) Then
                If Len(p) > bestLen Then best = p: bestLen = Len(p)
            End If
        End If
    Next i
    MatchPrefix = best
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then IsAllDigits = False: Exit Function
    Next i
    IsAllDigits = True
End Function

'================================ ACTUALISER LES TARIFS (WEB) =====================

Public Sub MAJ_Tarifs_Web()
    On Error GoTo EH
    If MsgBox("Tenter une verification de connexion aux sources de prix/facteurs ?" & vbCrLf & _
              "La saisie manuelle (feuille 09) reste toujours prioritaire.", _
              vbYesNo + vbQuestion, "Actualiser les tarifs") <> vbYes Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(SH_PARAM)

    ' URL de test (configurable en feuille 09)
    Dim url As String
    On Error Resume Next
    url = CStr(ws.Range("URL_TEST_MAJ").Value)
    On Error GoTo EH
    If Len(Trim(url)) = 0 Then url = "https://donnees.roulez-eco.fr/opendata/instantane"

    Dim ok As Boolean: ok = False
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If Not http Is Nothing Then
        http.Open "GET", url, False
        http.setRequestHeader "User-Agent", "Mozilla/5.0"
        http.send
        If http.Status = 200 Then ok = True
    End If
    On Error GoTo EH

    ' Horodatage de la colonne "Date MAJ" (col G) sur les blocs TARIFS et FACTEURS.
    ' Aucune VALEUR n'est ecrasee : seules les dates de verification sont inscrites.
    StampDates ws, "PRIX_ELEC", 7
    StampDates ws, "FE_ELEC", 7
    On Error Resume Next
    ws.Range("MAJ_GLOBALE").Value = Format(Now, "dd/mm/yyyy hh:nn")
    On Error GoTo EH

    If ok Then
        MsgBox "Connexion etablie avec la source." & vbCrLf & _
               "Pour un rafraichissement automatique des valeurs, utilisez la requete Power Query " & _
               "(voir le guide de connectivite). Sinon, ajustez les cellules jaunes en feuille 09." & vbCrLf & _
               "La date de verification a ete inscrite.", vbInformation
    Else
        MsgBox "Pas d'acces Internet ou source indisponible." & vbCrLf & _
               "Mettez a jour les prix/facteurs manuellement en feuille 09 (cellules jaunes).", vbExclamation
    End If
    Exit Sub
EH:
    MsgBox "Mise a jour web indisponible : " & Err.Description & vbCrLf & _
           "Saisissez les valeurs manuellement en feuille 09.", vbExclamation
End Sub

' Inscrit la date du jour dans la colonne 'colDate' pour un bloc de 7 lignes
' a partir de la ligne contenant 'anchor' en colonne A (ne touche aucune valeur).
Private Sub StampDates(ws As Worksheet, ByVal anchor As String, ByVal colDate As Long)
    Dim r0 As Long: r0 = FindRowText(ws, 1, anchor)
    If r0 = 0 Then Exit Sub
    Dim i As Long
    For i = r0 To r0 + 6
        If Len(Trim(CStr(ws.Cells(i, 1).Value))) > 0 Then ws.Cells(i, colDate).Value = Format(Date, "dd/mm/yyyy")
    Next i
End Sub

'================================ INSTALLER LES BOUTONS ===========================

Public Sub Installer_Outil()
    Dim ws As Worksheet: Set ws = Worksheets(SH_CMD)
    ' supprimer les anciens boutons de l'outil
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "BTN_" Then shp.Delete
    Next shp

    AddBtn ws, 1, 12, "1.  Importer le FEC", "Importer_FEC", RGB(27, 122, 82)
    AddBtn ws, 2, 14, "2.  Importer les DSN", "Importer_DSN", RGB(27, 122, 82)
    AddBtn ws, 3, 16, "3.  Generer la balance ESG", "Generer_Analyse", RGB(18, 70, 58)
    AddBtn ws, 4, 18, "4.  Actualiser le tableau de bord", "Actualiser_Dashboard", RGB(21, 101, 192)
    AddBtn ws, 5, 20, "5.  Actualiser les tarifs (web)", "MAJ_Tarifs_Web", RGB(106, 79, 179)
    AddBtn ws, 6, 22, "6.  Reinitialiser l'outil", "Reinitialiser_Outil", RGB(150, 60, 60)

    SetStatus "Outil installe. Renseignez les parametres d'import puis cliquez sur les boutons dans l'ordre."
    MsgBox "Boutons installes. L'outil est pret a l'emploi.", vbInformation
End Sub

Private Sub AddBtn(ws As Worksheet, ByVal idx As Long, ByVal rw As Long, _
                   ByVal caption As String, ByVal action As String, ByVal col As Long)
    Dim rng As Range: Set rng = ws.Range("A" & rw)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, rng.Left + 2, rng.Top + 1, 230, 26) ' 5 = rounded rectangle
    shp.Name = "BTN_" & idx
    shp.Fill.ForeColor.RGB = col
    shp.Line.Visible = msoFalse
    With shp.TextFrame
        .Characters.Text = caption
        .Characters.Font.Color = vbWhite
        .Characters.Font.Bold = True
        .Characters.Font.Size = 11
        .Characters.Font.Name = "Calibri"
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    shp.OnAction = action
End Sub

'================================ HELPERS =========================================

Private Function GetFolder() As String
    Dim f As String: f = Trim(CStr(Worksheets(SH_CMD).Range("C6").Value))
    If Len(f) = 0 Then
        MsgBox "Renseignez le dossier des fichiers en cellule C6 (feuille 01).", vbExclamation
        GetFolder = "": Exit Function
    End If
    If Right(f, 1) <> "\" Then f = f & "\"
    GetFolder = f
End Function

Private Function ResolveFecPath() As String
    Dim folder As String: folder = GetFolder()
    If Len(folder) = 0 Then Exit Function
    Dim fn As String: fn = Trim(CStr(Worksheets(SH_CMD).Range("C7").Value))
    Dim p As String
    If Len(fn) > 0 And (InStr(1, fn, "\") > 0 Or InStr(1, fn, ":") > 0) Then
        p = fn
    ElseIf Len(fn) > 0 Then
        p = folder & fn
    Else
        fn = Dir(folder & "*FEC*.txt")
        If Len(fn) = 0 Then fn = Dir(folder & "*.txt")
        If Len(fn) = 0 Then
            MsgBox "Aucun fichier FEC (.txt) trouve dans le dossier.", vbExclamation
            ResolveFecPath = "": Exit Function
        End If
        p = folder & fn
    End If
    If Dir(p) = "" Then
        MsgBox "Fichier FEC introuvable : " & p, vbCritical
        ResolveFecPath = ""
    Else
        ResolveFecPath = p
    End If
End Function

Private Function ReadTextSmart(ByVal path As String) As String
    On Error GoTo fallback
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile path
    ReadTextSmart = stm.ReadText: stm.Close
    Exit Function
fallback:
    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff
    ReadTextSmart = Input$(LOF(ff), ff)
    Close #ff
End Function

Private Function CleanField(ByVal s As String) As String
    s = Trim(CStr(s))
    If Len(s) >= 2 Then
        If Left(s, 1) = "'" And Right(s, 1) = "'" Then s = Mid(s, 2, Len(s) - 2)
        If Left(s, 1) = Chr(34) And Right(s, 1) = Chr(34) Then s = Mid(s, 2, Len(s) - 2)
    End If
    CleanField = s
End Function

Private Function ToDoubleFR(ByVal v As Variant) As Double
    On Error GoTo bad
    Dim s As String: s = CleanField(CStr(v))
    s = Replace(s, " ", ""): s = Replace(s, Chr(160), ""): s = Replace(s, ",", ".")
    If Len(s) = 0 Or s = "." Or s = "-" Then ToDoubleFR = 0 Else ToDoubleFR = CDbl(s)
    Exit Function
bad:
    ToDoubleFR = 0
End Function

Private Function MonthFromFilename(ByVal fn As String) As Long
    Dim i As Long, s As String
    For i = 1 To Len(fn) - 1
        s = Mid(fn, i, 2)
        If IsNumeric(s) Then
            If CLng(s) >= 1 And CLng(s) <= 12 Then
                If InStr(1, fn, "_" & s & "_", vbTextCompare) > 0 Or InStr(1, fn, "-" & s & "-", vbTextCompare) > 0 Then
                    MonthFromFilename = CLng(s): Exit Function
                End If
            End If
        End If
    Next i
    MonthFromFilename = 0
End Function

Private Function MoisNom(ByVal m As Long) As String
    Dim noms As Variant
    noms = Array("Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", _
                 "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre")
    If m >= 1 And m <= 12 Then MoisNom = noms(m - 1) Else MoisNom = "Mois " & m
End Function

Private Function NzL(d As Object, ByVal k As Variant) As Long
    If d.Exists(k) Then NzL = CLng(d(k)) Else NzL = 0
End Function

Private Function FindRowText(ws As Worksheet, ByVal col As Long, ByVal text As String) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim(CStr(ws.Cells(i, col).Value)), text, vbTextCompare) = 0 Then
            FindRowText = i: Exit Function
        End If
    Next i
    FindRowText = 0
End Function

Private Sub ClearFrom(ByVal sheetName As String, ByVal firstRow As Long)
    With Worksheets(sheetName)
        .Range(.Cells(firstRow, 1), .Cells(.Rows.Count, 30)).ClearContents
    End With
End Sub

Private Sub SetStatus(ByVal msg As String)
    On Error Resume Next
    Worksheets(SH_CMD).Range("A26").Value = msg
End Sub

Private Sub QuickSortKeys(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, midv As Variant, tmp As Variant
    low = first: high = last: midv = arr((first + last) \ 2)
    Do While low <= high
        Do While VBA.Val(arr(low)) < VBA.Val(midv): low = low + 1: Loop
        Do While VBA.Val(arr(high)) > VBA.Val(midv): high = high - 1: Loop
        If low <= high Then
            tmp = arr(low): arr(low) = arr(high): arr(high) = tmp
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSortKeys arr, first, high
    If low < last Then QuickSortKeys arr, low, last
End Sub
