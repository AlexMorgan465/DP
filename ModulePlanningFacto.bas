Attribute VB_Name = "ModulePlanningFacto"
Option Explicit

Public Const NOM_FEUILLE_BDD As String = "BDD"

Sub GenererPlanningFacto()
    Dim wsBDD As Worksheet, wsPlan As Worksheet
    Dim projectName As String, weekStartStr As String
    Dim weekStart As Date
    Dim lastRow As Long, r As Long
    Dim headers As Object

    On Error GoTo ErrHandler
    Set wsBDD = ThisWorkbook.Sheets(NOM_FEUILLE_BDD)

    projectName = "Facto"
    weekStartStr = InputBox("Date du LUNDI de la semaine à générer (jj/mm/aaaa) :", _
                             "Génération du planning Facto", _
                             Format(Date - Weekday(Date, vbMonday) + 1, "dd/mm/yyyy"))
    If Trim(weekStartStr) = "" Then Exit Sub
    weekStart = CDate(weekStartStr)
    weekStart = weekStart - Weekday(weekStart, vbMonday) + 1

    Set headers = GetHeaderMap(wsBDD)
    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "NOM")).End(xlUp).Row
    Set wsPlan = PreparePlanningSheet(projectName)

    Dim colActivite As Long, colManager As Long
    colActivite = GetCol(headers, "ACTIVITE")
    colManager = GetCol(headers, "MANAGER")

    Dim factRows() As Long, managerRows() As Long, otherRows() As Long
    Dim nFact As Long, nManager As Long, nOther As Long
    nFact = 0: nManager = 0: nOther = 0
    ReDim factRows(1 To lastRow): ReDim managerRows(1 To lastRow): ReDim otherRows(1 To lastRow)

    For r = 2 To lastRow
        Dim actVal As String, managerFlag As String, nom As String
        actVal = Trim(wsBDD.Cells(r, colActivite).Value)
        managerFlag = Trim(wsBDD.Cells(r, colManager).Value)
        nom = Trim(wsBDD.Cells(r, GetCol(headers, "NOMCOMPLET")).Value)
        If StrComp(actVal, projectName, vbTextCompare) = 0 Then
            If EstManager(managerFlag) Then
                nManager = nManager + 1: managerRows(nManager) = r
            ElseIf IsSpecialAgent(nom) Then
                nFact = nFact + 1: factRows(nFact) = r
            Else
                nOther = nOther + 1: otherRows(nOther) = r
            End If
        End If
    Next r

    Dim outRow As Long
    ' --- Section Agents Facto Spéciaux ---
    outRow = WriteSectionHeader(wsPlan, 1, weekStart, "Agents Facto")
    Dim i As Long
    For i = 1 To nFact
        outRow = ProcessRowFacto(wsBDD, wsPlan, headers, factRows(i), weekStart, outRow)
    Next i

    ' --- Section Autres Collaborateurs ---
    outRow = outRow + 2
    outRow = WriteSectionHeader(wsPlan, outRow, weekStart, "Collaborateur")
    For i = 1 To nOther
        outRow = ProcessRow(wsBDD, wsPlan, headers, otherRows(i), weekStart, outRow, False)
    Next i

    ' --- Section Manager ---
    outRow = outRow + 2
    outRow = WriteSectionHeader(wsPlan, outRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True)
    Next i

    wsPlan.Columns.AutoFit
    MsgBox "Planning Facto généré avec succès.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Function ProcessRowFacto(wsBDD As Worksheet, wsPlan As Worksheet, headers As Object, _
                     rowBDD As Long, weekStart As Date, outRow As Long) As Long
    Dim nomComplet As String, zone As String
    nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOMCOMPLET")).Value)
    If nomComplet = "" Then nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOM")).Value & " " & wsBDD.Cells(rowBDD, GetCol(headers, "PRENOM")).Value)
    zone = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "ZONES")).Value)

    wsPlan.Cells(outRow, 1).Value = zone: wsPlan.Cells(outRow, 2).Value = nomComplet
    wsPlan.Cells(outRow, 1).Font.Bold = True: wsPlan.Cells(outRow, 2).Font.Bold = True

    Dim offCount As Long, totalHeures As Double: offCount = 0: totalHeures = 0
    Dim comments As Object: Set comments = CreateObject("Scripting.Dictionary")

    Dim weekNum As Long: weekNum = Application.WorksheetFunction.IsoWeekNum(weekStart)
    Dim isEvenWeek As Boolean: isEvenWeek = (weekNum Mod 2 = 0)

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim dayDate As Date: dayDate = weekStart + (dayIndex - 1)
        Dim info As Variant: info = GetDayInfoFacto(wsBDD, headers, rowBDD, dayDate, dayIndex, nomComplet, isEvenWeek)
        Dim entreeH As Integer, entreeM As Integer, sortieH As Integer, sortieM As Integer, isOff As Boolean, comment As String
        entreeH = info(0): entreeM = info(1): sortieH = info(2): sortieM = info(3): isOff = info(4): comment = info(5)

        Dim colEntreePlan As Long, colSortiePlan As Long
        colEntreePlan = 3 + (dayIndex - 1) * 2: colSortiePlan = colEntreePlan + 1

        If isOff Then
            wsPlan.Cells(outRow, colEntreePlan).Value = "OFF": wsPlan.Cells(outRow, colSortiePlan).Value = "OFF"
            wsPlan.Range(wsPlan.Cells(outRow, colEntreePlan), wsPlan.Cells(outRow, colSortiePlan)).Font.Color = RGB(200, 0, 0)
            offCount = offCount + 1
        Else
            wsPlan.Cells(outRow, colEntreePlan).Value = TimeSerial(entreeH, entreeM, 0): wsPlan.Cells(outRow, colEntreePlan).NumberFormat = "h:mm"
            wsPlan.Cells(outRow, colSortiePlan).Value = TimeSerial(sortieH, sortieM, 0): wsPlan.Cells(outRow, colSortiePlan).NumberFormat = "h:mm"
            totalHeures = totalHeures + ((sortieH + sortieM / 60) - (entreeH + entreeM / 60) - 0.5) ' -30min pause
        End If
        If comment <> "" And StrComp(comment, "RAS", vbTextCompare) <> 0 Then If Not comments.Exists(comment) Then comments.Add comment, True
    Next dayIndex

    wsPlan.Cells(outRow, 17).Value = offCount
    wsPlan.Cells(outRow, 18).Value = totalHeures / 24: wsPlan.Cells(outRow, 18).NumberFormat = "[h]:mm:ss"
    If comments.Count = 0 Then wsPlan.Cells(outRow, 20).Value = "RAS" Else Dim k As Variant, txt As String: For Each k In comments.Keys: txt = txt & IIf(txt = "", "", " / ") & k: Next: wsPlan.Cells(outRow, 20).Value = txt
    ProcessRowFacto = outRow + 1
End Function

Function GetDayInfoFacto(wsBDD As Worksheet, headers As Object, rowBDD As Long, _
                     dayDate As Date, dayIndex As Integer, nomComplet As String, isEvenWeek As Boolean) As Variant
    Dim entreeH As Integer, entreeM As Integer, sortieH As Integer, sortieM As Integer, isOff As Boolean, comment As String
    isOff = False: comment = "RAS": entreeM = 0: sortieM = 0

    ' Congé/Maladie/Contrat
    Dim dEmbauche As Variant: dEmbauche = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEMBAUCHE")).Value
    If IsDate(dEmbauche) And dayDate < CDate(dEmbauche) Then isOff = True: comment = "Pas encore embauché"

    If Not isOff And dayIndex >= 6 Then isOff = True ' Weekend OFF

    If Not isOff Then
        ' Déterminer le groupe de la semaine
        Dim isGroupA As Boolean
        If IsInArray(UCase(nomComplet), Array("BARA", "DAMI")) Then
            isGroupA = isEvenWeek ' Bara+Dami = 7-17 semaine paire
        Else
            isGroupA = Not isEvenWeek ' El Moubarik+Intaj = 7-17 semaine impaire
        End If

        ' Horaire de base
        If isGroupA Then ' Groupe 7-17
            entreeH = 7: sortieH = 17
            ' Rotation Jeudi/Vendredi
            Dim isFirstInGroup As Boolean: isFirstInGroup = (InStr(UCase(nomComplet), "BARA") > 0 Or InStr(UCase(nomComplet), "EL MOUB") > 0)
            If isEvenWeek Then isFirstInGroup = Not isFirstInGroup ' Inverse la rotation chaque semaine
            If dayIndex = 4 Then ' Jeudi
                If isFirstInGroup Then sortieH = 16 Else sortieH = 17
            ElseIf dayIndex = 5 Then ' Vendredi
                If isFirstInGroup Then sortieH = 17 Else sortieH = 16
            End If
        Else ' Groupe 8-18
            entreeH = 8: sortieH = 18
            Dim isFirstInGroupB As Boolean: isFirstInGroupB = (InStr(UCase(nomComplet), "DAMI") > 0 Or InStr(UCase(nomComplet), "INTAJ") > 0)
            If isEvenWeek Then isFirstInGroupB = Not isFirstInGroupB
            If dayIndex = 4 Then
                If isFirstInGroupB Then sortieH = 17 Else sortieH = 18
            ElseIf dayIndex = 5 Then
                If isFirstInGroupB Then sortieH = 18 Else sortieH = 17
            End If
        End If
    End If
    GetDayInfoFacto = Array(entreeH, entreeM, sortieH, sortieM, isOff, comment)
End Function

' --- A AJOUTER DANS ProcessRow normal pour les autres ---
Function ProcessRow(wsBDD As Worksheet, wsPlan As Worksheet, headers As Object, rowBDD As Long, weekStart As Date, outRow As Long, isManager As Boolean) As Long
    ' Ton code ProcessRow d'origine mais avec pause 1h et horaire Lun-Jeu 8-18 Ven 8-17
    '... garde ton code existant ici...
    ProcessRow = outRow + 1
End Function

Function IsSpecialAgent(ByVal nomComplet As String) As Boolean
    Dim n As String: n = UCase(nomComplet)
    IsSpecialAgent = (InStr(n, "BARA") > 0 Or InStr(n, "DAMI") > 0 Or InStr(n, "EL MOUB") > 0 Or InStr(n, "INTAJ") > 0)
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Integer: For i = LBound(arr) To UBound(arr): If arr(i) = val Then IsInArray = True: Exit Function: Next
End Function

Function NormalizeHeader(ByVal s As String) As String
    Dim r As String: r = UCase(Trim(s))
    r = Replace(r, "É", "E"): r = Replace(r, "È", "E"): r = Replace(r, "Ê", "E"): r = Replace(r, "À", "A"): r = Replace(r, "Ç", "C")
    r = Replace(r, ".", ""): r = Replace(r, "'", ""): r = Replace(r, "-", ""): r = Replace(r, " ", "")
    NormalizeHeader = r
End Function

' Garde aussi: EstManager, DayColKey, DayLabel, PreparePlanningSheet, CleanSheetName, SheetExists, WriteSectionHeader, GetHeaderMap, GetCol
