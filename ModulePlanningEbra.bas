Option Explicit

Public Const NOM_FEUILLE_BDD As String = "BDD"

'--------------------------------------------------------------------
' POINT D'ENTREE
'--------------------------------------------------------------------
Sub GenererPlanningAccessibilite()

    Dim wsBDD As Worksheet, wsPlan As Worksheet
    Dim projectName As String, weekStartStr As String
    Dim weekStart As Date
    Dim lastRow As Long, r As Long
    Dim headers As Object

    On Error GoTo ErrHandler

    If Not SheetExists(NOM_FEUILLE_BDD) Then
        MsgBox "La feuille '" & NOM_FEUILLE_BDD & "' est introuvable.", vbCritical
        Exit Sub
    End If
    Set wsBDD = ThisWorkbook.Sheets(NOM_FEUILLE_BDD)

    projectName = InputBox("Nom du projet / de l'activité à générer (ex: Ebra presse) :", _
                            "Génération du planning", "Ebra presse")
    If Trim(projectName) = "" Then Exit Sub

    weekStartStr = InputBox("Date du LUNDI de la semaine à générer (jj/mm/aaaa) :", _
                             "Génération du planning", _
                             Format(Date - Weekday(Date, vbMonday) + 1, "dd/mm/yyyy"))
    If Trim(weekStartStr) = "" Then Exit Sub
    If Not IsDate(weekStartStr) Then
        MsgBox "Date invalide.", vbExclamation
        Exit Sub
    End If
    weekStart = CDate(weekStartStr)
    weekStart = weekStart - Weekday(weekStart, vbMonday) + 1 ' recale sur le lundi

    Set headers = GetHeaderMap(wsBDD)

    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "NOM")).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans la BDD.", vbExclamation
        Exit Sub
    End If

    Set wsPlan = PreparePlanningSheet(projectName)

    Dim colActivite As Long, colManager As Long
    colActivite = GetCol(headers, "ACTIVITE")
    colManager = GetCol(headers, "MANAGER")

    Dim collabRows() As Long, managerRows() As Long
    Dim nCollab As Long, nManager As Long
    nCollab = 0: nManager = 0
    ReDim collabRows(1 To lastRow)
    ReDim managerRows(1 To lastRow)

    For r = 2 To lastRow
        Dim actVal As String, managerFlag As String
        actVal = Trim(wsBDD.Cells(r, colActivite).Value)
        managerFlag = Trim(wsBDD.Cells(r, colManager).Value)
        If StrComp(actVal, projectName, vbTextCompare) = 0 Then
            If EstManager(managerFlag) Then
                nManager = nManager + 1
                managerRows(nManager) = r
            Else
                nCollab = nCollab + 1
                collabRows(nCollab) = r
            End If
        End If
    Next r

    If nCollab = 0 And nManager = 0 Then
        MsgBox "Aucune ligne trouvée pour l'activité '" & projectName & "' dans la BDD.", vbExclamation
        Exit Sub
    End If

    Dim outRow As Long
    Dim i As Long

    ' --- Section collaborateurs : shifts ---
    outRow = WriteSectionHeader(wsPlan, 1, weekStart, "Collaborateur")
    For i = 1 To nCollab
        outRow = ProcessRow(wsBDD, wsPlan, headers, collabRows(i), weekStart, outRow, False)
    Next i

    ' --- Section collaborateurs : pause dejeuner ---
    outRow = outRow + 1
    outRow = WritePauseSectionHeader(wsPlan, outRow, weekStart, "Collaborateur")
    For i = 1 To nCollab
        outRow = ProcessPauseRow(wsBDD, wsPlan, headers, collabRows(i), weekStart, outRow, False, i)
    Next i

    ' --- Section manager : shifts ---
    outRow = outRow + 2
    outRow = WriteSectionHeader(wsPlan, outRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True)
    Next i

    ' --- Section manager : pause dejeuner ---
    outRow = outRow + 1
    outRow = WritePauseSectionHeader(wsPlan, outRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessPauseRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True, i)
    Next i

    ' --- Table de reference des vagues de pause ---
    WriteShiftReferenceTable wsPlan, outRow + 2

    wsPlan.Columns.AutoFit
    MsgBox "Planning généré avec succès dans la feuille '" & wsPlan.Name & "'.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------
' Detection du role Manager
'--------------------------------------------------------------------
Function EstManager(ByVal managerFlagValue As String) As Boolean
    EstManager = (StrComp(Trim(managerFlagValue), "OUI", vbTextCompare) = 0)
End Function

'--------------------------------------------------------------------
' Détermine le type de planning d'un manager
' A = Chouifi / El Bahlouly : 7h-16h
' B = Mounaji / autres : 8h-17h
'--------------------------------------------------------------------
Function GetManagerScheduleType(ByVal nomComplet As String) As String
    Dim nom As String
    nom = LCase(Trim(nomComplet))
    If InStr(nom, "chouifi") > 0 Or InStr(nom, "el bahlouly") > 0 Or InStr(nom, "elbahlouly") > 0 Then
        GetManagerScheduleType = "A"
    Else
        GetManagerScheduleType = "B"
    End If
End Function

'--------------------------------------------------------------------
' Traite une ligne BDD
'--------------------------------------------------------------------
Function ProcessRow(wsBDD As Worksheet, wsPlan As Worksheet, headers As Object, _
                     rowBDD As Long, weekStart As Date, outRow As Long, _
                     isManager As Boolean) As Long

    Dim nomComplet As String, zone As String
    nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOMCOMPLET")).Value)
    If nomComplet = "" Then
        nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOM")).Value & " & _
                           wsBDD.Cells(rowBDD, GetCol(headers, "PRENOM")).Value)
    End If
    zone = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "ZONES")).Value)

    wsPlan.Cells(outRow, 1).Value = zone
    wsPlan.Cells(outRow, 2).Value = nomComplet
    wsPlan.Cells(outRow, 1).Font.Bold = True
    wsPlan.Cells(outRow, 2).Font.Bold = True

    Dim offCount As Long, totalHeures As Double
    Dim comments As Object
    Set comments = CreateObject("Scripting.Dictionary")
    offCount = 0
    totalHeures = 0

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim dayDate As Date
        dayDate = weekStart + (dayIndex - 1)

        Dim info As Variant
        info = GetDayInfo(wsBDD, headers, rowBDD, dayDate, dayIndex, isManager, nomComplet)
        Dim entreeH As Integer, sortieH As Integer, isOff As Boolean, comment As String
        entreeH = info(0): sortieH = info(1): isOff = info(2): comment = info(3)

        Dim colEntreeBDD As Long, colSortieBDD As Long
        colEntreeBDD = GetCol(headers, DayColKey(dayIndex, True))
        colSortieBDD = GetCol(headers, DayColKey(dayIndex, False))

        Dim colEntreePlan As Long, colSortiePlan As Long
        colEntreePlan = 3 + (dayIndex - 1) * 2
        colSortiePlan = colEntreePlan + 1

        If isOff Then
            wsBDD.Cells(rowBDD, colEntreeBDD).Value = "OFF"
            wsBDD.Cells(rowBDD, colSortieBDD).Value = "OFF"
            wsPlan.Cells(outRow, colEntreePlan).Value = "OFF"
            wsPlan.Cells(outRow, colSortiePlan).Value = "OFF"
            wsPlan.Range(wsPlan.Cells(outRow, colEntreePlan), wsPlan.Cells(outRow, colSortiePlan)).Font.Color = RGB(200, 0, 0)
            offCount = offCount + 1
        Else
            wsBDD.Cells(rowBDD, colEntreeBDD).Value = TimeSerial(entreeH, 0, 0)
            wsBDD.Cells(rowBDD, colEntreeBDD).NumberFormat = "hh""H"""
            wsBDD.Cells(rowBDD, colSortieBDD).Value = TimeSerial(sortieH, 0, 0)
            wsBDD.Cells(rowBDD, colSortieBDD).NumberFormat = "hh""H"""

            wsPlan.Cells(outRow, colEntreePlan).Value = TimeSerial(entreeH, 0, 0)
            wsPlan.Cells(outRow, colEntreePlan).NumberFormat = "h:mm"
            wsPlan.Cells(outRow, colSortiePlan).Value = TimeSerial(sortieH, 0, 0)
            wsPlan.Cells(outRow, colSortiePlan).NumberFormat = "h:mm"

            ' Déduction pause 1h seulement Lun-Ven si > 5h
            Dim duree As Integer
            duree = sortieH - entreeH
            If dayIndex <= 5 And duree > 5 Then
                totalHeures = totalHeures + (duree - 1)
            Else
                totalHeures = totalHeures + duree
            End If
        End If

        If comment <> "" And StrComp(comment, "RAS", vbTextCompare) <> 0 Then
            If Not comments.Exists(comment) Then comments.Add comment, True
        End If
    Next dayIndex

    wsPlan.Cells(outRow, 17).Value = offCount
    wsPlan.Cells(outRow, 18).Value = totalHeures / 24
    wsPlan.Cells(outRow, 18).NumberFormat = "[h]:mm:ss"

    Dim ttRaw As String
    ttRaw = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TT")).Value)
    wsPlan.Cells(outRow, 19).Value = IIf(ttRaw <> "" And StrComp(ttRaw, "NON", vbTextCompare) <> 0, "O", "N")

    If comments.Count = 0 Then
        wsPlan.Cells(outRow, 20).Value = "RAS"
    Else
        Dim k As Variant, txt As String
        For Each k In comments.Keys
            txt = txt & IIf(txt = "", "", " / ") & k
        Next k
        wsPlan.Cells(outRow, 20).Value = txt
    End If

    ProcessRow = outRow + 1
End Function

'--------------------------------------------------------------------
' Calcule l'horaire d'un jour
'--------------------------------------------------------------------
Function GetDayInfo(wsBDD As Worksheet, headers As Object, rowBDD As Long, _
                     dayDate As Date, dayIndex As Integer, isManager As Boolean, _
                     Optional nomComplet As String = "") As Variant

    Dim entreeH As Integer, sortieH As Integer, isOff As Boolean, comment As String
    isOff = False: comment = ""

    ' Contrats, Maladie, Congé
    Dim dEmbauche As Variant, dSortie As Variant, typeContrat As String
    dEmbauche = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEMBAUCHE")).Value
    dSortie = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDESORTIE")).Value
    typeContrat = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONTRAT")).Value)

    If IsDate(dEmbauche) Then If dayDate < CDate(dEmbauche) Then isOff = True: comment = "Pas encore embauché"
    If Not isOff And IsDate(dSortie) Then If dayDate >= CDate(dSortie) Then isOff = True: comment = "Contrat terminé"
    If Not isOff And (StrComp(typeContrat, "Terminé", vbTextCompare) = 0 Or StrComp(typeContrat, "Sorti", vbTextCompare) = 0) Then isOff = True: comment = "Contrat terminé"

    If Not isOff Then
        Dim maladieVal As String, dArret As Variant, dRepr As Variant
        maladieVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "MALADIE")).Value)
        dArret = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDARRET")).Value
        dRepr = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEREPRISE")).Value
        If maladieVal <> "" And StrComp(maladieVal, "NON", vbTextCompare) <> 0 Then
            If (Not IsDate(dArret) Or dayDate >= CDate(dArret)) And (Not IsDate(dRepr) Or dayDate <= CDate(dRepr)) Then
                isOff = True: comment = "Maladie"
            End If
        End If
    End If

    If Not isOff Then
        Dim congeVal As String, cD As Variant, cF As Variant, typeConge As String
        congeVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "CONGE")).Value)
        cD = wsBDD.Cells(rowBDD, GetCol(headers, "CONGED")).Value
        cF = wsBDD.Cells(rowBDD, GetCol(headers, "CONGEF")).Value
        typeConge = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONGE")).Value)
        If congeVal <> "" And StrComp(congeVal, "NON", vbTextCompare) <> 0 Then
            If (Not IsDate(cD) Or dayDate >= CDate(cD)) And (Not IsDate(cF) Or dayDate <= CDate(cF)) Then
                isOff = True: comment = IIf(typeConge <> "", typeConge, "Congé")
            End If
        End If
    End If

    ' Horaire par défaut
    If Not isOff Then
        If isManager Then
            Dim mgrType As String
            mgrType = GetManagerScheduleType(nomComplet)
            If mgrType = "A" Then ' Chouifi / El Bahlouly
                Select Case dayIndex
                    Case 1 To 5: entreeH = 7: sortieH = 16
                    Case 6: entreeH = 7: sortieH = 11
                    Case Else: isOff = True
                End Select
            Else ' Mounaji et autres
                Select Case dayIndex
                    Case 1 To 5: entreeH = 8: sortieH = 17
                    Case Else: isOff = True
                End Select
            End If
        Else ' Collaborateur
            Select Case dayIndex
                Case 1 To 5: entreeH = 7: sortieH = 16
                Case 6: entreeH = 7: sortieH = 11
                Case Else: isOff = True
            End Select
        End If
        If comment = "" Then comment = "RAS"
    Else
        If comment = "" Then comment = "RAS"
    End If

    ' Télétravail
    If Not isOff Then
        Dim ttVal As String, ttD As Variant, ttF As Variant
        ttVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TT")).Value)
        ttD = wsBDD.Cells(rowBDD, GetCol(headers, "TTD")).Value
        ttF = wsBDD.Cells(rowBDD, GetCol(headers, "TTF")).Value
        If ttVal <> "" And StrComp(ttVal, "NON", vbTextCompare) <> 0 Then
            If (Not IsDate(ttD) Or dayDate >= CDate(ttD)) And (Not IsDate(ttF) Or dayDate <= CDate(ttF)) Then
                comment = IIf(StrComp(comment, "RAS", vbTextCompare) = 0, "Télétravail", comment & " / Télétravail")
            End If
        End If
    End If

    GetDayInfo = Array(entreeH, sortieH, isOff, comment)
End Function

'--------------------------------------------------------------------
' Pause Managers selon les règles
'--------------------------------------------------------------------
Sub GetManagerPause(ByVal nomComplet As String, ByVal dayIndex As Integer, _
                     ByRef h1 As Integer, ByRef m1 As Integer, _
                     ByRef h2 As Integer, ByRef m2 As Integer)
    Dim nom As String
    nom = LCase(Trim(nomComplet))

    If InStr(nom, "chouifi") > 0 Then
        ' Chouifi : Lun-Jeu 12h-13h, Ven 13h30-14h30
        If dayIndex >= 1 And dayIndex <= 4 Then
            h1 = 12: m1 = 0: h2 = 13: m2 = 0
        ElseIf dayIndex = 5 Then
            h1 = 13: m1 = 30: h2 = 14: m2 = 30
        End If
    ElseIf InStr(nom, "el bahlouly") > 0 Or InStr(nom, "elbahlouly") > 0 Then
        ' El Bahlouly : Lun-Ven 13h-14h
        h1 = 13: m1 = 0: h2 = 14: m2 = 0
    Else
        ' Mounaji et autres : Lun-Jeu 13h-14h, Ven 12h30-13h30
        If dayIndex >= 1 And dayIndex <= 4 Then
            h1 = 13: m1 = 0: h2 = 14: m2 = 0
        ElseIf dayIndex = 5 Then
            h1 = 12: m1 = 30: h2 = 13: m2 = 30
        End If
    End If
End Sub

'--------------------------------------------------------------------
' NormalizeHeader CORRIGE
'--------------------------------------------------------------------
Function NormalizeHeader(ByVal s As String) As String
    Dim r As String
    r = UCase(Trim(s))
    r = Replace(r, "É", "E"): r = Replace(r, "È", "E"): r = Replace(r, "Ê", "E"): r = Replace(r, "Ë", "E")
    r = Replace(r, "À", "A"): r = Replace(r, "Â", "A")
    r = Replace(r, "Ô", "O")
    r = Replace(r, "Î", "I"): r = Replace(r, "Ï", "I")
    r = Replace(r, "Ù", "U"): r = Replace(r, "Û", "U")
    r = Replace(r, "Ç", "C")
    r = Replace(r, ".", ""): r = Replace(r, "'", ""): r = Replace(r, "-", ""): r = Replace(r, " ", "")
    NormalizeHeader = r
End Function

' --- Le reste des fonctions : DayColKey, DayLabel, PreparePlanningSheet,
' CleanSheetName, SheetExists, WriteSectionHeader, WritePauseSectionHeader,
' ProcessPauseRow, GetPauseWaveIndex, GetWaveTimes, WriteShiftReferenceTable, GetHeaderMap, GetCol
' --- SONT IDENTIQUES A TON CODE D'ORIGINE ---
