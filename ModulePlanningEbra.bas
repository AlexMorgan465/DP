Attribute VB_Name = "ModulePlanningEbra"

Option Explicit

'=====================================================================================
' GENERATEUR DE PLANNING - Projet "Ebra presse" (version adaptée)
'=====================================================================================
' ... (même entête que précédemment)
'=====================================================================================

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

    projectName = InputBox("Nom du projet / de l'activité ŕ générer (ex: Ebra presse) :", _
                            "Génération du planning", "Ebra presse")
    If Trim(projectName) = "" Then Exit Sub

    weekStartStr = InputBox("Date du LUNDI de la semaine ŕ générer (jj/mm/aaaa) :", _
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

    ' --- REMPLACEMENT DE "MATRICULE" PAR "NOM" ---
    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "NOM")).End(xlUp).Row
    ' ---------------------------------------------

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
        MsgBox "Aucune ligne trouvée pour l'activité '" & projectName & _
               "' (ou 'Manager') dans la BDD.", vbExclamation
        Exit Sub
    End If

    Dim outRow As Long
    Dim i As Long

    ' --- Section collaborateurs : shifts ---
    outRow = WriteSectionHeader(wsPlan, 1, weekStart, "Collaborateur")
    For i = 1 To nCollab
        outRow = ProcessRow(wsBDD, wsPlan, headers, collabRows(i), weekStart, outRow, False)
    Next i

    ' --- Section collaborateurs : pause dejeuner (rotation par vagues) ---
    outRow = outRow + 1
    outRow = WritePauseSectionHeader(wsPlan, outRow, weekStart, "Collaborateur")
    For i = 1 To nCollab
        outRow = ProcessPauseRow(wsBDD, wsPlan, headers, collabRows(i), weekStart, outRow, False, i)
    Next i

    ' --- Section manager : shifts ---
    outRow = outRow + 2
    Dim managerStartRow As Long
    managerStartRow = outRow
    outRow = WriteSectionHeader(wsPlan, managerStartRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True)
    Next i

    ' --- Section manager : pause dejeuner (personnalisée selon le nom) ---
    outRow = outRow + 1
    outRow = WritePauseSectionHeader(wsPlan, outRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessPauseRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True, i)
    Next i

    ' --- Table de reference des vagues de pause (3 vagues pour collaborateurs) ---
    WriteShiftReferenceTable wsPlan, outRow + 2

    wsPlan.Columns.AutoFit
    MsgBox "Planning généré avec succčs dans la feuille '" & wsPlan.Name & _
           "'." & vbCrLf & "La BDD a été mise ŕ jour pour la semaine du " & _
           Format(weekStart, "dd/mm/yyyy") & "." & vbCrLf & _
           "Les pauses déjeuner des collaborateurs tournent chaque semaine sur 3 vagues " & _
           "(11h-12h / 11h30-12h30 / 12h-13h). Les pauses des managers sont fixes selon leur nom.", _
           vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------
' Le reste des fonctions est identique à la version précédente
' (EstManager, GetManagerScheduleType, ProcessRow, GetDayInfo, etc.)
'--------------------------------------------------------------------
' ... (copiez ici toutes les autres fonctions inchangées)
