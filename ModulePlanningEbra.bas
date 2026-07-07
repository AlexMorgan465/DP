Attribute VB_Name = "ModulePlanningEbra"

Option Explicit

'=====================================================================================
' GENERATEUR DE PLANNING - Projet "Ebra presse" (version adaptée)
'=====================================================================================
' ... (en-tête)
'=====================================================================================

Public Const NOM_FEUILLE_BDD As String = "BDD"
Public Const COL_PROJET As String = "ACTIVITE"   ' <-- ADAPTEZ ICI (ex: "PROJET", "ACTIVITE", ...)

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
    weekStart = weekStart - Weekday(weekStart, vbMonday) + 1

    Set headers = GetHeaderMap(wsBDD)

    ' Utilisation de la colonne "NOM" pour la dernière ligne
    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "NOM")).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Aucune donnée trouvée dans la BDD.", vbExclamation
        Exit Sub
    End If

    Set wsPlan = PreparePlanningSheet(projectName)

    Dim colActivite As Long, colManager As Long
    colActivite = GetCol(headers, COL_PROJET)        ' <-- Utilisation de la constante
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

    ' --- Reste du code inchangé (appels à ProcessRow, ProcessPauseRow, etc.) ---
    ' ... (copiez ici toutes les autres fonctions telles que fournies précédemment)

    wsPlan.Columns.AutoFit
    MsgBox "Planning généré avec succès dans la feuille '" & wsPlan.Name & "'." & vbCrLf & _
           "La BDD a été mise à jour pour la semaine du " & Format(weekStart, "dd/mm/yyyy") & ".", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

' --- Toutes les autres fonctions (EstManager, ProcessRow, GetDayInfo, GetHeaderMap, etc.) restent identiques ---
