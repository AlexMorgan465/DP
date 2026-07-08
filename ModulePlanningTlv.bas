Attribute VB_Name = "ModulePlanningTlv"

Option Explicit

'=====================================================================================
' GENERATEUR DE PLANNING - Projet "TLV"
'=====================================================================================
' HYPOTHESES retenues (a adapter si besoin dans le code ci-dessous) :
'
' 1. Tout le monde (collaborateurs ET manager) a la même valeur dans "Activité"
'    (ex: "TLV") -> cette colonne sert uniquement de filtre projet.
'    Le rôle est déterminé par la colonne "MANAGER" de la BDD, qui vaut "OUI"
'    pour le(s) manager(s) et "NON" pour les collaborateurs.
'    -> Si votre BDD utilise un autre libellé que "OUI", adaptez EstManager().
'
' 2. Congé actif pour un jour donne si colonne "Congé" est renseignee et differente
'    de "NON", ET que le jour est compris entre "Congé D" et "Congé F".
'
' 3. Maladie active si colonne "Maladie" renseignee et differente de "NON", ET que
'    le jour est compris entre "Date D'Arret" et "DATE DE REPRISE".
'
' 4. Télétravail (TT) : si colonne "TT" renseignee et differente de "NON", ET jour
'    compris entre "TT D" et "TT F" -> le collaborateur reste aux memes horaires,
'    on ajoute juste la mention "Télétravail" en commentaire.
'    Le flag "TT" (O/N) affiche dans le planning est une simple reprise de la
'    colonne "TT" de la BDD.
'
' 5. Contrat / statut :
'      - Si "Date d'embauche" > jour  -> OFF, commentaire "Pas encore embauché"
'      - Si "Date de sortie" renseignee et <= jour -> OFF, "Contrat terminé"
'      - Si "Type de contrat" = "Terminé" ou "Sorti" -> OFF, "Contrat terminé"
'    (adaptez les libelles dans la fonction GetDayInfo si vos valeurs different)
'
' 6. Horaire par defaut :
'      - Collaborateurs : Lun-Jeu 8h-18h, Ven 8h-17h, Sam/Dim OFF
'      - Manager        : Lun-Ven 8h-17h, Sam/Dim OFF
'      - Pause dejeuner fixe 13h-14h (1h), deduite du total d'heures planifiees
'
' Priorite des regles (du + fort au + faible) : Contrat > Maladie > Congé > Défaut,
' puis annotation Télétravail si applicable et que le jour est travaille.

' 7. Rotation "Abdelaoui / Aziane" (2 plannings-types de 40h qui s'échangent
'    entre ces deux collaboratrices chaque semaine, selon la parité du n°
'    de semaine ISO) :
'      Planning A (40h) : Lun 8h-18h, Mar 9h-18h, Mer 9h-18h, Jeu 9h-18h,
'                          Ven 10h-18h, Sam/Dim OFF
'      Planning B (40h) : Lun OFF, Mar 8h-18h, Mer 8h-18h, Jeu 8h-17h,
'                          Ven 8h-17h, Sam 8h-14h (pas de pause dej ce jour-
'                          la, journée courte), Dim OFF
'    Semaine paire (ex: S28) -> Abdelaoui = Planning A, Aziane = Planning B
'    Semaine impaire (ex: S27) -> inversé (Abdelaoui = B, Aziane = A)
'    Diop et Sylla gardent le planning fixe standard (Lun-Ven 8h-17h,
'    week-end OFF = 40h), déjà couvert par l'horaire par défaut ci-dessus.
'    -> Adaptez GetAgentRole() si les noms exacts dans la BDD diffèrent.
'
' La macro :
'   a) Met a jour les colonnes horaires (LUN. Entrée ... DIM. Sortie) de la feuille
'      "BDD" pour chaque collaborateur/manager du projet et de la semaine choisis.
'   b) Construit / reconstruit la feuille de planning visuel, nommee comme le projet.
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

    projectName = InputBox("Nom du projet / de l'activité à générer (ex: TLV) :", _
                            "Génération du planning", "TLV")
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

    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "MATRICULE")).End(xlUp).Row
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

    ' --- Section collaborateurs ---
    outRow = WriteSectionHeader(wsPlan, 1, weekStart, "Collaborateur")
    For i = 1 To nCollab
        outRow = ProcessRow(wsBDD, wsPlan, headers, collabRows(i), weekStart, outRow, False)
    Next i

    ' --- Table de reference Shift / Pause dejeuner ---
    WriteShiftReferenceTable wsPlan, outRow + 2

    ' --- Section manager ---
    Dim managerStartRow As Long
    managerStartRow = outRow + 5
    outRow = WriteSectionHeader(wsPlan, managerStartRow, weekStart, "Manager")
    For i = 1 To nManager
        outRow = ProcessRow(wsBDD, wsPlan, headers, managerRows(i), weekStart, outRow, True)
    Next i

    wsPlan.Columns.AutoFit
    MsgBox "Planning généré avec succès dans la feuille '" & wsPlan.Name & _
           "'." & vbCrLf & "La BDD a été mise à jour pour la semaine du " & _
           Format(weekStart, "dd/mm/yyyy") & ".", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------
' Detection du role Manager - ADAPTER ICI si votre critère est different
' (se base sur la colonne "MANAGER" de la BDD : "OUI" = manager)
'--------------------------------------------------------------------
Function EstManager(ByVal managerFlagValue As String) As Boolean
    EstManager = (StrComp(Trim(managerFlagValue), "OUI", vbTextCompare) = 0)
End Function

'--------------------------------------------------------------------
' Identifie le role special d'un agent parmi Abdelaoui / Aziane (rotation
' de 2 plannings-types, cf. point 7 de l'en-tete du module) a partir de
' son nom complet. 0 = agent normal (horaire par defaut standard).
'--------------------------------------------------------------------
Function GetAgentRole(ByVal nomComplet As String) As Integer
    Dim u As String
    u = UCase(nomComplet)
    If InStr(u, "ABDELAOUI") > 0 Then
        GetAgentRole = 1
    ElseIf InStr(u, "AZIANE") > 0 Then
        GetAgentRole = 2
    Else
        GetAgentRole = 0
    End If
End Function

'--------------------------------------------------------------------
' Traite une ligne BDD (collaborateur ou manager) : calcule les horaires
' de la semaine, met à jour la BDD, écrit la ligne dans le planning.
' Retourne la prochaine ligne libre dans la feuille planning.
'--------------------------------------------------------------------
Function ProcessRow(wsBDD As Worksheet, wsPlan As Worksheet, headers As Object, _
                     rowBDD As Long, weekStart As Date, outRow As Long, _
                     isManager As Boolean) As Long

    Dim nomComplet As String, zone As String
    nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOMCOMPLET")).Value)
    If nomComplet = "" Then
        nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOM")).Value & " " & _
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

    ' Role special Abdelaoui/Aziane (rotation 2 plannings-types, point 7 de
    ' l'en-tete du module) : impacte uniquement la deduction de pause du
    ' samedi (journee courte 8h-14h sans pause dejeuner)
    Dim agentRole As Integer
    agentRole = GetAgentRole(nomComplet)

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
            wsPlan.Range(wsPlan.Cells(outRow, colEntreePlan), wsPlan.Cells(outRow, colSortiePlan)) _
                .Font.Color = RGB(200, 0, 0)
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

            Dim pauseDeduite As Double
            If agentRole > 0 And dayIndex = 6 Then
                pauseDeduite = 0 ' samedi : journee courte, pas de pause dejeuner
            Else
                pauseDeduite = 1
            End If
            totalHeures = totalHeures + (sortieH - entreeH - pauseDeduite)
        End If

        If comment <> "" And StrComp(comment, "RAS", vbTextCompare) <> 0 Then
            If Not comments.Exists(comment) Then comments.Add comment, True
        End If
    Next dayIndex

    ' OFF
    wsPlan.Cells(outRow, 17).Value = offCount
    ' NB heures planifiées (format cumulé sur plus de 24h)
    wsPlan.Cells(outRow, 18).Value = totalHeures / 24
    wsPlan.Cells(outRow, 18).NumberFormat = "[h]:mm:ss"
    ' TT (reprise simple de la colonne TT de la BDD)
    Dim ttRaw As String
    ttRaw = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TT")).Value)
    wsPlan.Cells(outRow, 19).Value = IIf(ttRaw <> "" And StrComp(ttRaw, "NON", vbTextCompare) <> 0, "O", "N")
    ' Commentaires
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
' Calcule l'horaire/l'etat d'un jour donne pour une ligne BDD
' Retourne un tableau : (entreeHeure, sortieHeure, isOff, commentaire)
'--------------------------------------------------------------------
Function GetDayInfo(wsBDD As Worksheet, headers As Object, rowBDD As Long, _
                     dayDate As Date, dayIndex As Integer, isManager As Boolean, _
                     Optional ByVal nomComplet As String = "") As Variant

    Dim entreeH As Integer, sortieH As Integer, isOff As Boolean, comment As String
    isOff = False
    comment = ""

    ' 1) Contrat / statut ------------------------------------------------
    Dim dEmbauche As Variant, dSortie As Variant, typeContrat As String
    dEmbauche = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEMBAUCHE")).Value
    dSortie = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDESORTIE")).Value
    typeContrat = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONTRAT")).Value)

    If IsDate(dEmbauche) Then
        If dayDate < CDate(dEmbauche) Then
            isOff = True: comment = "Pas encore embauché"
        End If
    End If
    If Not isOff And IsDate(dSortie) Then
        If dayDate >= CDate(dSortie) Then
            isOff = True: comment = "Contrat terminé"
        End If
    End If
    If Not isOff And (StrComp(typeContrat, "Terminé", vbTextCompare) = 0 _
                       Or StrComp(typeContrat, "Sorti", vbTextCompare) = 0) Then
        isOff = True: comment = "Contrat terminé"
    End If

    ' 2) Maladie ----------------------------------------------------------
    If Not isOff Then
        Dim maladieVal As String, dArret As Variant, dRepr As Variant
        maladieVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "MALADIE")).Value)
        dArret = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDARRET")).Value
        dRepr = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEREPRISE")).Value
        If maladieVal <> "" And StrComp(maladieVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartM As Boolean, okEndM As Boolean
            okStartM = (Not IsDate(dArret)) Or (dayDate >= CDate(dArret))
            okEndM = (Not IsDate(dRepr)) Or (dayDate <= CDate(dRepr))
            If okStartM And okEndM Then
                isOff = True: comment = "Maladie"
            End If
        End If
    End If

    ' 3) Congé --------------------------------------------------------------
    If Not isOff Then
        Dim congeVal As String, cD As Variant, cF As Variant, typeConge As String
        congeVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "CONGE")).Value)
        cD = wsBDD.Cells(rowBDD, GetCol(headers, "CONGED")).Value
        cF = wsBDD.Cells(rowBDD, GetCol(headers, "CONGEF")).Value
        typeConge = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONGE")).Value)
        If congeVal <> "" And StrComp(congeVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartC As Boolean, okEndC As Boolean
            okStartC = (Not IsDate(cD)) Or (dayDate >= CDate(cD))
            okEndC = (Not IsDate(cF)) Or (dayDate <= CDate(cF))
            If okStartC And okEndC Then
                isOff = True
                comment = IIf(typeConge <> "", typeConge, "Congé")
            End If
        End If
    End If

    ' 4) Horaire par defaut si aucune exception n'a mis la journee en OFF ---
    If Not isOff Then
        Dim agentRoleDI As Integer
        agentRoleDI = GetAgentRole(nomComplet)

        If agentRoleDI > 0 Then
            ' ---- Rotation 2 plannings-types Abdelaoui/Aziane (point 7) ----
            Dim weekStartDI As Date, isoWkDI As Long, weekIsEven As Boolean
            Dim role1HasPlanA As Boolean, thisRoleHasPlanA As Boolean
            weekStartDI = dayDate - (dayIndex - 1)
            isoWkDI = Application.WorksheetFunction.IsoWeekNum(weekStartDI)
            weekIsEven = (isoWkDI Mod 2 = 0)

            ' Semaine paire -> Abdelaoui (role 1) a le Planning A ; impaire -> inverse
            role1HasPlanA = weekIsEven
            thisRoleHasPlanA = IIf(agentRoleDI = 1, role1HasPlanA, Not role1HasPlanA)

            If thisRoleHasPlanA Then
                ' Planning A (40h) : Lun 8-18, Mar-Jeu 9-18, Ven 10-18, Sam/Dim OFF
                Select Case dayIndex
                    Case 1: entreeH = 8: sortieH = 18
                    Case 2 To 4: entreeH = 9: sortieH = 18
                    Case 5: entreeH = 10: sortieH = 18
                    Case Else: isOff = True ' Samedi / Dimanche
                End Select
            Else
                ' Planning B (40h) : Lun OFF, Mar-Mer 8-18, Jeu-Ven 8-17,
                ' Sam 8-14 (journee courte), Dim OFF
                Select Case dayIndex
                    Case 1: isOff = True ' Lundi OFF
                    Case 2, 3: entreeH = 8: sortieH = 18
                    Case 4, 5: entreeH = 8: sortieH = 17
                    Case 6: entreeH = 8: sortieH = 14
                    Case 7: isOff = True ' Dimanche
                End Select
            End If
            If comment = "" Then comment = "RAS"
        ElseIf isManager Then
            entreeH = 8
            If dayIndex <= 5 Then
                sortieH = 17
            Else
                isOff = True
            End If
            If comment = "" Then comment = "RAS"
        Else
            entreeH = 8
            Select Case dayIndex
                Case 1 To 4: sortieH = 17
                Case 5: sortieH = 17
                Case Else: isOff = True ' Samedi / Dimanche
            End Select
            If comment = "" Then comment = "RAS"
        End If
    Else
        If comment = "" Then comment = "RAS" ' week-end normal, sans cause particuliere
    End If

    ' 5) Télétravail : annotation seule, horaire inchangé ---------------------
    If Not isOff Then
        Dim ttVal As String, ttD As Variant, ttF As Variant
        ttVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TT")).Value)
        ttD = wsBDD.Cells(rowBDD, GetCol(headers, "TTD")).Value
        ttF = wsBDD.Cells(rowBDD, GetCol(headers, "TTF")).Value
        If ttVal <> "" And StrComp(ttVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartT As Boolean, okEndT As Boolean
            okStartT = (Not IsDate(ttD)) Or (dayDate >= CDate(ttD))
            okEndT = (Not IsDate(ttF)) Or (dayDate <= CDate(ttF))
            If okStartT And okEndT Then
                If StrComp(comment, "RAS", vbTextCompare) = 0 Then
                    comment = "Télétravail"
                Else
                    comment = comment & " / Télétravail"
                End If
            End If
        End If
    End If

    GetDayInfo = Array(entreeH, sortieH, isOff, comment)
End Function

'--------------------------------------------------------------------
' Clé de colonne BDD normalisée pour un jour/sens donné (ex: "LUNENTREE")
'--------------------------------------------------------------------
Function DayColKey(dayIndex As Integer, isEntree As Boolean) As String
    Dim prefixes As Variant
    prefixes = Array("LUN", "MAR", "MER", "JEU", "VEN", "SAM", "DIM")
    DayColKey = prefixes(dayIndex - 1) & IIf(isEntree, "ENTREE", "SORTIE")
End Function

Function DayLabel(dayIndex As Integer) As String
    Dim labels As Variant
    labels = Array("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    DayLabel = labels(dayIndex - 1)
End Function

'--------------------------------------------------------------------
' Prépare (recrée) la feuille de planning nommée d'après le projet
'--------------------------------------------------------------------
Function PreparePlanningSheet(ByVal projectName As String) As Worksheet
    Dim sheetName As String
    sheetName = CleanSheetName(projectName)

    Application.DisplayAlerts = False
    If SheetExists(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Delete
    End If
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set PreparePlanningSheet = ws
End Function

Function CleanSheetName(ByVal s As String) As String
    Dim r As String
    r = s
    Dim badChars As Variant, ch As Variant
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In badChars
        r = Replace(r, ch, "")
    Next ch
    If Len(r) > 31 Then r = Left(r, 31)
    CleanSheetName = r
End Function

Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'--------------------------------------------------------------------
' Ecrit les deux lignes d'en-tête (semaine + libellés colonnes) et
' retourne la première ligne de données disponible.
'--------------------------------------------------------------------
Function WriteSectionHeader(wsPlan As Worksheet, startRow As Long, weekStart As Date, _
                             roleLabel As String) As Long

    Dim r1 As Long, r2 As Long
    r1 = startRow
    r2 = startRow + 1

    Dim weekNum As Long
    On Error Resume Next
    weekNum = Application.WorksheetFunction.IsoWeekNum(weekStart)
    On Error GoTo 0

    Dim headerFill As Long, headerFont As Long
    headerFill = RGB(31, 73, 125)
    headerFont = RGB(255, 255, 255)

    With wsPlan.Range(wsPlan.Cells(r1, 1), wsPlan.Cells(r1, 2))
        .Merge
        .Value = "S" & weekNum
        .HorizontalAlignment = xlCenter
        .Interior.Color = headerFill
        .Font.Color = headerFont
        .Font.Bold = True
    End With

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim c1 As Long, c2 As Long
        c1 = 3 + (dayIndex - 1) * 2
        c2 = c1 + 1
        Dim dayDate As Date
        dayDate = weekStart + (dayIndex - 1)
        With wsPlan.Range(wsPlan.Cells(r1, c1), wsPlan.Cells(r1, c2))
            .Merge
            .Value = DayLabel(dayIndex) & " " & Format(dayDate, "dd mmmm yyyy")
            .HorizontalAlignment = xlCenter
            .Interior.Color = headerFill
            .Font.Color = headerFont
            .Font.Bold = True
        End With
        wsPlan.Cells(r2, c1).Value = "Début de shift"
        wsPlan.Cells(r2, c2).Value = "Fin de shift"
    Next dayIndex

    wsPlan.Cells(r2, 1).Value = "Zones"
    wsPlan.Cells(r2, 2).Value = roleLabel
    wsPlan.Cells(r2, 17).Value = "OFF"
    wsPlan.Cells(r2, 18).Value = "NB heures planifiées"
    wsPlan.Cells(r2, 19).Value = "TT"
    wsPlan.Cells(r2, 20).Value = "Commentaires"

    With wsPlan.Range(wsPlan.Cells(r2, 1), wsPlan.Cells(r2, 20))
        .Interior.Color = RGB(217, 226, 243)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    WriteSectionHeader = r2 + 1
End Function

'--------------------------------------------------------------------
' Petit tableau de reference Shift / Pause dejeuner
'--------------------------------------------------------------------
Sub WriteShiftReferenceTable(wsPlan As Worksheet, atRow As Long)
    With wsPlan
        .Range(.Cells(atRow, 4), .Cells(atRow, 5)).Merge
        .Cells(atRow, 4).Value = "Shift"
        .Range(.Cells(atRow, 6), .Cells(atRow, 7)).Merge
        .Cells(atRow, 6).Value = "Pause déjeuner"
        .Range(.Cells(atRow, 4), .Cells(atRow, 7)).Interior.Color = RGB(31, 73, 125)
        .Range(.Cells(atRow, 4), .Cells(atRow, 7)).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(atRow, 4), .Cells(atRow, 7)).Font.Bold = True
        .Range(.Cells(atRow, 4), .Cells(atRow, 7)).HorizontalAlignment = xlCenter

        .Cells(atRow + 1, 4).Value = "8:00"
        .Cells(atRow + 1, 5).Value = "18:00"
        .Cells(atRow + 1, 6).Value = "13:00"
        .Cells(atRow + 1, 7).Value = "14:00"
        .Range(.Cells(atRow + 1, 4), .Cells(atRow + 1, 7)).HorizontalAlignment = xlCenter
        .Range(.Cells(atRow + 1, 4), .Cells(atRow + 1, 7)).Font.Bold = True
    End With
End Sub

'--------------------------------------------------------------------
' Construit un dictionnaire {en-tête normalisé -> numéro de colonne}
' en lisant la ligne 1 de la feuille BDD.
'--------------------------------------------------------------------
Function GetHeaderMap(wsBDD As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long, c As Long
    lastCol = wsBDD.Cells(1, wsBDD.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        Dim key As String
        key = NormalizeHeader(CStr(wsBDD.Cells(1, c).Value))
        If key <> "" And Not dict.Exists(key) Then
            dict.Add key, c
        End If
    Next c

    Set GetHeaderMap = dict
End Function

Function GetCol(headers As Object, ByVal key As String) As Long
    If headers.Exists(key) Then
        GetCol = headers(key)
    Else
        Err.Raise vbObjectError + 1, , "Colonne introuvable dans la BDD pour la clé : " & key
    End If
End Function

'--------------------------------------------------------------------
' Normalise un en-tête : majuscules, sans accents, sans espaces/points/apostrophes
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
    r = Replace(r, ".", "")
    r = Replace(r, "'", "")
    r = Replace(r, "-", "")
    r = Replace(r, " ", "")
    NormalizeHeader = r
End Function



