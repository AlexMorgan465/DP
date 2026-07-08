Attribute VB_Name = "ModulePlanningFacto"

Option Explicit

'=====================================================================================
' GENERATEUR DE PLANNING - Projet "Facto"
'=====================================================================================
' HYPOTHESES retenues (a adapter si besoin dans le code ci-dessous) :
'
' 1. Tout le monde (collaborateurs ET manager) a la même valeur dans "Activité"
'    (ex: "Facto") -> cette colonne sert uniquement de filtre projet.
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

' 7. Planning special "4 agents" (Bara / Dami / El Moubarik / Intaj) :
'      - 2 binômes : {Bara, Dami} et {El Moubarik, Intaj}.
'      - Chaque semaine, un binôme est à 7h-17h (Lun-Mer) et l'autre à 8h-18h
'        (Lun-Mer). L'alternance se fait automatiquement selon la parité du
'        n° de semaine ISO (semaine impaire -> Bara/Dami = 7h-17h ;
'        semaine paire -> Bara/Dami = 8h-18h, et l'autre binôme inversement).
'      - Jeudi/Vendredi : au sein du binôme qui est en horaire "7h-X", rotation
'        de shift réduit entre les 2 membres : l'un fait Jeudi court/Vendredi
'        normal, l'autre fait Jeudi normal/Vendredi court.
'          Dami / El Moubarik = "court le jeudi" (jeudi -1h, vendredi normal)
'          Bara / Intaj       = "court le vendredi" (jeudi normal, vendredi -1h)
'      - Week-end OFF comme tout collaborateur.
'      - Pause déjeuner spécifique : 12h-12h30 (30 min, au lieu d'1h), une
'        seule personne à la fois parmi les 4, avec 30 min d'écart après
'        chaque retour de pause. L'ordre de passage est tiré au sort à
'        chaque génération du planning (cf. table "Rotation pause déjeuner").
'      -> Adaptez GetAgentRole() si les noms exacts dans la BDD diffèrent.
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

    projectName = InputBox("Nom du projet / de l'activité à générer (ex: Facto) :", _
                            "Génération du planning", "Facto")
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

    ' --- Rotation pause dejeuner des 4 agents (Bara/Dami/El Moubarik/Intaj) ---
    WriteLunchBreakRotation wsPlan, outRow + 5

    ' --- Section manager ---
    Dim managerStartRow As Long
    managerStartRow = outRow + 13
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
' Identifie le role special d'un agent parmi les 4 (Bara/Dami/El Moubarik/
' Intaj) a partir de son nom complet. 0 = agent normal (pas de regle
' speciale). Comparaison par mot-cle (insensible a la casse) pour rester
' tolerant au format exact du nom dans la BDD (ex: "Bara Ahmed").
'--------------------------------------------------------------------
Function GetAgentRole(ByVal nomComplet As String) As Integer
    Dim u As String
    u = UCase(nomComplet)
    If InStr(u, "BARA") > 0 Then
        GetAgentRole = 1        ' Bara      -> court le vendredi
    ElseIf InStr(u, "DAMI") > 0 Then
        GetAgentRole = 2        ' Dami      -> court le jeudi
    ElseIf InStr(u, "MOUBARIK") > 0 Then
        GetAgentRole = 3        ' El Moubarik -> court le jeudi
    ElseIf InStr(u, "INTAJ") > 0 Then
        GetAgentRole = 4        ' Intaj     -> court le vendredi
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

    ' Role special (Bara/Dami/El Moubarik/Intaj) pour horaires + pause dejeuner
    ' specifiques (cf. point 7 de l'en-tete du module)
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
            If agentRole > 0 And dayIndex <= 5 Then
                pauseDeduite = 0.5 ' pause dejeuner specifique 12h-12h30 (30 min)
            Else
                pauseDeduite = 1   ' pause dejeuner standard 13h-14h (1h)
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

        If agentRoleDI > 0 And dayIndex <= 5 Then
            ' ---- Planning special 4 agents (Bara/Dami/El Moubarik/Intaj) ----
            ' cf. point 7 de l'en-tete du module pour le detail de la regle
            Dim weekStartDI As Date, isoWkDI As Long, pairA_is7 As Boolean
            Dim isPairA As Boolean, pairIs7 As Boolean, baseEntreeDI As Integer
            weekStartDI = dayDate - (dayIndex - 1)
            isoWkDI = Application.WorksheetFunction.IsoWeekNum(weekStartDI)
            pairA_is7 = (isoWkDI Mod 2 = 1) ' semaine impaire -> Bara/Dami = 7h-17h

            isPairA = (agentRoleDI = 1 Or agentRoleDI = 2) ' Bara ou Dami
            If isPairA Then
                pairIs7 = pairA_is7
            Else
                pairIs7 = Not pairA_is7 ' El Moubarik / Intaj = binome inverse
            End If

            baseEntreeDI = IIf(pairIs7, 7, 8) ' 7h-17h ou 8h-18h

            Select Case dayIndex
                Case 1 To 3 ' Lundi a Mercredi : horaire plein du binome
                    entreeH = baseEntreeDI
                    sortieH = baseEntreeDI + 10
                Case 4, 5 ' Jeudi / Vendredi : rotation shift reduit (-1h)
                    Dim courtLeJeudi As Boolean
                    courtLeJeudi = (agentRoleDI = 2 Or agentRoleDI = 3) ' Dami / El Moubarik
                    entreeH = baseEntreeDI
                    If dayIndex = 4 Then ' Jeudi
                        sortieH = IIf(courtLeJeudi, baseEntreeDI + 9, baseEntreeDI + 10)
                    Else ' Vendredi
                        sortieH = IIf(courtLeJeudi, baseEntreeDI + 10, baseEntreeDI + 9)
                    End If
            End Select
            If comment = "" Then comment = "Rotation 4 agents (pause 12h-12h30)"
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
                Case 1 To 4: sortieH = 18
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
' Ecrit la table de rotation des pauses dejeuner (12h-12h30) des 4 agents
' Bara / Dami / El Moubarik / Intaj. Une seule personne a la fois, avec
' 30 min d'ecart apres chaque retour de pause avant le passage suivant :
'   Passage 1 : 12:00-12:30
'   Passage 2 : 13:00-13:30   (30 min d'ecart apres le retour du 1er)
'   Passage 3 : 14:00-14:30
'   Passage 4 : 15:00-15:30
' L'ordre de passage est tire au sort a chaque generation, pour chaque
' jour ouvre (Lundi a Vendredi ; week-end OFF donc pas de pause).
'--------------------------------------------------------------------
Sub WriteLunchBreakRotation(wsPlan As Worksheet, atRow As Long)
    Dim agentNames As Variant
    agentNames = Array("Bara", "Dami", "El Moubarik", "Intaj")

    With wsPlan
        .Range(.Cells(atRow, 1), .Cells(atRow, 5)).Merge
        .Cells(atRow, 1).Value = "Rotation pause déjeuner - Bara / Dami / El Moubarik / Intaj (aléatoire, 30 min d'écart après chaque retour)"
        .Cells(atRow, 1).Font.Bold = True
        .Cells(atRow, 1).Interior.Color = RGB(31, 73, 125)
        .Cells(atRow, 1).Font.Color = RGB(255, 255, 255)
    End With

    Dim headerRow As Long
    headerRow = atRow + 1
    wsPlan.Cells(headerRow, 1).Value = "Jour"
    Dim c As Long
    For c = 1 To 4
        wsPlan.Cells(headerRow, c + 1).Value = "Passage " & c
    Next c
    With wsPlan.Range(wsPlan.Cells(headerRow, 1), wsPlan.Cells(headerRow, 5))
        .Font.Bold = True
        .Interior.Color = RGB(217, 226, 243)
    End With

    Dim jours As Variant
    jours = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi")

    Randomize

    Dim d As Integer
    For d = 0 To 4 ' Lundi a Vendredi uniquement (week-end OFF)
        Dim r As Long
        r = headerRow + 1 + d
        wsPlan.Cells(r, 1).Value = jours(d)

        Dim order() As Integer
        order = ShuffleOrder(4)

        Dim slotStart As Date
        slotStart = TimeSerial(12, 0, 0)
        Dim i As Integer
        For i = 0 To 3
            Dim nm As String
            nm = agentNames(order(i))
            Dim slotEnd As Date
            slotEnd = slotStart + TimeSerial(0, 30, 0)
            wsPlan.Cells(r, i + 2).Value = nm & " (" & Format(slotStart, "hh:mm") & "-" & Format(slotEnd, "hh:mm") & ")"
            slotStart = slotStart + TimeSerial(1, 0, 0) ' pause 30 min + 30 min d'ecart
        Next i
    Next d
End Sub

'--------------------------------------------------------------------
' Renvoie un tableau (0 a n-1) contenant une permutation aleatoire des
' entiers 0..n-1 (algorithme de Fisher-Yates).
'--------------------------------------------------------------------
Function ShuffleOrder(ByVal n As Integer) As Integer()
    Dim arr() As Integer
    ReDim arr(0 To n - 1)
    Dim i As Integer
    For i = 0 To n - 1
        arr(i) = i
    Next i
    For i = n - 1 To 1 Step -1
        Dim j As Integer
        j = Int((i + 1) * Rnd)
        Dim tmp As Integer
        tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
    Next i
    ShuffleOrder = arr
End Function

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



