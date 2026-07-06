VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFGL 
   Caption         =   "UserForm2"
   ClientHeight    =   10104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13464
   OleObjectBlob   =   "UFGL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' UFGL - Saisie manuelle du planning GOOGLE LEADS
' ============================================================
' CE FICHIER EST DU CODE, PAS UN .frm COMPLET.
' Anthropic/Claude ne peut pas generer le binaire .frx d'un
' UserForm (positions/polices des controles). Vous devez donc
' creer le formulaire dans l'editeur VBA, puis coller ce code
' dans son module de code. Voir INSTRUCTIONS.md pour la liste
' exacte des controles a creer et leurs noms.
'
' WORKFLOW :
'   1) On selectionne (Ctrl/Shift-clic) un ou plusieurs collaborateurs
'      dans lstCollabs, on choisit une Vague / un Shift Reduit / un mode
'      Dimanche, puis on clique "Planifier" -> la condition est stockee
'      pour ces collaborateurs (colonne 2 de la liste mise a jour).
'   2) On peut recommencer avec une autre selection et d'autres
'      conditions autant de fois que necessaire.
'   3) "Generer" ecrit dans la feuille GOOGLE LEADS toutes les
'      affectations planifiees d'un coup.
' ============================================================

Option Explicit

Private m_semaines() As Date            ' index -> Lundi de la semaine correspondante
Private m_affectations() As AffectationGL
Private m_nbAffect As Integer

Private Sub UserForm_Initialize()
    Me.Caption = "Google Leads - Saisie du planning"

    ' --- Combo Semaine : semaine courante + 6 suivantes ---
    cboSemaine.Clear
    ReDim m_semaines(0 To 6)
    Dim i As Integer
    Dim wd As Integer: wd = Weekday(Date, vbMonday)
    Dim lundiCourant As Date: lundiCourant = Date - (wd - 1)
    For i = 0 To 6
        Dim lundi As Date: lundi = lundiCourant + (i * 7)
        m_semaines(i) = lundi
        Dim sem As Integer: sem = Application.WorksheetFunction.WeekNum(lundi, 2)
        cboSemaine.AddItem "S" & sem & "  (" & Format(lundi, "dd/mm") & " - " & Format(lundi + 6, "dd/mm") & ")"
    Next i
    cboSemaine.ListIndex = 0

    ' --- Vague (shift Lundi -> Samedi), par defaut 7h-17h ---
    optVague1.Value = True
    txtVagueSpec.Text = ""
    txtVagueSpec.Enabled = False

    ' --- Shift Reduit (applique le Dimanche), par defaut 11h-20h ---
    optRed1.Value = True
    txtRedSpec.Text = ""
    txtRedSpec.Enabled = False

    ' --- Dimanche : ni TT ni OFF coche par defaut (= shift reduit normal) ---
    optTT.Value = False
    optOFF.Value = False

    ' --- Liste des collaborateurs GOOGLE LEADS (depuis Utilisateurs) ---
    ChargerListeCollabs

    m_nbAffect = 0
    ReDim m_affectations(1 To 1)
End Sub

' Remplit lstCollabs avec les collaborateurs dont le projet = GOOGLE LEADS.
' Colonne 1 = Nom complet, Colonne 2 = condition planifiee (vide au depart).
' Pense a activer, sur lstCollabs : ColumnCount = 2, MultiSelect = fmMultiSelectExtended.
Private Sub ChargerListeCollabs()
    lstCollabs.Clear
    If Not FeuilleExiste("Utilisateurs") Then Exit Sub

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)

    Dim i As Integer
    For i = 1 To nbCollab
        If UCase(Trim(collaborateurs(i).projet)) = "GOOGLE LEADS" Then
            lstCollabs.AddItem collaborateurs(i).nomComplet
            lstCollabs.List(lstCollabs.ListCount - 1, 1) = ""
        End If
    Next i

    If lstCollabs.ListCount = 0 Then
        MsgBox "Aucun collaborateur avec le projet ""GOOGLE LEADS"" trouve dans Utilisateurs.", vbExclamation
    End If
End Sub

' Active/desactive la zone de saisie libre "Entree speciale" (Vague)
Private Sub optVagueSpec_Click()
    txtVagueSpec.Enabled = True
    txtVagueSpec.SetFocus
End Sub
Private Sub optVague1_Click(): txtVagueSpec.Enabled = False: End Sub
Private Sub optVague2_Click(): txtVagueSpec.Enabled = False: End Sub
Private Sub optVague3_Click(): txtVagueSpec.Enabled = False: End Sub
Private Sub optVague4_Click(): txtVagueSpec.Enabled = False: End Sub
Private Sub optVague5_Click(): txtVagueSpec.Enabled = False: End Sub

' Active/desactive la zone de saisie libre "Entree speciale" (Shift Reduit)
Private Sub optRedSpec_Click()
    txtRedSpec.Enabled = True
    txtRedSpec.SetFocus
End Sub
Private Sub optRed1_Click(): txtRedSpec.Enabled = False: End Sub
Private Sub optRed2_Click(): txtRedSpec.Enabled = False: End Sub
Private Sub optRed3_Click(): txtRedSpec.Enabled = False: End Sub
Private Sub optRed4_Click(): txtRedSpec.Enabled = False: End Sub
Private Sub optRed5_Click(): txtRedSpec.Enabled = False: End Sub

' Lit les options actuellement choisies dans le formulaire.
' Retourne False si une saisie speciale est invalide.
Private Function LireConditionCourante(ByRef EntreeVague As String, ByRef SortieVague As String, _
        ByRef EntreeReduit As String, ByRef SortieReduit As String, ByRef ModeDimanche As String) As Boolean

    LireConditionCourante = False

    If optVagueSpec.Value Then
        If Not ParsePlageHoraire(txtVagueSpec.Text, EntreeVague, SortieVague) Then
            MsgBox "Entree speciale (Vague) invalide. Format attendu : HH:MM-HH:MM", vbExclamation
            txtVagueSpec.SetFocus: Exit Function
        End If
    ElseIf optVague1.Value Then: EntreeVague = "07:00": SortieVague = "17:00"
    ElseIf optVague2.Value Then: EntreeVague = "08:00": SortieVague = "18:00"
    ElseIf optVague3.Value Then: EntreeVague = "09:00": SortieVague = "19:00"
    ElseIf optVague4.Value Then: EntreeVague = "10:00": SortieVague = "20:00"
    ElseIf optVague5.Value Then: EntreeVague = "11:00": SortieVague = "20:00"
    Else
        MsgBox "Choisissez une Vague.", vbExclamation: Exit Function
    End If

    If optRedSpec.Value Then
        If Not ParsePlageHoraire(txtRedSpec.Text, EntreeReduit, SortieReduit) Then
            MsgBox "Entree speciale (Shift Reduit) invalide. Format attendu : HH:MM-HH:MM", vbExclamation
            txtRedSpec.SetFocus: Exit Function
        End If
    ElseIf optRed1.Value Then: EntreeReduit = "11:00": SortieReduit = "20:00"
    ElseIf optRed2.Value Then: EntreeReduit = "07:00": SortieReduit = "16:00"
    ElseIf optRed3.Value Then: EntreeReduit = "08:00": SortieReduit = "17:00"
    ElseIf optRed4.Value Then: EntreeReduit = "09:00": SortieReduit = "18:00"
    ElseIf optRed5.Value Then: EntreeReduit = "10:00": SortieReduit = "19:00"
    Else
        MsgBox "Choisissez un Shift Reduit.", vbExclamation: Exit Function
    End If

    If optTT.Value Then
        ModeDimanche = "TT"
    ElseIf optOFF.Value Then
        ModeDimanche = "OFF"
    Else
        ModeDimanche = "TRAVAIL"
    End If

    LireConditionCourante = True
End Function

' ------------------------------------------------------------
' BOUTON PLANIFIER : applique la condition courante a tous les
' collaborateurs selectionnes dans lstCollabs (multi-selection).
' ------------------------------------------------------------
Private Sub cmdPlanifier_Click()
    Dim EntreeVague As String, SortieVague As String
    Dim EntreeReduit As String, SortieReduit As String
    Dim ModeDimanche As String
    If Not LireConditionCourante(EntreeVague, SortieVague, EntreeReduit, SortieReduit, ModeDimanche) Then Exit Sub

    Dim nbSelect As Integer: nbSelect = 0
    Dim r As Integer
    For r = 0 To lstCollabs.ListCount - 1
        If lstCollabs.Selected(r) Then
            nbSelect = nbSelect + 1
            Dim nomComplet As String: nomComplet = lstCollabs.List(r, 0)

            ' Chercher si ce collaborateur a deja une affectation en cours -> on la remplace
            Dim k As Integer, trouve As Integer: trouve = 0
            For k = 1 To m_nbAffect
                If UCase(m_affectations(k).nomComplet) = UCase(nomComplet) Then trouve = k: Exit For
            Next k
            If trouve = 0 Then
                m_nbAffect = m_nbAffect + 1
                ReDim Preserve m_affectations(1 To m_nbAffect)
                trouve = m_nbAffect
            End If

            With m_affectations(trouve)
                .nomComplet = nomComplet
                .EntreeVague = EntreeVague
                .SortieVague = SortieVague
                .EntreeReduit = EntreeReduit
                .SortieReduit = SortieReduit
                .ModeDimanche = ModeDimanche
            End With

            ' Feedback visuel dans la 2e colonne de la liste
            Dim libelleDim As String
            Select Case ModeDimanche
                Case "TT": libelleDim = "Dim TT " & EntreeReduit & "-" & SortieReduit
                Case "OFF": libelleDim = "Dim OFF"
                Case Else: libelleDim = "Dim " & EntreeReduit & "-" & SortieReduit
            End Select
            lstCollabs.List(r, 1) = EntreeVague & "-" & SortieVague & "  |  " & libelleDim
        End If
    Next r

    If nbSelect = 0 Then
        MsgBox "Selectionnez au moins un collaborateur dans la liste (Ctrl+clic pour plusieurs).", vbExclamation
        Exit Sub
    End If

    MsgBox nbSelect & " collaborateur(s) planifie(s) avec cette condition.", vbInformation
End Sub

' ------------------------------------------------------------
' BOUTON GENERER : ecrit toutes les affectations planifiees dans
' la feuille GOOGLE LEADS (+ PLANNING / CONSOLIDATION).
' ------------------------------------------------------------
Private Sub cmdGenerer_Click()
    If m_nbAffect = 0 Then
        MsgBox "Aucun collaborateur planifie. Selectionnez des collaborateurs et cliquez sur ""Planifier"" d'abord.", vbExclamation
        Exit Sub
    End If

    If cboSemaine.ListIndex < 0 Then
        MsgBox "Choisissez une semaine.", vbExclamation: Exit Sub
    End If
    g_LundiCible = m_semaines(cboSemaine.ListIndex)
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(g_LundiCible, 2)

    Dim rep As Integer
    rep = MsgBox("Generer le planning GOOGLE LEADS pour " & m_nbAffect & " collaborateur(s), semaine " & sem & " ?", _
                 vbYesNo + vbQuestion, "Confirmation")
    If rep = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)
    If nbCollab = 0 Then
        MsgBox "Aucun collaborateur trouve dans la feuille Utilisateurs.", vbExclamation
        GoTo Cleanup
    End If

    If Not FeuilleExiste("CONSOLIDATION") Then InitialiserFeuilleConsolidation
    If Not FeuilleExiste("PLANNING") Then InitialiserFeuillePlanning

    GenererPlanningGL_Multi collaborateurs, nbCollab, m_affectations, m_nbAffect

    MsgBox "Planning GOOGLE LEADS genere pour la semaine " & sem & " !", vbInformation, "Termine"

    ' Reinitialiser le plan en cours (feuille ecrite)
    m_nbAffect = 0
    ReDim m_affectations(1 To 1)
    ChargerListeCollabs

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

