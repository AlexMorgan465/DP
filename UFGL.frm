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
' UFGL - Saisie manuelle du planning GOOGLE LEADS (par jour)
' ============================================================
' CE FICHIER EST DU CODE, PAS UN .frm COMPLET.
' Anthropic/Claude ne peut pas generer le binaire .frx d'un
' UserForm (positions/polices des controles). Collez ce code
' dans le module de code du formulaire deja cree dans l'editeur VBA.
'
' *** CONTROLES A AJOUTER / VERIFIER DANS LE DESIGNER ***
'   - cboJour (ComboBox) : deja ajoute par vous. Rempli en code avec
'     Lundi, Mardi, Mercredi, Jeudi, Vendredi, Samedi, Dimanche,
'     "Toutes les semaines".
'   - optChoixVague (OptionButton) et optChoixReduit (OptionButton) :
'     A AJOUTER, places n'importe ou sur le formulaire (par ex. juste
'     avant les libelles "Vague" et "Shift Reduit"). Dans la fenetre
'     Proprietes de CHACUN des deux, mettre la propriete
'         GroupName = grpHoraire
'     (la MEME valeur pour les deux). Cela les rend mutuellement
'     exclusifs entre eux, mais INDEPENDANTS de optTT/optOFF, meme
'     s'ils sont poses directement sur le formulaire.
'     Legende suggeree : optChoixVague = "Vague", optChoixReduit = "Shift Reduit".
'   - optTT / optOFF : deja presents, inchanges. Ils forment un groupe
'     a 2 (deja mutuellement exclusifs car dans le meme conteneur).
'     Aucun des deux coche = jour "travail normal" (ni TT ni OFF).
'
' *** LOGIQUE ***
'   Pour un clic sur "Planifier", la condition appliquee au(x) jour(s)
'   choisi(s) dans cboJour est determinee par 2 choix INDEPENDANTS :
'     1) OFF / TT / (aucun = normal)          -> via optOFF / optTT
'     2) Vague / Shift Reduit (quel horaire)  -> via optChoixVague / optChoixReduit
'   Exemple : optChoixReduit + optTT coches + cboJour="Dimanche"
'             => Dimanche sera "TT" avec les heures du Shift Reduit choisi.
'   Si optOFF est coche, le choix Vague/Reduit est ignore (jour = OFF).
'
'   "Toutes les semaines" dans cboJour applique la MEME condition aux
'   7 jours (Lundi a Dimanche) en un seul clic sur "Planifier".
'
' WORKFLOW :
'   1) Choisir un jour (ou "Toutes les semaines"), une condition
'      (OFF / TT / normal) et un horaire (Vague ou Shift Reduit + le
'      shift precis), selectionner un ou plusieurs collaborateurs dans
'      lstCollabs, cliquer "Planifier".
'   2) Recommencer autant de fois que necessaire (autre jour, autre
'      condition, autres collaborateurs...). Chaque jour deja planifie
'      pour un collaborateur est ecrase si on le replanifie.
'   3) "Generer" ecrit dans la feuille GOOGLE LEADS toutes les
'      affectations planifiees d'un coup. Les jours jamais touches
'      pour un collaborateur sont traites comme OFF (avec confirmation).
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

    ' --- Combo Jour : Lundi..Dimanche + Toutes les semaines ---
    cboJour.Clear
    cboJour.AddItem "Lundi"
    cboJour.AddItem "Mardi"
    cboJour.AddItem "Mercredi"
    cboJour.AddItem "Jeudi"
    cboJour.AddItem "Vendredi"
    cboJour.AddItem "Samedi"
    cboJour.AddItem "Dimanche"
    cboJour.AddItem "Toutes les semaines"
    cboJour.ListIndex = 0   ' par defaut : Lundi

    ' --- Vague (shift), par defaut 7h-17h ---
    optVague1.Value = True
    txtVagueSpec.Text = ""
    txtVagueSpec.Enabled = False

    ' --- Shift Reduit, par defaut 11h-20h ---
    optRed1.Value = True
    txtRedSpec.Text = ""
    txtRedSpec.Enabled = False

    ' --- Quel horaire utiliser pour la condition en cours : Vague par defaut ---
    optChoixVague.Value = True

    ' --- OFF / TT : aucun coche par defaut (= jour travaille normal) ---
    optTT.Value = False
    optOFF.Value = False

    ' --- Liste des collaborateurs GOOGLE LEADS (depuis Utilisateurs) ---
    ChargerListeCollabs

    m_nbAffect = 0
    ReDim m_affectations(1 To 1)
End Sub

' Remplit lstCollabs avec les collaborateurs dont le projet = GOOGLE LEADS.
' Colonne 1 = Nom complet, Colonne 2 = resume des jours deja planifies.
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

' ------------------------------------------------------------
' Lit l'horaire choisi dans la section VAGUE.
' ------------------------------------------------------------
Private Function LireHoraireVague(ByRef entree As String, ByRef sortie As String) As Boolean
    LireHoraireVague = False
    If optVagueSpec.Value Then
        If Not ParsePlageHoraire(txtVagueSpec.Text, entree, sortie) Then
            MsgBox "Entree speciale (Vague) invalide. Format attendu : HH:MM-HH:MM", vbExclamation
            txtVagueSpec.SetFocus: Exit Function
        End If
    ElseIf optVague1.Value Then: entree = "07:00": sortie = "17:00"
    ElseIf optVague2.Value Then: entree = "08:00": sortie = "18:00"
    ElseIf optVague3.Value Then: entree = "09:00": sortie = "19:00"
    ElseIf optVague4.Value Then: entree = "10:00": sortie = "20:00"
    ElseIf optVague5.Value Then: entree = "11:00": sortie = "20:00"
    Else
        MsgBox "Choisissez une Vague.", vbExclamation: Exit Function
    End If
    LireHoraireVague = True
End Function

' ------------------------------------------------------------
' Lit l'horaire choisi dans la section SHIFT REDUIT.
' ------------------------------------------------------------
Private Function LireHoraireReduit(ByRef entree As String, ByRef sortie As String) As Boolean
    LireHoraireReduit = False
    If optRedSpec.Value Then
        If Not ParsePlageHoraire(txtRedSpec.Text, entree, sortie) Then
            MsgBox "Entree speciale (Shift Reduit) invalide. Format attendu : HH:MM-HH:MM", vbExclamation
            txtRedSpec.SetFocus: Exit Function
        End If
    ElseIf optRed1.Value Then: entree = "11:00": sortie = "20:00"
    ElseIf optRed2.Value Then: entree = "07:00": sortie = "16:00"
    ElseIf optRed3.Value Then: entree = "08:00": sortie = "17:00"
    ElseIf optRed4.Value Then: entree = "09:00": sortie = "18:00"
    ElseIf optRed5.Value Then: entree = "10:00": sortie = "19:00"
    Else
        MsgBox "Choisissez un Shift Reduit.", vbExclamation: Exit Function
    End If
    LireHoraireReduit = True
End Function

' ------------------------------------------------------------
' Determine, a partir des options cochees, le MODE (TRAVAIL/TT/OFF)
' et l'horaire (Entree/Sortie) a appliquer au(x) jour(s) choisi(s).
' Retourne False si la saisie est invalide/incomplete (message deja affiche).
' ------------------------------------------------------------
Private Function LireModeEtHoraire(ByRef mode As String, ByRef entree As String, ByRef sortie As String) As Boolean
    LireModeEtHoraire = False

    If optOFF.Value Then
        mode = "OFF": entree = "": sortie = ""
        LireModeEtHoraire = True
        Exit Function
    End If

    Dim horaireOK As Boolean
    If optChoixReduit.Value Then
        horaireOK = LireHoraireReduit(entree, sortie)
    Else
        horaireOK = LireHoraireVague(entree, sortie)
    End If
    If Not horaireOK Then Exit Function

    If optTT.Value Then
        mode = "TT"
    Else
        mode = "TRAVAIL"
    End If
    LireModeEtHoraire = True
End Function

' ------------------------------------------------------------
' Lit cboJour et retourne la liste des jours cibles (1=Lundi..7=Dimanche).
' "Toutes les semaines" -> les 7 jours.
' ------------------------------------------------------------
Private Function ListeJoursCibles(ByRef joursCibles() As Integer, ByRef nbJours As Integer) As Boolean
    ListeJoursCibles = False
    If cboJour.ListIndex < 0 Then
        MsgBox "Choisissez un jour (ou ""Toutes les semaines"").", vbExclamation
        Exit Function
    End If

    If cboJour.ListIndex = 7 Then   ' "Toutes les semaines"
        nbJours = 7
        ReDim joursCibles(1 To 7)
        Dim j As Integer
        For j = 1 To 7: joursCibles(j) = j: Next j
    Else
        nbJours = 1
        ReDim joursCibles(1 To 1)
        joursCibles(1) = cboJour.ListIndex + 1   ' 0-based -> 1=Lundi..7=Dimanche
    End If
    ListeJoursCibles = True
End Function

' Libelle court des jours cibles pour les messages utilisateur.
Private Function LibelleJoursCibles(joursCibles() As Integer, nbJours As Integer) As String
    Dim nomsJours As Variant
    nomsJours = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche")
    If nbJours = 7 Then
        LibelleJoursCibles = "toute la semaine"
        Exit Function
    End If
    LibelleJoursCibles = nomsJours(joursCibles(1) - 1)
End Function

' Cherche l'affectation existante d'un collaborateur, ou en cree une nouvelle.
Private Function TrouverOuCreerAffectation(nomComplet As String) As Integer
    Dim k As Integer
    For k = 1 To m_nbAffect
        If UCase(m_affectations(k).nomComplet) = UCase(nomComplet) Then
            TrouverOuCreerAffectation = k
            Exit Function
        End If
    Next k
    m_nbAffect = m_nbAffect + 1
    ReDim Preserve m_affectations(1 To m_nbAffect)
    m_affectations(m_nbAffect).nomComplet = nomComplet
    TrouverOuCreerAffectation = m_nbAffect
End Function

' Construit le resume affiche dans la 2e colonne de lstCollabs
' (ex : "L:07-17 Ma:07-17 Me:OFF Je:07-17 Ve:07-17 Sa:07-17 Di:TT 11-20").
Private Function ResumeAffectation(aff As AffectationGL) As String
    Dim initiales As Variant
    initiales = Array("L", "Ma", "Me", "Je", "Ve", "Sa", "Di")
    Dim s As String: s = ""
    Dim j As Integer
    For j = 1 To 7
        Dim morceau As String
        With aff.jours(j)
            If Not .Defini Then
                morceau = "-"
            ElseIf UCase(.Mode) = "OFF" Then
                morceau = "OFF"
            ElseIf UCase(.Mode) = "TT" Then
                morceau = "TT " & .Entree & "-" & .Sortie
            Else
                morceau = .Entree & "-" & .Sortie
            End If
        End With
        s = s & initiales(j - 1) & ":" & morceau & "  "
    Next j
    ResumeAffectation = Trim(s)
End Function

' ------------------------------------------------------------
' BOUTON PLANIFIER : applique la condition courante (mode + horaire)
' au(x) jour(s) choisi(s) dans cboJour, pour tous les collaborateurs
' selectionnes dans lstCollabs (multi-selection).
' ------------------------------------------------------------
Private Sub cmdPlanifier_Click()
    Dim mode As String, entree As String, sortie As String
    If Not LireModeEtHoraire(mode, entree, sortie) Then Exit Sub

    Dim joursCibles() As Integer
    Dim nbJours As Integer
    If Not ListeJoursCibles(joursCibles, nbJours) Then Exit Sub

    Dim nbSelect As Integer: nbSelect = 0
    Dim r As Integer
    For r = 0 To lstCollabs.ListCount - 1
        If lstCollabs.Selected(r) Then
            nbSelect = nbSelect + 1
            Dim nomComplet As String: nomComplet = lstCollabs.List(r, 0)
            Dim idxAff As Integer: idxAff = TrouverOuCreerAffectation(nomComplet)

            Dim jj As Integer
            For jj = 1 To nbJours
                Dim jCible As Integer: jCible = joursCibles(jj)
                With m_affectations(idxAff).jours(jCible)
                    .Defini = True
                    .Mode = mode
                    .Entree = entree
                    .Sortie = sortie
                End With
            Next jj

            lstCollabs.List(r, 1) = ResumeAffectation(m_affectations(idxAff))
        End If
    Next r

    If nbSelect = 0 Then
        MsgBox "Selectionnez au moins un collaborateur dans la liste (Ctrl+clic pour plusieurs).", vbExclamation
        Exit Sub
    End If

    MsgBox nbSelect & " collaborateur(s) planifie(s) pour " & LibelleJoursCibles(joursCibles, nbJours) & ".", vbInformation
End Sub

' ------------------------------------------------------------
' BOUTON GENERER : ecrit toutes les affectations planifiees dans
' la feuille GOOGLE LEADS (+ PLANNING / CONSOLIDATION). Les jours
' jamais planifies pour un collaborateur sont traites comme OFF
' (apres confirmation).
' ------------------------------------------------------------
Private Sub cmdGenerer_Click()
    If m_nbAffect = 0 Then
        MsgBox "Aucun collaborateur planifie. Selectionnez des collaborateurs, choisissez un jour et une condition, puis cliquez sur ""Planifier"" d'abord.", vbExclamation
        Exit Sub
    End If

    If cboSemaine.ListIndex < 0 Then
        MsgBox "Choisissez une semaine.", vbExclamation: Exit Sub
    End If
    g_LundiCible = m_semaines(cboSemaine.ListIndex)
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(g_LundiCible, 2)

    ' Verifier les jours non planifies (seront traites comme OFF)
    Dim messageManquants As String: messageManquants = ""
    Dim k As Integer, j As Integer
    For k = 1 To m_nbAffect
        Dim joursManquants As String: joursManquants = ""
        For j = 1 To 7
            If Not m_affectations(k).jours(j).Defini Then
                Dim nomsJours As Variant
                nomsJours = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
                joursManquants = joursManquants & nomsJours(j - 1) & " "
            End If
        Next j
        If joursManquants <> "" Then
            messageManquants = messageManquants & "- " & m_affectations(k).nomComplet & " : " & Trim(joursManquants) & Chr(10)
        End If
    Next k

    If messageManquants <> "" Then
        Dim repManquants As Integer
        repManquants = MsgBox("Certains jours n'ont pas ete planifies et seront mis en OFF par defaut :" & Chr(10) & Chr(10) & _
                               messageManquants & Chr(10) & "Continuer ?", vbYesNo + vbExclamation, "Jours non planifies")
        If repManquants = vbNo Then Exit Sub
    End If

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
