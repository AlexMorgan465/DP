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
'   - lstJours (ListBox) : REMPLACE cboJour. Rempli en code avec
'     Lundi, Mardi, Mercredi, Jeudi, Vendredi, Samedi, Dimanche,
'     "Toutes les semaines" (8 lignes, index 0 a 7).
'     Proprietes a regler dans le Designer : MultiSelect = fmMultiSelectMulti
'     (2 - fmMultiSelectMulti), pour pouvoir cocher plusieurs jours d'un
'     coup (ex: Lundi + Mercredi + Vendredi) sans passer par "Toutes les
'     semaines". Si "Toutes les semaines" est coche, les autres jours
'     coches sont ignores (voir ListeJoursCibles plus bas).
'   - txtRecherche (TextBox) : NOUVEAU, filtre lstCollabs en direct pendant
'     la saisie (recherche par sous-chaine dans le nom, insensible a la
'     casse). Voir txtRecherche_Change / FiltrerListeCollabs plus bas.
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
'   coche(s) dans lstJours est determinee par 2 choix INDEPENDANTS :
'     1) OFF / TT / (aucun = normal)          -> via optOFF / optTT
'     2) Vague / Shift Reduit (quel horaire)  -> via optChoixVague / optChoixReduit
'   Exemple : optChoixReduit + optTT coches + lstJours="Dimanche"
'             => Dimanche sera "TT" avec les heures du Shift Reduit choisi.
'   Si optOFF est coche, le choix Vague/Reduit est ignore (jour = OFF).
'
'   lstJours est multi-selection : on peut cocher plusieurs jours a la
'   fois (ex: Lundi + Mercredi + Vendredi) pour leur appliquer la meme
'   condition en un seul clic sur "Planifier". Cocher "Toutes les
'   semaines" applique la condition aux 7 jours (Lundi a Dimanche) et
'   ignore les autres cases cochees dans lstJours.
'
' *** CORRECTIONS DE CETTE VERSION ***
'   - Au demarrage, AUCUN jour n'est preselectionne dans lstJours (avant :
'     "Lundi" etait coche par defaut). L'utilisateur doit cocher un ou
'     plusieurs jours, ou "Toutes les semaines", explicitement avant de
'     planifier.
'   - OFF et TT sont totalement independants du/des jour(s) choisi(s) :
'     cocher OFF ou TT n'applique JAMAIS "toute la semaine" tout seul ;
'     c'est uniquement la selection dans lstJours qui determine le(s)
'     jour(s) cible(s). Handlers optOFF_Click/optTT_Click ajoutes pour
'     le garantir.
'   - Rappel : pour pouvoir cocher Vague/Reduit EN MEME TEMPS que OFF ou
'     TT (et meme avec "Toutes les semaines"), optChoixVague et
'     optChoixReduit doivent avoir GroupName = grpHoraire dans le
'     Designer (voir plus haut). Sans ce reglage, VBA les rend
'     mutuellement exclusifs avec optTT/optOFF.
'
' *** MISE A JOUR : optTT / optChoixVague / optChoixReduit en ToggleButton ***
'   Ces 3 controles ont ete changes d'OptionButton vers ToggleButton pour ne
'   plus etre lies automatiquement a un groupe d'exclusion mutuelle. Un
'   ToggleButton n'impose PLUS l'exclusivite tout seul : c'est desormais le
'   code (optOFF_Click / optTT_Click / optChoixVague_Click / optChoixReduit_Click,
'   plus bas) qui force manuellement :
'     - optOFF et optTT restent mutuellement exclusifs entre eux (un jour ne
'       peut pas etre a la fois OFF et TT), mais ne touchent jamais lstJours.
'     - optChoixVague et optChoixReduit restent mutuellement exclusifs entre
'       eux (un seul horaire actif a la fois), mais restent independants de
'       optOFF/optTT.
'   Hypothese : optOFF reste un OptionButton classique et le controle Reduit
'   s'appelle toujours "optChoixReduit". Si vous l'avez renomme differemment
'   (ex: optChoixRed), remplacez le nom dans le code ci-dessous.
'
' *** NOUVEAU : lblTerminal (rapport en direct) ***
'   Un Label nomme lblTerminal a ete ajoute. A chaque clic sur "Planifier",
'   MettreAJourTerminal (plus bas) y affiche : le detail jour par jour de
'   chaque collaborateur deja planifie, puis la liste des collaborateurs de
'   lstCollabs qui n'ont encore AUCUN jour defini.
'   Conseil Designer pour lblTerminal : WordWrap = True, AutoSize = False,
'   redimensionner assez haut/large, police a chasse fixe (ex: Consolas)
'   pour un rendu "terminal" lisible.
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
Private m_tousCollabsNoms() As String   ' liste complete (non filtree) des noms GOOGLE LEADS
Private m_nbTousCollabs As Integer

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

    ' --- lstJours : Lundi..Dimanche + Toutes les semaines (multi-selection) ---
    lstJours.Clear
    lstJours.AddItem "Lundi"
    lstJours.AddItem "Mardi"
    lstJours.AddItem "Mercredi"
    lstJours.AddItem "Jeudi"
    lstJours.AddItem "Vendredi"
    lstJours.AddItem "Samedi"
    lstJours.AddItem "Dimanche"
    lstJours.AddItem "Toutes les semaines"
    ' Rappel Designer : lstJours.MultiSelect doit etre regle sur
    ' 2 - fmMultiSelectMulti pour pouvoir cocher plusieurs jours.
    ' Aucun jour n'est preselectionne par defaut (Selected() = False pour
    ' toutes les lignes tant qu'on n'y touche pas).

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

    MettreAJourTerminal
End Sub

' Remplit lstCollabs avec les collaborateurs dont le projet = GOOGLE LEADS.
' Colonne 1 = Nom complet, Colonne 2 = resume des jours deja planifies.
' Pense a activer, sur lstCollabs : ColumnCount = 2, MultiSelect = fmMultiSelectExtended.
' La liste complete (non filtree) est gardee dans m_tousCollabsNoms ; lstCollabs
' n'affiche que le sous-ensemble filtre par txtRecherche (voir FiltrerListeCollabs).
Private Sub ChargerListeCollabs()
    m_nbTousCollabs = 0
    If Not FeuilleExiste("Utilisateurs") Then
        lstCollabs.Clear
        Exit Sub
    End If

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)

    ReDim m_tousCollabsNoms(1 To IIf(nbCollab > 0, nbCollab, 1))
    Dim i As Integer
    For i = 1 To nbCollab
        If UCase(Trim(collaborateurs(i).projet)) = "GOOGLE LEADS" Then
            m_nbTousCollabs = m_nbTousCollabs + 1
            m_tousCollabsNoms(m_nbTousCollabs) = collaborateurs(i).nomComplet
        End If
    Next i

    txtRecherche.Text = ""
    FiltrerListeCollabs

    If m_nbTousCollabs = 0 Then
        MsgBox "Aucun collaborateur avec le projet ""GOOGLE LEADS"" trouve dans Utilisateurs.", vbExclamation
    End If
End Sub

' ------------------------------------------------------------
' Reconstruit lstCollabs a partir de m_tousCollabsNoms, en ne gardant que
' les noms contenant le texte de txtRecherche (recherche par sous-chaine,
' insensible a la casse). Le resume (colonne 2) de chaque collaborateur
' deja planifie (present dans m_affectations) est reaffiche apres filtrage.
' NOTE : changer le texte de recherche reconstruit la liste, donc les
' cases cochees dans lstCollabs sont reinitialisees a chaque frappe. Ce
' n'est pas un probleme puisque les affectations deja "Planifie" sont
' sauvegardees dans m_affectations, independamment des cases cochees.
' ------------------------------------------------------------
Private Sub FiltrerListeCollabs()
    Dim filtre As String
    filtre = UCase(Trim(txtRecherche.Text))

    lstCollabs.Clear
    Dim i As Integer, k As Integer
    For i = 1 To m_nbTousCollabs
        Dim nom As String: nom = m_tousCollabsNoms(i)
        If filtre = "" Or InStr(1, UCase(nom), filtre) > 0 Then
            lstCollabs.AddItem nom
            Dim ligne As Integer: ligne = lstCollabs.ListCount - 1
            lstCollabs.List(ligne, 1) = ""
            For k = 1 To m_nbAffect
                If UCase(m_affectations(k).nomComplet) = UCase(nom) Then
                    lstCollabs.List(ligne, 1) = ResumeAffectation(m_affectations(k))
                    Exit For
                End If
            Next k
        End If
    Next i
End Sub

' Filtre lstCollabs a chaque frappe dans la zone de recherche.
Private Sub txtRecherche_Change()
    FiltrerListeCollabs
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
' optOFF / optTT : ces 2 handlers existent uniquement pour garantir
' explicitement qu'AUCUN code ne touche lstJours ni optChoixVague /
' optChoixReduit quand on coche OFF ou TT. Le(s) jour(s) applique(s)
' reste(nt) TOUJOURS exactement ceux coches dans lstJours : cocher OFF
' ou TT ne force JAMAIS "toute la semaine" tout seul.
' De meme, OFF/TT (mode) et Vague/Reduit (horaire) sont 2 groupes
' INDEPENDANTS : cocher OFF ou TT n'empeche pas de cocher Vague ou
' Shift Reduit egalement (utile si vous decochez OFF/TT ensuite).
' IMPORTANT : pour que Vague/Reduit restent cochables EN MEME TEMPS
' que OFF/TT et que "Toutes les semaines", verifiez bien dans le
' Designer que optChoixVague et optChoixReduit ont la MEME propriete
' GroupName = grpHoraire (voir en-tete du fichier). Sans cela, VBA
' les considere comme faisant partie du groupe par defaut du
' formulaire et les rend mutuellement exclusifs avec optTT/optOFF.
' ------------------------------------------------------------
Private Sub optOFF_Click()
    ' optOFF est reste OptionButton : s'il vient d'etre coche, on decoche
    ' manuellement optTT (ToggleButton) puisque celui-ci n'est plus dans
    ' le meme groupe d'exclusion automatique. Ne touche ni lstJours, ni
    ' optChoixVague/optChoixReduit.
    If optOFF.Value Then optTT.Value = False
End Sub
Private Sub optTT_Click()
    ' optTT est un ToggleButton : on force manuellement l'exclusivite avec
    ' optOFF. Ne touche ni lstJours, ni optChoixVague/optChoixReduit.
    If optTT.Value Then optOFF.Value = False
End Sub

' optChoixVague / optChoixReduit sont maintenant des ToggleButton : ils ne
' s'excluent plus automatiquement (l'ancien GroupName=grpHoraire n'a plus
' d'effet sur un ToggleButton). On force ici manuellement qu'un seul des
' deux soit actif, sans jamais toucher optOFF/optTT ni lstJours.
Private Sub optChoixVague_Click()
    If optChoixVague.Value Then optChoixReduit.Value = False
End Sub
Private Sub optChoixReduit_Click()
    If optChoixReduit.Value Then optChoixVague.Value = False
End Sub

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
' Lit lstJours (multi-selection) et retourne la liste des jours cibles
' (1=Lundi..7=Dimanche). Si "Toutes les semaines" (index 7) est coche,
' les 7 jours sont retournes et les autres cases cochees sont ignorees.
' ------------------------------------------------------------
Private Function ListeJoursCibles(ByRef joursCibles() As Integer, ByRef nbJours As Integer) As Boolean
    ListeJoursCibles = False

    If lstJours.Selected(7) Then   ' "Toutes les semaines"
        nbJours = 7
        ReDim joursCibles(1 To 7)
        Dim j As Integer
        For j = 1 To 7: joursCibles(j) = j: Next j
        ListeJoursCibles = True
        Exit Function
    End If

    Dim i As Integer, n As Integer
    Dim indices(1 To 7) As Integer
    n = 0
    For i = 0 To 6   ' Lundi(0)..Dimanche(6)
        If lstJours.Selected(i) Then
            n = n + 1
            indices(n) = i + 1
        End If
    Next i

    If n = 0 Then
        MsgBox "Cochez au moins un jour (ou ""Toutes les semaines"").", vbExclamation
        Exit Function
    End If

    nbJours = n
    ReDim joursCibles(1 To n)
    For i = 1 To n: joursCibles(i) = indices(i): Next i
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
    Dim s As String, i As Integer
    For i = 1 To nbJours
        s = s & nomsJours(joursCibles(i) - 1)
        If i < nbJours Then s = s & ", "
    Next i
    LibelleJoursCibles = s
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
' Construit et affiche dans lblTerminal l'etat du planning en cours :
'   - pour chaque collaborateur deja planifie, le detail jour par jour
'     (reutilise ResumeAffectation : mode + horaire par jour)
'   - la liste des collaborateurs de lstCollabs n'ayant ENCORE AUCUN
'     jour defini (non planifies)
' Appelee a chaque clic sur "Planifier".
' ------------------------------------------------------------
Private Sub MettreAJourTerminal()
    Dim s As String
    s = "=== PLANNING EN COURS (" & m_nbAffect & " collaborateur(s)) ===" & vbCrLf & vbCrLf

    Dim k As Integer
    If m_nbAffect = 0 Then
        s = s & "(aucun collaborateur planifie pour le moment)" & vbCrLf
    Else
        For k = 1 To m_nbAffect
            s = s & m_affectations(k).nomComplet & " :" & vbCrLf
            s = s & "    " & ResumeAffectation(m_affectations(k)) & vbCrLf
        Next k
    End If

    ' --- Collaborateurs de la liste sans aucun jour planifie ---
    Dim r As Integer, j As Integer
    Dim nomComplet As String
    Dim nonPlanifies As String: nonPlanifies = ""
    Dim nbNonPlan As Integer: nbNonPlan = 0

    For r = 0 To lstCollabs.ListCount - 1
        nomComplet = lstCollabs.List(r, 0)
        Dim idxAff As Integer: idxAff = -1
        For k = 1 To m_nbAffect
            If UCase(m_affectations(k).nomComplet) = UCase(nomComplet) Then
                idxAff = k
                Exit For
            End If
        Next k

        Dim aAuMoinsUnJour As Boolean: aAuMoinsUnJour = False
        If idxAff > 0 Then
            For j = 1 To 7
                If m_affectations(idxAff).jours(j).Defini Then aAuMoinsUnJour = True: Exit For
            Next j
        End If

        If Not aAuMoinsUnJour Then
            nonPlanifies = nonPlanifies & "  - " & nomComplet & vbCrLf
            nbNonPlan = nbNonPlan + 1
        End If
    Next r

    s = s & vbCrLf & "--- Collaborateurs non planifies (" & nbNonPlan & ") ---" & vbCrLf
    If nbNonPlan = 0 Then
        s = s & "  (aucun, tous ont au moins un jour planifie)" & vbCrLf
    Else
        s = s & nonPlanifies
    End If

    lblTerminal.Caption = s
End Sub

' ------------------------------------------------------------
' BOUTON PLANIFIER : applique la condition courante (mode + horaire)
' au(x) jour(s) choisi(s) dans lstJours, pour tous les collaborateurs
' selectionnes dans lstCollabs (multi-selection).
' ------------------------------------------------------------
Private Sub cmdPlanifier_Click()
    ' --- DEBUG TEMPORAIRE : capture l'etat brut des boutons juste avant
    ' d'appliquer quoi que ce soit, pour diagnostiquer le bug "TT applique
    ' sans avoir coche TT". A retirer une fois le probleme confirme/corrige.
    Dim dbg As String
    dbg = "[DEBUG clic Planifier] OFF=" & optOFF.Value & "  TT=" & optTT.Value & _
          "  Vague=" & optChoixVague.Value & "  Reduit=" & optChoixReduit.Value & vbCrLf
    Dim dbgJours As String: dbgJours = "[DEBUG jours coches] "
    Dim di As Integer
    Dim nomsJoursDbg As Variant
    nomsJoursDbg = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche", "ToutesLesSemaines")
    For di = 0 To 7
        If lstJours.Selected(di) Then dbgJours = dbgJours & nomsJoursDbg(di) & " "
    Next di
    dbg = dbg & dbgJours & vbCrLf & vbCrLf

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

    MettreAJourTerminal
    lblTerminal.Caption = dbg & lblTerminal.Caption

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
