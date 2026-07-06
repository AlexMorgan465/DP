VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFMain 
   Caption         =   "UserForm1"
   ClientHeight    =   10740
   ClientLeft      =   0
   ClientTop       =   96
   ClientWidth     =   7512
   OleObjectBlob   =   "UFMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExporter_Click()
ExporterPlanningsParActivite
End Sub

Private Sub cmdGenererMail_Click()
GenererEmailRenfort
End Sub

Private Sub lblStatut_Click()

End Sub

' ============================================================
' INITIALISATION
' ============================================================
Private Sub UserForm_Initialize()
    ' Afficher infos semaine courante
    Dim lundi As Date
    Dim wd As Integer
    wd = Weekday(Date, vbMonday)
    lundi = Date - (wd - 1)
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(lundi, 2)

    lblSemInfo.Caption = "Semaine courante : S" & sem & "  |  " & _
                         Format(lundi, "dd/mm/yyyy") & " à " & Format(lundi + 6, "dd/mm/yyyy") & _
                         "  |  Aujourd'hui : " & Format(Date, "dddd dd/mm/yyyy")

    ' Compter collaborateurs si feuille existe
    Dim nbCollab As String: nbCollab = ""
    If FeuilleExiste("Utilisateurs") Then
        Dim wsU As Worksheet
        Set wsU = ThisWorkbook.Sheets("Utilisateurs")
        Dim lr As Long
        lr = wsU.Cells(wsU.Rows.Count, 1).End(xlUp).Row
        If lr > 1 Then nbCollab = CStr(lr - 1)
    End If

    ' Compter besoins si feuille existe
    Dim nbBesoins As String: nbBesoins = ""
    If FeuilleExiste("BESOINS") Then
        Dim wsB As Worksheet
        Set wsB = ThisWorkbook.Sheets("BESOINS")
        Dim lrB As Long
        lrB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
        If lrB > 1 Then nbBesoins = CStr(lrB - 1)
    End If

    lblStatut.Caption = "Collaborateurs enregistrés : " & nbCollab & _
                        "   |   Besoins de renfort : " & nbBesoins
End Sub

' ============================================================
' BOUTON : GÃ‰NÃ‰RER LE PLANNING
' ============================================================
Private Sub cmdGenerer_Click()
    Me.Hide
    UFGenerer.Show
    ' Après fermeture de UFGenerer, revenir au lobby
    If Me.Visible = False Then
        ' RafraÃ®chir le statut
        UserForm_Initialize
        Me.Show
    End If
End Sub

' ============================================================
' BOUTON : GESTION UTILISATEURS
' ============================================================
Private Sub cmdUtilisateurs_Click()
    lblStatut.Caption = "Ouverture de la gestion des utilisateurs..."
    Me.Hide
    UFUtilisateurs.Show
    UserForm_Initialize
    Me.Show
End Sub

' ============================================================
' BOUTON : GESTION BESOINS
' ============================================================
Private Sub cmdBesoins_Click()
    lblStatut.Caption = "Ouverture de la gestion des besoins..."
    Me.Hide
    UFBesoins.Show
    UserForm_Initialize
    Me.Show
End Sub

' ============================================================
' BOUTON : RÃ‰INITIALISER ROTATIONS
' ============================================================
Private Sub cmdResetRotations_Click()
    Dim rep As Integer
    rep = MsgBox("Réinitialiser toutes les rotations ?" & Chr(10) & _
                 "Cette action remet tous les index Ã  0.", _
                 vbYesNo + vbWarning, "Confirmation")
    If rep = vbNo Then Exit Sub

    If FeuilleExiste("ROTATION") Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ROTATION")
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
            ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 6)).Clear
        End If
        lblStatut.Caption = "âœ” Rotations réinitialisées avec succès."
        MsgBox "Rotations réinitialisées.", vbInformation
    Else
        lblStatut.Caption = "âš  Feuille ROTATION introuvable."
    End If
End Sub

' ============================================================
' BOUTON : VOIR CONSOLIDATION
' ============================================================
Private Sub cmdVoirConsolidation_Click()
    If FeuilleExiste("CONSOLIDATION") Then
        ThisWorkbook.Sheets("CONSOLIDATION").Activate
        lblStatut.Caption = "Feuille CONSOLIDATION activée."
        Me.Hide
    Else
        MsgBox "La feuille CONSOLIDATION n'existe pas encore." & Chr(10) & _
               "Générez d'abord un planning.", vbExclamation
    End If
End Sub

' ============================================================
' BOUTON : VOIR PLANNING GLOBAL
' ============================================================
Private Sub cmdVoirPlanning_Click()
    If FeuilleExiste("PLANNING") Then
        ThisWorkbook.Sheets("PLANNING").Activate
        lblStatut.Caption = "Feuille PLANNING activée."
        Me.Hide
    Else
        MsgBox "La feuille PLANNING n'existe pas encore." & Chr(10) & _
               "Générez d'abord un planning.", vbExclamation
    End If
End Sub

' ============================================================
' BOUTON : VOIR FEUILLE BESOINS
' ============================================================
Private Sub cmdVoirBesoins_Click()
    If FeuilleExiste("BESOINS") Then
        ThisWorkbook.Sheets("BESOINS").Activate
        lblStatut.Caption = "Feuille BESOINS activée."
        Me.Hide
    Else
        MsgBox "La feuille BESOINS est introuvable.", vbExclamation
    End If
End Sub

' ============================================================
' FERMER
' ============================================================
Private Sub cmdFermer_Click()
    Unload Me
End Sub

' ============================================================
' HELPER LOCAL
' ============================================================
Private Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

' ============================================================
' POINT D'ENTRÃ‰E PUBLIC â€” assigner Ã  un bouton ruban
' ============================================================
Public Sub Ouvrir()
    UFMain.Show
End Sub


