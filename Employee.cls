VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Employee.cls
' Wraps the legacy "Collaborateur" Type with behavior.
' Conversion helpers keep this compatible with all existing
' code that still passes around Collaborateur() arrays.
' ============================================================

Public NomComplet      As String
Public Nom             As String
Public Prenom          As String
Public Matricule       As String
Public Projet          As String
Public Ville           As String
Public Zone            As String
Public PointRepere     As String
Public Telephone       As String
Public DateEmbauche    As String
Public IndexRotation   As Integer
Public EnConge         As Boolean
Public CongeDebut      As Date
Public CongeFin        As Date
Public EnTT            As Boolean
Public TTDebut         As Date
Public TTFin           As Date
Public RenforcPress    As Boolean
Public RenforcItaly    As Boolean

' --- Conversion: legacy Collaborateur -> Employee ---
Public Sub LoadFromCollaborateur(c As Collaborateur)
    NomComplet = c.nomComplet
    Nom = c.nom
    Prenom = c.Prenom
    Matricule = c.Matricule
    Projet = c.projet
    Ville = c.ville
    Zone = c.zone
    PointRepere = c.PointRepere
    Telephone = c.Telephone
    DateEmbauche = c.DateEmbauche
    IndexRotation = c.IndexRotation
    EnConge = c.EnConge
    CongeDebut = c.CongeDebut
    CongeFin = c.CongeFin
    EnTT = c.EnTT
    TTDebut = c.TTDebut
    TTFin = c.TTFin
    RenforcPress = c.RenforcPress
    RenforcItaly = c.RenforcItaly
End Sub

' --- Conversion: Employee -> legacy Collaborateur ---
' Needed because most of BOOM.bas (writers, exporters, rotation)
' still expects a Collaborateur array. This lets new code use
' Employee objects while old code keeps working untouched.
Public Function ToCollaborateur() As Collaborateur
    Dim c As Collaborateur
    c.nomComplet = NomComplet
    c.nom = Nom
    c.Prenom = Prenom
    c.Matricule = Matricule
    c.projet = Projet
    c.ville = Ville
    c.zone = Zone
    c.PointRepere = PointRepere
    c.Telephone = Telephone
    c.DateEmbauche = DateEmbauche
    c.IndexRotation = IndexRotation
    c.EnConge = EnConge
    c.CongeDebut = CongeDebut
    c.CongeFin = CongeFin
    c.EnTT = EnTT
    c.TTDebut = TTDebut
    c.TTFin = TTFin
    c.RenforcPress = RenforcPress
    c.RenforcItaly = RenforcItaly
    ToCollaborateur = c
End Function

' --- Behavior (this is the point of having a class at all) ---

Public Function EstEnCongeLe(d As Date) As Boolean
    EstEnCongeLe = EnConge And d >= CongeDebut And d <= CongeFin
End Function

Public Function EstEnTTLe(d As Date) As Boolean
    EstEnTTLe = EnTT And d >= TTDebut And d <= TTFin
End Function

Public Function EstDisponibleLe(d As Date) As Boolean
    EstDisponibleLe = Not EstEnCongeLe(d) And Not EstEnTTLe(d)
End Function

Public Function AppartientAuProjet(nomProjet As String) As Boolean
    AppartientAuProjet = (UCase(Trim(Projet)) = UCase(Trim(nomProjet)))
End Function
