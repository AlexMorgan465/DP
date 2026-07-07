VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ProjectRule.cls
' Data-driven description of a project's weekly schedule.
' This REPLACES hardcoded logic like:
'
'   If Project = "DAC" Then ... ElseIf Project = "TLV" Then ...
'
' with a single object the SchedulingEngine can read.
'
' Two things a project can have:
'   1) A fixed shift per day of the week (Jour 1..7)
'   2) An optional rotating "wave" pause time (like GLF), where
'      each employee's pause slot shifts week to week based on
'      their IndexRotation.
' ============================================================

Public NomProjet             As String   ' e.g. "AFEDIM"
Public NomFeuille            As String   ' Excel sheet name (can differ, e.g. "CM Leasing")
Public UtiliseRotationVagues As Boolean
Public TailleGroupeVague     As Integer
Public FooterNote            As String   ' footnote printed when NOT using wave rotation

Private m_Jours(1 To 7)  As ShiftDay
Private m_VaguesPause()  As String
Private m_NbVagues       As Integer

Private Sub Class_Initialize()
    Dim j As Integer
    For j = 1 To 7
        Set m_Jours(j) = New ShiftDay
    Next j
    TailleGroupeVague = 1
    UtiliseRotationVagues = False
    m_NbVagues = 0
End Sub

' j = 1 (Lundi) .. 7 (Dimanche), matching the existing JOURS() convention
Public Property Get Jour(j As Integer) As ShiftDay
    Set Jour = m_Jours(j)
End Property

Public Sub DefinirVagues(vagues() As String)
    Dim n As Integer: n = UBound(vagues) - LBound(vagues) + 1
    ReDim m_VaguesPause(1 To n)
    Dim i As Integer
    For i = 1 To n
        m_VaguesPause(i) = vagues(LBound(vagues) + i - 1)
    Next i
    m_NbVagues = n
End Sub

Public Function NombreVagues() As Integer
    NombreVagues = m_NbVagues
End Function

' index is 1-based
Public Function PauseVague(index As Integer) As String
    PauseVague = m_VaguesPause(index)
End Function
