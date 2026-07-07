VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShiftDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ShiftDay.cls
' Describes ONE day of a project's weekly pattern: either a
' working shift (start/end/pause) or an OFF day.
' ============================================================

Public IsOff        As Boolean
Public HeureDebut   As String
Public HeureFin     As String
Public PauseDebut   As String
Public PauseFin     As String

Private Sub Class_Initialize()
    DefinirOff
End Sub

Public Sub DefinirTravail(hDebut As String, hFin As String, pDebut As String, pFin As String)
    IsOff = False
    HeureDebut = hDebut
    HeureFin = hFin
    PauseDebut = pDebut
    PauseFin = pFin
End Sub

Public Sub DefinirOff()
    IsOff = True
    HeureDebut = ""
    HeureFin = ""
    PauseDebut = ""
    PauseFin = ""
End Sub
