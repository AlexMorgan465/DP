Attribute VB_Name = "ProjectRulesRepository"
Option Explicit

' ============================================================
' ProjectRulesRepository.bas
'
' Single place that knows each project's weekly rules.
' Today the rules are built in code (BuildFixedRule / BuildGLFRule)
' so behavior is IDENTICAL to the current GenererPlanningFixe /
' GenererPlanningGLF subs -- this is a safe, zero-behavior-change
' extraction.
'
' NEXT STEP (once this is verified working): replace the bodies
' of BuildFixedRule/BuildGLFRule with code that reads a
' "Config_Projets" worksheet instead, so adding a new project
' means adding rows in Excel, not new VBA. The public function
' LoadProjectRule() is the only thing the rest of the app calls,
' so that change is invisible to everything else.
' ============================================================

Public Function LoadProjectRule(nomProjet As String) As ProjectRule
    Dim rule As ProjectRule
    Select Case UCase(Trim(nomProjet))
        Case "AFEDIM"
            Set rule = BuildFixedRule("AFEDIM", "AFEDIM")
        Case "ACCESSIBILITE"
            Set rule = BuildFixedRule("ACCESSIBILITE", "ACCESSIBILITE")
        Case "CMLEASING", "CM LEASING"
            Set rule = BuildFixedRule("CMLEASING", "CM Leasing")
        Case "GLF"
            Set rule = BuildGLFRule()
        Case Else
            Set rule = Nothing   ' project not yet migrated to the engine
    End Select
    Set LoadProjectRule = rule
End Function

' Mon-Thu 08:00-18:00, Fri 08:00-17:00, pause 13:00-14:00, weekend off.
' Matches GenererPlanningFixe exactly.
Private Function BuildFixedRule(nomProjet As String, nomFeuille As String) As ProjectRule
    Dim r As New ProjectRule
    r.NomProjet = nomProjet
    r.NomFeuille = nomFeuille
    r.UtiliseRotationVagues = False
    r.FooterNote = "Total : 44h | Pause fixe 13:00-14:00 | TT = fond violet"

    Dim j As Integer
    For j = 1 To 4
        r.Jour(j).DefinirTravail "08:00", "18:00", "13:00", "14:00"
    Next j
    r.Jour(5).DefinirTravail "08:00", "17:00", "13:00", "14:00"
    r.Jour(6).DefinirOff
    r.Jour(7).DefinirOff

    Set BuildFixedRule = r
End Function

' Same weekly hours as the fixed rule, but the pause slot rotates
' across 5 "waves" based on each employee's IndexRotation.
' Matches GenererPlanningGLF exactly.
Private Function BuildGLFRule() As ProjectRule
    Dim r As New ProjectRule
    r.NomProjet = "GLF"
    r.NomFeuille = "GLF"
    r.UtiliseRotationVagues = True
    r.TailleGroupeVague = 5

    Dim vagues(1 To 5) As String
    vagues(1) = "12:00": vagues(2) = "12:30": vagues(3) = "13:00"
    vagues(4) = "13:30": vagues(5) = "14:00"
    r.DefinirVagues vagues

    Dim j As Integer
    For j = 1 To 4
        r.Jour(j).DefinirTravail "08:00", "18:00", "", ""   ' pause supplied by the wave
    Next j
    r.Jour(5).DefinirTravail "08:00", "17:00", "", ""
    r.Jour(6).DefinirOff
    r.Jour(7).DefinirOff

    Set BuildGLFRule = r
End Function
