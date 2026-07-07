Attribute VB_Name = "SchedulingEngine"
Option Explicit

' ============================================================
' SchedulingEngine.bas
'
' ONE generic sub that replaces:
'   - GenererPlanningFixe          (AFEDIM / ACCESSIBILITE / CMLEASING)
'   - the hardcoded body of GenererPlanningGLF
'
' It calls the SAME low-level helpers already in BOOM.bas
' (FormatCelluleJour, AjouterMinutes, AppliquerCongesEtTT,
'  EcrireLigneAvecConsolidation, EcrireEnTeteHorizontale,
'  AppliquerBorduresH) so nothing else in the workbook needs to
' change. Output is identical to the current code, row for row.
' ============================================================

Public Sub GenerateProjectPlanning(rule As ProjectRule, collabs() As Collaborateur, nb As Integer)
    If rule Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(rule.NomFeuille)
    EcrireEnTeteHorizontale ws, rule.NomFeuille

    ' --- Filter employees belonging to this project ---
    Dim idxProjet() As Integer
    Dim nbProjet As Integer: nbProjet = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(Trim(rule.NomProjet)) Then
            nbProjet = nbProjet + 1
            ReDim Preserve idxProjet(1 To nbProjet)
            idxProjet(nbProjet) = i
        End If
    Next i
    If nbProjet = 0 Then Exit Sub

    Dim ligne As Long: ligne = 4
    Dim k As Integer
    For k = 1 To nbProjet
        Dim idx As Integer: idx = idxProjet(k)

        Dim cellules(1 To 7) As String
        Dim entrees(1 To 7)  As String
        Dim sorties(1 To 7)  As String
        Dim pDs(1 To 7)      As String
        Dim pFs(1 To 7)      As String
        Dim j As Integer

        Dim utiliseVague As Boolean
        utiliseVague = rule.UtiliseRotationVagues And rule.NombreVagues() > 0

        Dim pauseH As String, pauseF As String
        If utiliseVague Then
            Dim groupeBase As Integer: groupeBase = (k - 1) Mod rule.TailleGroupeVague
            Dim vagueIdx As Integer
            vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod rule.NombreVagues()) + 1
            pauseH = rule.PauseVague(vagueIdx)
            pauseF = AjouterMinutes(pauseH, 60)
        End If

        For j = 1 To 7
            Dim jourDef As ShiftDay
            Set jourDef = rule.Jour(j)

            If jourDef.IsOff Then
                cellules(j) = "OFF"
                entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
            Else
                Dim pD As String, pF As String
                If utiliseVague Then
                    pD = pauseH: pF = pauseF
                Else
                    pD = jourDef.PauseDebut: pF = jourDef.PauseFin
                End If
                cellules(j) = FormatCelluleJour(jourDef.HeureDebut, jourDef.HeureFin, pD, pF)
                entrees(j) = jourDef.HeureDebut
                sorties(j) = jourDef.HeureFin
                pDs(j) = pD
                pFs(j) = pF
            End If
        Next j

        AppliquerCongesEtTT cellules, entrees, sorties, pDs, pFs, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entrees, sorties, pDs, pFs
        ligne = ligne + 1
    Next k

    If ligne > 4 Then
        If rule.UtiliseRotationVagues Then
            ligne = ligne + 1
            ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES " & rule.NomProjet & _
                " (groupes ~" & rule.TailleGroupeVague & ", rotation hebdo)"
            ws.Cells(ligne, 1).Font.Bold = True
            ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
            Dim v As Integer
            For v = 1 To rule.NombreVagues()
                ws.Cells(ligne + v, 1).Value = "Vague " & v & " : " & _
                    rule.PauseVague(v) & "-" & AjouterMinutes(rule.PauseVague(v), 60)
            Next v
            AppliquerBorduresH ws, 4, ligne - 2
        Else
            ws.Cells(ligne + 1, 1).Value = rule.FooterNote
            ws.Cells(ligne + 1, 1).Font.Italic = True
            ws.Cells(ligne + 1, 1).Font.Color = RGB(31, 73, 125)
            AppliquerBorduresH ws, 4, ligne - 1
        End If
    End If
End Sub
