Option Explicit

Sub GenererPlanning()

    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long

    Set ws = Worksheets("ACCESSIBILITE")

    'Première ligne des collaborateurs
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 6 To derniereLigne

        'Lundi
        ws.Cells(i, 4) = "08:00 - 18:00" & vbCrLf & _
                         "Pause : 13:00-14:00"

        'Mardi
        ws.Cells(i, 5) = "08:00 - 18:00" & vbCrLf & _
                         "Pause : 13:00-14:00"

        'Mercredi
        ws.Cells(i, 6) = "08:00 - 18:00" & vbCrLf & _
                         "Pause : 13:00-14:00"

        'Jeudi
        ws.Cells(i, 7) = "08:00 - 18:00" & vbCrLf & _
                         "Pause : 13:00-14:00"

        'Vendredi
        ws.Cells(i, 8) = "08:00 - 17:00" & vbCrLf & _
                         "Pause : 13:00-14:00"

        'Samedi
        ws.Cells(i, 9) = "OFF"

        'Dimanche
        ws.Cells(i, 10) = "OFF"

        'Calcul des heures
        ws.Cells(i, 11) = 44

    Next i

    MsgBox "Planning généré avec succès !"

End Sub
