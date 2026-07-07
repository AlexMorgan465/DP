Case 6

    Dim weekNo As Long
    Dim semaineRotation As Integer
    Dim nom As String

    weekNo = Application.WorksheetFunction.IsoWeekNum(dayDate)
    semaineRotation = (weekNo - 1) Mod 2
    nom = LCase(nomComplet)

    If semaineRotation = 0 Then
        ' Semaine 1 : El Bahlouly + Chouifi travaillent
        If InStr(nom, "chouifi") > 0 _
        Or InStr(nom, "el bahlouly") > 0 _
        Or InStr(nom, "elbahlouly") > 0 Then

            entreeH = 7
            sortieH = 11

        Else
            isOff = True
            comment = "OFF"
        End If

    Else
        ' Semaine 2 : Mounaji travaille
        If InStr(nom, "mounaji") > 0 Then

            entreeH = 7
            sortieH = 11

        Else
            isOff = True
            comment = "OFF"
        End If

    End If
