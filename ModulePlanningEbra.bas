If isManager Then

    Dim weekNo As Long
    Dim semaineRotation As Integer
    Dim nom As String

    weekNo = Application.WorksheetFunction.IsoWeekNum(dayDate)
    semaineRotation = (weekNo - 1) Mod 2

    nom = LCase(Trim(nomComplet))

    Select Case dayIndex

        Case 1 To 5

            If InStr(nom, "chouifi") > 0 _
            Or InStr(nom, "el bahlouly") > 0 _
            Or InStr(nom, "elbahlouly") > 0 Then

                entreeH = 8
                sortieH = 17

            Else

                entreeH = 7
                sortieH = 16

            End If

        Case 6

            If semaineRotation = 0 Then
                ' El Bahlouly + Chouifi travaillent

                If InStr(nom, "chouifi") > 0 _
                Or InStr(nom, "el bahlouly") > 0 _
                Or InStr(nom, "elbahlouly") > 0 Then

                    entreeH = 7
                    sortieH = 11

                Else
                    isOff = True
                End If

            Else
                ' Mounaji travaille

                If InStr(nom, "mounaji") > 0 Then

                    entreeH = 7
                    sortieH = 11

                Else
                    isOff = True
                End If

            End If

        Case Else

            isOff = True

    End Select
