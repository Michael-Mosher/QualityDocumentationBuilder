Public Sub Reset()

    Dim i As Integer

    Dim delete_sheet As Boolean

    Application.DisplayAlerts = False

    Application.ScreenUpdating = False

    With ActiveWorkbook

        For i = .Worksheets.Count To 1 Step -1

            delete_sheet = False

            With .Worksheets(i)

                If .Visible = xlSheetHidden Or .Visible = xlSheetVeryHidden And Not .name = "Form Constants" Then

                    .Visible = xlSheetVisible

                End If

                If Not .name = "Raw" And Not .name = "Quality Ranking" And Not .name = "SM-SP" And Not .name = "Form Constants" Then

                    delete_sheet = True

                ElseIf .name = "Raw" Then

                    .UsedRange.ClearContents

                End If

            End With

            If delete_sheet Then

                Sheets(i).Delete

            End If

        Next i

    End With

    'dtpBeginDate.Value = Now()

    'dtpEndDate.Value = Now()

    'sltEvalMonth.Value = Month(Now())

    'sltEvalYear.Value = Year(Now())

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

End Sub
