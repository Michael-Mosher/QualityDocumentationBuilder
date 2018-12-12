Implements IArrayFilter

Public Function IArrayFilter_evaluate(ByRef incumbent As Variant, Optional ByRef oNextIndex As Variant) As Integer

    Dim oIncumbentEvaluation As Evaluation

    Dim oNextEvaluation As Evaluation

    Set oIncumbentEvaluation = incumbent

    Set oNextEvaluation = oNextIndex

    If oIncumbentEvaluation.getAgentName < oNextEvaluation.getAgentName Then

        IArrayFilter_evaluate = 1

    ElseIf oIncumbentEvaluation.getAgentName = oNextEvaluation.getAgentName Then

        If oIncumbentEvaluation.getCurrentTimeStamp < oNextEvaluation.getCurrentTimeStamp Then

            IArrayFilter_evaluate = 1

        ElseIf oIncumbentEvaluation.getCurrentTimeStamp = oNextEvaluation.getCurrentTimeStamp Then

            IArrayFilter_evaluate = 0

        Else

            IArrayFilter_evaluate = -1

        End If

    Else

        IArrayFilter_evaluate = -1

    End If

End Function
