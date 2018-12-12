Option Explicit

Implements IArrayFilter

 

Dim aMetricTypes() As String

 

Public Function IArrayFilter_evaluate(ByRef incumbent As Variant, Optional ByRef oNextIndex As Variant) As Integer

    Dim EcIncumbent As EvaluationComment

    Dim EcNextIndex As EvaluationComment

    On Error Resume Next

    If Not oNextIndex Is Nothing And Not oNextIndex Is Null Then

        Set EcIncumbent = incumbent

    End If

    Set EcNextIndex = oNextIndex

    On Error GoTo 0

    If getTypeIndex(EcIncumbent.getMetricType) < getTypeIndex(EcNextIndex.getMetricType) Then

        IArrayFilter_evaluate = 1

    ElseIf getTypeIndex(EcIncumbent.getMetricType) = getTypeIndex(EcNextIndex.getMetricType) Then

        IArrayFilter_evaluate = 0

    Else

        IArrayFilter_evaluate = -1

    End If

End Function

' Pass an array of Strings to define the order the EvaluationComment.EvalType should be

Public Sub initialize(ByRef some_arg As Variant)

    Dim type_list_ordered() As String

    type_list_ordered = some_arg

    ReDim aMetricTypes(LBound(type_list_ordered) To UBound(type_list_ordered))

    aMetricTypes = type_list_ordered

End Sub

 

Private Function getTypeIndex(sThisType As String) As Integer

    Dim iAnswer As Integer

    For iAnswer = LBound(aMetricTypes) To UBound(aMetricTypes)

        If aMetricTypes(iAnswer) = sThisType Then

            getTypeIndex = iAnswer

            Exit For

        End If

    Next iAnswer

    getTypeIndex = iAnswer

End Function
