Option Explicit

 

Private Note As String

Private score As Variant

Private cmnt_type As Long ' EvalCmntType

Private score_max As Variant

Private score_percentage As Long

Private yes_no As String

Private is_garbage As Boolean

Private static_types() As EvalCmntType

Private dynamic_types() As EvalCmntType

Private non_partial_types() As EvalCmntType

 

Private cells() As Variant

 

Private agent_name_clmn As Integer

Private metric_type_clmn As Integer

Private comment_clmn As Integer

Private eval_type_clmn As Integer

Private time_stamp_clmn As Integer

Private score_label_clmn As Integer

Private metric_score_clmn As Integer

Private max_score_clmn As Integer

Private metric_pct_clmn As Integer

Private primary_score_clmn As Integer

Private secondary_score_clmn As Integer

Private bad_format_comment As String

 

Public Enum EvalCmntType

    Comment1 = 0

    EvaluatorSatisfaction = 1

    Verification1 = 2

    BusinessComment = 16

    AccurateInformation = 4

    Process__002F__Procedures = 5

    Expectations = 6

    Hold__002F__Transfer = 7

    CallLog = 8

    Added__002F__Updated = 9

    Survey100 = 10

    Callback = 11

    Opening__002F__Farewell = 12

    ActivelyListened = 13

    ControlledCall = 14

    Clear__002F__Confident = 15

    HoldComment = 3

    Garbage17 = 17

    Garbage18 = 18

    Garbage19 = 19

End Enum

 

Private Sub Class_Initialize()

    Dim i As Integer

    ReDim static_types(0 To 7)

   

    agent_name_clmn = 0

    metric_type_clmn = 1

    comment_clmn = 2

    eval_type_clmn = 3

    time_stamp_clmn = 4

    score_label_clmn = 5

    metric_score_clmn = 6

    max_score_clmn = 7

    metric_pct_clmn = 8

    primary_score_clmn = 9

    secondary_score_clmn = 10

    ReDim cells(agent_name_clmn To secondary_score_clmn)

   

    static_types(0) = EvaluatorSatisfaction

    static_types(1) = AccurateInformation

    static_types(2) = Process__002F__Procedures

    static_types(3) = Expectations

    static_types(4) = Opening__002F__Farewell

    static_types(5) = ActivelyListened

    static_types(6) = ControlledCall

    static_types(7) = Clear__002F__Confident

   

    ReDim dynamic_types(0 To 4)

    dynamic_types(0) = Hold__002F__Transfer

    dynamic_types(1) = CallLog

    dynamic_types(2) = Added__002F__Updated

    dynamic_types(3) = Survey100

    dynamic_types(4) = Callback

   

    ReDim non_partial_types(0 To 4)

    non_partial_types(0) = Hold__002F__Transfer

    non_partial_types(1) = CallLog

    non_partial_types(2) = Added__002F__Updated

    non_partial_types(3) = Survey100

    non_partial_types(4) = Callback

   

    is_garbage = False

End Sub

 

Public Sub setGarbageTrue()

    is_garbage = True

End Sub

 

Public Sub setOriginalComment(sRawComment As String)

  bad_format_comment = sRawComment

End Sub

 

Public Sub setAgent(sEvalAgent As String)

    cells(agent_name_clmn) = sEvalAgent

End Sub

 

Public Sub setMetricType(sMetricType As String)

    cells(metric_type_clmn) = sMetricType

End Sub

 

Public Sub setComment(sComment As String)

    cells(comment_clmn) = sComment

End Sub

 

Public Sub setEvalType(sEvalType As String)

    cells(eval_type_clmn) = sEvalType

End Sub

 

Public Sub setTimeStamp(oTimeStamp As Date)

    cells(time_stamp_clmn) = oTimeStamp

End Sub

 

Public Sub setScoreLabel(sScoreLabel As String)

    cells(score_label_clmn) = sScoreLabel

End Sub

 

Public Sub setMetricScore(dMetricScore As Double)

    cells(metric_score_clmn) = dMetricScore

End Sub

 

Public Sub setMaxScore(dMaxScore As Double)

    cells(max_score_clmn) = dMaxScore

End Sub

 

Public Sub setMetricPercentage(dMetricPercentage As Double)

    cells(metric_pct_clmn) = dMetricPercentage

End Sub

 

Public Sub setPrimaryScore(dPrimaryScore As Double)

    cells(primary_score_clmn) = dPrimaryScore

End Sub

 

Public Sub setSecondaryScore(dSecondaryScore As Double)

    cells(secondary_score_clmn) = dSecondaryScore

End Sub

 

Public Function getOriginalComment()

  getOriginalComment = bad_format_comment

End Function

 

Public Function isGarbage()

    isGarbage = is_garbage

End Function

 

Public Function getAgent()

    getAgent = cells(agent_name_clmn)

End Function

 

Public Function getMetricType()

    getMetricType = cells(metric_type_clmn)

End Function

 

Public Function getComment()

    getComment = cells(comment_clmn)

End Function

 

Public Function getEvalType()

    getEvalType = cells(eval_type_clmn)

End Function

 

Public Function getTimeStamp()

    getTimeStamp = cells(time_stamp_clmn)

End Function

 

Public Function getScoreLabel()

    getScoreLabel = cells(score_label_clmn)

End Function

 

Public Function getMetricScore()

    getMetricScore = cells(metric_score_clmn)

End Function

 

Public Function getMaxScore()

    getMaxScore = cells(max_score_clmn)

End Function

 

Public Function getMetricPercentage() As Variant

    getMetricPercentage = cells(metric_pct_clmn)

End Function

 

Public Function getPrimaryScore()

    getPrimaryScore = cells(primary_score_clmn)

End Function

 

Public Function getSecondaryScore()

    getSecondaryScore = cells(secondary_score_clmn)

End Function

 

'Public Function compareTo(oOtherEvalComment As EvaluationComment) As Integer

'    Me.

'End Function

 

Public Sub InitiateProperties(metric_type As String, comment As Variant)

 

    ' wrong pipe format - garbage

    ' good pipe, string first arg, type not recognized - garbage

    ' good pipe, string first arg, type non-standard - change metric_type to first arg

    ' good pipe, string first arg, type recognized, ESAT, missing score - garbage

    On Error GoTo GarbageBlock

    If InStr(1, comment, "||", vbBinaryCompare) > 0 And InStr(InStr(1, comment, "||", vbBinaryCompare), comment, "||", vbBinaryCompare) > 0 Then

        Dim meta_text As Variant

        meta_text = LCase(LTrim(RTrim(Mid(comment, InStr(1, comment, "||", vbBinaryCompare) + 2, InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) - (InStr(1, comment, "||", vbBinaryCompare) + 2)))))

        If Not IsNumeric(meta_text) Then

            If (meta_text = "yes" Or meta_text = "partial" Or meta_text = "no" Or meta_text = "n/a" Or meta_text = "") Then

                MetricType = metric_type

                If (meta_text = "n/a" Or meta_text = "") And Not cmnt_type = -1 Then

                    is_garbage = False

                    score = "--"

                    Note = comment

                ElseIf Not cmnt_type = -1 Then

                    score = applyScore(meta_text, cmnt_type)

                    Note = comment

                End If

            Else

                ' Non-standard metric type

                MetricType = meta_text

                If Not cmnt_type = -1 Then

                    ' Looking for Evaluator Satisfaction

                    If InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) > 0 And InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) Then

                        Dim second_arg As String

                        second_arg = Mid(comment, InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) - (InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2))

                        If IsNumeric(second_arg) Then

                            score = second_arg

                            is_garbage = False

                            Note = comment

                        Else

                            is_garbage = True

                        End If

                    Else ' Non-ESAT non-standard metric type

                        score = "--"

                        Note = comment

                        is_garbage = False

                    End If

                Else ' Comment Type not recognized

GarbageBlock:

                    is_garbage = True

                    Note = comment

                    MetricType = metric_type

                    score = "--"

                End If

            End If

        ' Manually provided score for metric

        ElseIf IsNumeric(meta_text) Then

            MetricType = metric_type

            If Not cmnt_type = -1 Then

                score = meta_text

                is_garbage = False

                Note = comment

            End If

       

        End If

    Else

        is_garbage = True

    End If

 

End Sub

Private Property Let CmntType(Value As EvalCmntType)

    cmnt_type = Value

End Property

 

Private Property Get CmntType() As EvalCmntType

    CmntType = cmnt_type

End Property

 

Public Property Get MetricType() As String

    Select Case cmnt_type

        Case EvalCmntType.Comment1

            MetricType = "Comment"

        Case EvalCmntType.EvaluatorSatisfaction

            MetricType = "Evaluator Satisfaction"

        Case EvalCmntType.Verification1

            MetricType = "Verification"

        Case EvalCmntType.BusinessComment

            MetricType = "Business Comment"

        Case EvalCmntType.AccurateInformation

            MetricType = "Accurate Information"

        Case EvalCmntType.Process__002F__Procedures

            MetricType = "Process / Procedures"

        Case EvalCmntType.Expectations

            MetricType = "Expectations"

        Case EvalCmntType.Hold__002F__Transfer

            MetricType = "Hold / Transfer"

        Case EvalCmntType.CallLog

            MetricType = "Call Log"

        Case EvalCmntType.Added__002F__Updated

            MetricType = "Added / Updated"

        Case EvalCmntType.Survey100

            MetricType = "Survey"

        Case EvalCmntType.Callback

            MetricType = "Call Back"

        Case EvalCmntType.Opening__002F__Farewell

            MetricType = "Opening / Farewell"

        Case EvalCmntType.ActivelyListened

            MetricType = "Actively Listened"

        Case EvalCmntType.ControlledCall

            MetricType = "Controlled Call"

        Case EvalCmntType.Clear__002F__Confident

            MetricType = "Clear / Confident"

        Case EvalCmntType.Garbage17

            MetricType = "Comment"

        Case EvalCmntType.Garbage18

            MetricType = "Comment"

        Case EvalCmntType.Garbage19

            MetricType = "Comment"

           

    End Select

End Property

 

Public Property Let MetricType(raw_metric As String)

    Select Case raw_metric

        Case "Actively listened to client and correctly identified the root cause of the call"

            cmnt_type = EvalCmntType.ActivelyListened

        Case "Appropriately controlled the call"

            cmnt_type = EvalCmntType.ControlledCall

        Case "Comment"

            cmnt_type = Comment1

        Case "Communicated in a clear and confident manner"

            cmnt_type = Clear__002F__Confident

        Case "Followed correct processes and procedures"

            cmnt_type = Process__002F__Procedures

        Case "Logged call correctly and added necessary notes"

            cmnt_type = CallLog

        Case "Provided a warm opening and a fond farewell"

            cmnt_type = Opening__002F__Farewell

        Case "Provided accurate information"

            cmnt_type = AccurateInformation

        Case "Set appropriate expectation with client"

           cmnt_type = Expectations

        Case "What is the likelihood that the caller will need to call again due to the agent's handling of the interaction?"

            cmnt_type = Callback

        Case "Added or updated all required information"

            cmnt_type = Added__002F__Updated

        Case "Offered survey at the end of the call"

            cmnt_type = Survey100

        Case "Followed appropriate hold / dial procedure"

            cmnt_type = Hold__002F__Transfer

        Case "evaluator satisfaction"

            cmnt_type = EvaluatorSatisfaction

        Case "business comment"

            cmnt_type = BusinessComment

        Case "business evaluation"

            cmnt_type = BusinessComment

        Case "negative"

            cmnt_type = Comment1

        Case "verbal"

            cmnt_type = Comment1

        Case "verification"

            cmnt_type = Verification1

        Case "survey"

            cmnt_type = Comment1

        Case "busines"

            cmnt_type = Comment1

        Case "written"

            cmnt_type = Comment1

        Case "esat"

            cmnt_type = EvaluatorSatisfaction

        Case "hold comment"

            cmnt_type = HoldComment

        Case Else

            cmnt_type = -1

    End Select

 

End Property

 

Public Function MetricEnumeration(raw_metric As String) As EvalCmntType

    Select Case raw_metric

        Case "Actively listened to client and correctly identified the root cause of the call"

            MetricEnumeration = EvalCmntType.ActivelyListened

        Case "Appropriately controlled the call"

            MetricEnumeration = EvalCmntType.ControlledCall

        Case "Comment"

            MetricEnumeration = Comment1

        Case "Communicated in a clear and confident manner"

            MetricEnumeration = Clear__002F__Confident

        Case "Followed correct processes and procedures"

            MetricEnumeration = Process__002F__Procedures

        Case "Logged call correctly and added necessary notes"

            MetricEnumeration = CallLog

        Case "Provided a warm opening and a fond farewell"

            MetricEnumeration = Opening__002F__Farewell

        Case "Provided accurate information"

            MetricEnumeration = AccurateInformation

        Case "Set appropriate expectation with client"

            MetricEnumeration = Expectations

        Case "What is the likelihood that the caller will need to call again due to the agent's handling of the interaction?"

            MetricEnumeration = Callback

        Case "Added or updated all required information"

            MetricEnumeration = Added__002F__Updated

        Case "Offered survey at the end of the call"

            MetricEnumeration = Survey100

        Case "Followed appropriate hold / dial procedure"

            MetricEnumeration = Hold__002F__Transfer

        Case "evaluator satisfaction"

            MetricEnumeration = EvaluatorSatisfaction

        Case "business comment"

            MetricEnumeration = BusinessComment

        Case "business evaluation"

            MetricEnumeration = BusinessComment

        Case "negative"

            MetricEnumeration = Comment1

        Case "verbal"

            MetricEnumeration = Comment1

        Case "verification"

            MetricEnumeration = Verification1

        Case "survey"

            MetricEnumeration = Comment1

        Case "busines"

            MetricEnumeration = Comment1

        Case "written"

            MetricEnumeration = Comment1

        Case "esat"

            MetricEnumeration = EvaluatorSatisfaction

        Case "hold comment"

            MetricEnumeration = HoldComment

        Case Else

            MetricEnumeration = -1

    End Select

 

End Function

 

 

Private Function IsInArray(ByVal stringToBeFound As String, arr As Variant) As Boolean

  Dim aArrayToSearch() As String

  ReDim aArrayToSearch(LBound(arr) To UBound(arr))

  aArrayToSearch = arr

  IsInArray = (UBound(Filter(aArrayToSearch, stringToBeFound)) > -1)

  Exit Function

End Function

 

Private Function applyScore(ByVal arg As String, ByVal this_type As Long) As Long

    arg = LCase(LTrim(RTrim(arg)))

    If arg = "partial" And IsInArray(this_type, non_partial_types) Then

        Select Case this_type

            Case EvalCmntType.Process__002F__Procedures

                applyScore = 1

                Exit Function

            Case EvalCmntType.AccurateInformation

                applyScore = 0.25

                Exit Function

            Case EvalCmntType.Expectations

                applyScore = 0.25

                Exit Function

            Case EvalCmntType.Opening__002F__Farewell

                applyScore = 0.13

                Exit Function

            Case EvalCmntType.ActivelyListened

                applyScore = 0.13

                Exit Function

            Case EvalCmntType.Clear__002F__Confident

                applyScore = 0.13

                Exit Function

            Case EvalCmntType.ControlledCall

                applyScore = 0.13

        End Select

        yes_no = arg

    ElseIf arg = "yes" And IsInArray(this_type, static_types) Then

        Select Case this_type

            Case EvalCmntType.Process__002F__Procedures

                applyScore = 2

                Exit Function

            Case EvalCmntType.AccurateInformation

                applyScore = 0.5

                Exit Function

            Case EvalCmntType.Expectations

                applyScore = 0.5

                Exit Function

            Case EvalCmntType.Opening__002F__Farewell

                applyScore = 0.25

                Exit Function

            Case EvalCmntType.ActivelyListened

                applyScore = 0.25

                Exit Function

            Case EvalCmntType.Clear__002F__Confident

                applyScore = 0.25

                Exit Function

            Case EvalCmntType.ControlledCall

                applyScore = 0.25

        End Select

        yes_no = arg

    ElseIf arg = "Yes" And IsInArray(this_type, dynamic_types) Then

        yes_no = "Yes"

        Exit Function

    End If

   

End Function
