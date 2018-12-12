Option Explicit

 

Private proc_scores() As Double

Private agents_proc_scores As Collection ' of arrays of Doubles

Private esat_scores() As Double

Private agents_esat_scores As Collection ' of arrays of Doubles

Private evals_missing_esat As Collection '

Private sp_evaldate_collection As Collection ' of arrays of Date

Private oAgentsMissingSecondaryScore As Collection ' String (agent_name) _

pointing to array of Date for those evaluation dates with insufficient secondary score

Private all_evaluations() As Evaluation ' To be sorted with func ArrayFilterEvaluation and sub Module1.treeSortExpandTree

Private bClientSatisfaction As Boolean

Private bNNA As Boolean

Private bCRD As Boolean

Private bMSSN As Boolean

Private bFirstAgent As Boolean

Private duplicate_evaluation As Boolean

Private first_metric As Boolean

Private first_eval As Boolean

Private logic_switch_token As Boolean

Private business_expectation_na_resolved As Boolean

Private verification_comment_supplied As Boolean

Private sCurrentAgent As String

Private iEvalCollectionIndex As Integer

Private sPreviousAgent As String

Private oCurrentWorkingEvaluation As Evaluation

Private fArrayFilterEvaluation As ArrayFilterEvaluation

Private this_agent_eval_qty As Integer

       

 

Private Sub Class_Initialize()

    Set agents_proc_scores = New Collection

    Set agents_esat_scores = New Collection

    Set sp_evaldate_collection = New Collection

    bClientSatisfaction = True

    bNNA = False

    bCRD = False

    bMSSN = False

    duplicate_evaluation = False

    first_metric = True

    iEvalCollectionIndex = 0

    this_agent_eval_qty = 0

    bFirstAgent = True

    first_eval = True

    business_expectation_na_resolved = False

    verification_comment_supplied = False

    Set fArrayFilterEvaluation = New ArrayFilterEvaluation

    Set oAgentsMissingSecondaryScore = New Collection

End Sub

 

Public Sub setClientSatisfaction(bIsCSat As Boolean)

    bClientSatisfaction = bIsCSat

End Sub

 

Public Sub setIsCrd(bIsCrd As Boolean)

    bCRD = bIsCrd

End Sub

 

Public Sub setIsNna(bIsNna As Boolean)

    bNNA = bIsNna

End Sub

 

Public Sub setIsMssnMisc(bIsMssn As Boolean)

    bMSSN = bIsMssn

End Sub

 

Public Function getPrimaryAvgForAgent(sAgentName As String) As Double

    Dim aScoreArr() As Double

    Dim iIndex As Integer

    Dim dSubtotal As Double

    aScoreArr = agents_proc_scores.Item(sAgentName)

    dSubtotal = 0

    For iIndex = LBound(aScoreArr) To UBound(aScoreArr)

        dSubtotal = dSubtotal + aScoreArr(iIndex)

    Next iIndex

    getPrimaryAvgForAgent = dSubtotal / CDbl(UBound(aScoreArr) - LBound(aScoreArr) + 1)

End Function

 

Public Function getSecondaryAvgForAgent(sAgentName As String) As Double

    If isSecondaryAvgAvailableForAgent(sAgentName) Then

        Dim aScoreArr() As Double

        Dim iIndex As Integer

        Dim dSubtotal As Double

        dSubtotal = 0

        aScoreArr = agents_esat_scores.Item(sAgentName)

       

        For iIndex = LBound(aScoreArr) To UBound(aScoreArr)

            dSubtotal = dSubtotal + aScoreArr(iIndex)

        Next iIndex

        getSecondaryAvgForAgent = dSubtotal / CDbl(UBound(aScoreArr) - LBound(aScoreArr) + 1)

    Else

        getSecondaryAvgForAgent = 0

    End If

End Function

 

Public Function isSecondaryAvgAvailableForAgent(sAgentName As String) As Boolean

  Dim dates() As Date

  Dim thisDate As Date

  Dim bDateFound As Boolean

  bDateFound = False

  If frmReportBuilderSubmit.isKeyOfCollection(oAgentsMissingSecondaryScore, sAgentName) Then

    isSecondaryAvgAvailableForAgent = False

  Else

    isSecondaryAvgAvailableForAgent = True

  End If

End Function

 

Public Function isSecondaryAvgAvailableForEval(sAgentName As String, oEvalDate As Date) As Boolean

  Dim dates() As Date

  Dim thisDate As Date

  Dim el As Variant

  Dim bDateFound As Boolean

  bDateFound = False

  If frmReportBuilderSubmit.isKeyOfCollection(oAgentsMissingSecondaryScore, sAgentName) Then

    dates = oAgentsMissingSecondaryScore(sAgentName)

    For Each el In dates

      thisDate = el

      If thisDate = oEvalDate Then

        isSecondaryAvgAvailableForEval = False

        Exit Function

      End If

    Next el

  End If

  isSecondaryAvgAvailableForEval = True

End Function

 

 

Public Function getPrimaryAvgForAll() As Double

  Dim dSubtotal As Double

  Dim iIndex As Integer

  For iIndex = LBound(proc_scores) To UBound(proc_scores)

    dSubtotal = dSubtotal + proc_scores(iIndex)

  Next iIndex

  getPrimaryAvgForAll = dSubtotal / CDbl(UBound(proc_scores) - LBound(proc_scores) + 1)

End Function

 

Public Function getSecondaryAvgForAll() As Double

  Dim dSubtotal As Double

  Dim iIndex As Integer

  If (Not Not esat_scores) = 0 Then

    getSecondaryAvgForAll = 0#

  Else

    For iIndex = LBound(esat_scores) To UBound(esat_scores)

      dSubtotal = dSubtotal + esat_scores(iIndex)

    Next iIndex

    getSecondaryAvgForAll = dSubtotal / CDbl(UBound(esat_scores) - LBound(esat_scores) + 1)

  End If

End Function

 

Public Function isSecondaryAvgAvailableForAll() As Boolean

  isSecondaryAvgAvailableForAll = oAgentsMissingSecondaryScore.Count < 1

End Function

 

Public Sub insertData(metric_type As Variant, comment As Variant)

    'Core evaluation parsing code below

        'If IsNumeric(r.Value) Then

        '    metric_type = Format(r.Value, "Long Date") & " " & Format(r.Value, "Long Time")

        'Else

        Dim oTempEvaluation As Evaluation

        Dim temp_esatscore_arr() As Double

        Dim this_sm_proc() As Double

        Dim aDates() As Date

        Dim current_eval As Date

        Dim dScore As Double

       

        ' Begin scenarios

        If Not bFirstAgent And (IsDate(metric_type) Or IsNumeric(metric_type)) And Not IsEmpty(comment) And IsNumeric(comment) Then

            logic_switch_token = getLogicSwitchToken

            If frmReportBuilderSubmit.isKeyOfCollection(agents_proc_scores, sCurrentAgent) And Not duplicate_evaluation _

                And Not first_metric Then

              If Not bClientSatisfaction Then

                Call oCurrentWorkingEvaluation.addCallHandlingMax

              End If

              Call oCurrentWorkingEvaluation.processRaw

              Call oCurrentWorkingEvaluation.setSortOrder

              If Not oCurrentWorkingEvaluation.hasSecondaryScore Then

                Call processMissingEsat

              End If

              'oTempEvaluation.isClientSatisfaction Then

              ' agents_esat_scores As Collection ' of arrays of EvalProcEsatTypeDate

              If oCurrentWorkingEvaluation.isSecondaryScoreValid Then

                If frmReportBuilderSubmit.isCollectionKey(sCurrentAgent, agents_esat_scores) Then

                  temp_esatscore_arr = agents_esat_scores(sCurrentAgent)

                  ReDim Preserve temp_esatscore_arr(0 To UBound(temp_esatscore_arr) + 1)

                  agents_esat_scores.Remove (sCurrentAgent)

                Else

                  ReDim temp_esatscore_arr(0 To 0)

                End If

                temp_esatscore_arr(UBound(temp_esatscore_arr)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

                agents_esat_scores.Add key:=sCurrentAgent, Item:=temp_esatscore_arr

                If (Not Not esat_scores) <> 0 Then

                  ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

                Else

                  ReDim esat_scores(0 To 0)

                End If

                esat_scores(UBound(esat_scores)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

              End If

           

              If Not oCurrentWorkingEvaluation.isVerificationFound Then

                Call oCurrentWorkingEvaluation.processDefaultVerification

              End If

              Call Module1.treeSortExpandTree(all_evaluations, getNewEvaluationIndex( _

                  all_evaluations, oCurrentWorkingEvaluation _

                  ), oCurrentWorkingEvaluation)

            End If

'            If isAgentsFirstEvalPopulated And Not duplicate_evaluation Then

'                Call incrementOffsetOmnibus

'                temp_adate = sp_evaldate_collection(sCurrentAgent)

'                Call addAgentNameOmnibus(sCurrentAgent)

'                temp_date = temp_adate(UBound(temp_adate))

'                Call addTimeStampOmnibus(temp_date)

'                Call addMetricTypeOmnibus("FILLER-IGNORE ME")

'                Call formatFillerRowOmnibus

'            End If

            current_eval = generateEvaluationCreationDate(metric_type)

            dScore = comment

            duplicate_evaluation = isAgentEvalDateKnown(sCurrentAgent, current_eval)

 

            If Not duplicate_evaluation Then

              Call createNewCurrentEvaluation

              Call oCurrentWorkingEvaluation.setRawComment(current_eval, dScore)

              If frmReportBuilderSubmit.isKeyOfCollection(sp_evaldate_collection, sCurrentAgent) Then

                aDates = sp_evaldate_collection(sCurrentAgent)

                sp_evaldate_collection.Remove (sCurrentAgent)

                ReDim Preserve aDates(LBound(aDates) To UBound(aDates) + 1)

              Else

                ReDim aDates(0 To 0)

              End If

              aDates(UBound(aDates)) = current_eval

              sp_evaldate_collection.Add key:=sCurrentAgent, Item:=aDates

              If frmReportBuilderSubmit.isKeyOfCollection(agents_proc_scores, sCurrentAgent) Then

                temp_esatscore_arr = agents_proc_scores(sCurrentAgent)

                ReDim Preserve temp_esatscore_arr(LBound(temp_esatscore_arr) To UBound(temp_esatscore_arr) + 1)

                agents_proc_scores.Remove (sCurrentAgent)

              Else

                ReDim temp_esatscore_arr(0 To 0)

              End If

              temp_esatscore_arr(UBound(temp_esatscore_arr)) = dScore

              agents_proc_scores.Add key:=sCurrentAgent, Item:=temp_esatscore_arr

              If (Not Not proc_scores) <> 0 Then

                ReDim Preserve proc_scores(LBound(proc_scores) To UBound(proc_scores) + 1)

              Else

                ReDim proc_scores(0 To 0)

              End If

              proc_scores(UBound(proc_scores)) = dScore

            End If

        ElseIf TypeName(metric_type) = "String" And metric_type = "Group:" Then

           ' Do nothing

        ElseIf TypeName(metric_type) = "String" And metric_type = "Agent:" Then

            logic_switch_token = getLogicSwitchToken()

            If Not bFirstAgent And logic_switch_token And Not duplicate_evaluation Then

                If Not bClientSatisfaction Then

                  Call oCurrentWorkingEvaluation.addCallHandlingMax

                End If

                Call oCurrentWorkingEvaluation.processRaw

                Call oCurrentWorkingEvaluation.setSortOrder

                If Not oCurrentWorkingEvaluation.hasSecondaryScore Then

                    Call processMissingEsat

                End If

                'oTempEvaluation.isClientSatisfaction Then

                    ' agents_esat_scores As Collection ' of arrays of EvalProcEsatTypeDate

                If oCurrentWorkingEvaluation.isSecondaryScoreValid Then

                    If frmReportBuilderSubmit.isCollectionKey(sCurrentAgent, agents_esat_scores) Then

                        temp_esatscore_arr = agents_esat_scores(sCurrentAgent)

                        ReDim Preserve temp_esatscore_arr(0 To UBound(temp_esatscore_arr) + 1)

                        agents_esat_scores.Remove (sCurrentAgent)

                    Else

                        ReDim temp_esatscore_arr(0 To 0)

                    End If

                    temp_esatscore_arr(UBound(temp_esatscore_arr)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

                    agents_esat_scores.Add key:=sCurrentAgent, Item:=temp_esatscore_arr

                    If (Not Not esat_scores) <> 0 Then

                        ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

                    Else

                        ReDim esat_scores(0 To 0)

                    End If

                    esat_scores(UBound(esat_scores)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

                       

'                        If sm_known Then

'                            If frmReportBuilderSubmit.isKeyOfCollection(sm_esat_scores, current_sm) Then

'                                this_sm_proc = sm_esat_scores(current_sm)

'                                sm_esat_scores.Remove current_sm

'                                ReDim Preserve this_sm_proc(LBound(this_sm_proc) To UBound(this_sm_proc) + 1)

'                            Else

'                                ReDim this_sm_proc(0 To 0)

'                            End If

'                            this_sm_proc(UBound(this_sm_proc)) = temp_esateval.esat

'                            sm_esat_scores.Add key:=current_sm, Item:=this_sm_proc

'                        End If

                        ' Put the values on the row

'                       Call addEsatScoreOmnibus(temp_esateval.esat)

                        'garbage_text = getEsatScore(output_row_offset - 1, primary_output_tab_n)

 

                End If

               

                If Not oCurrentWorkingEvaluation.isVerificationFound Then

                    Call oCurrentWorkingEvaluation.processDefaultVerification

                End If

'                If frmReportBuilderSubmit.SheetExists(sCurrentAgent) Then 'And Not first_eval Then

'                    Call incrementOffsetOmnibus

'                    temp_adate = sp_evaldate_collection(sCurrentAgent)

'                    'Call addAgentNameOmnibus(sCurrentAgent)

'                    temp_date = temp_adate(UBound(temp_adate))

'                    'Call addTimeStampOmnibus(temp_date)

'                    'Call addMetricTypeOmnibus("FILLER-IGNORE ME")

'                    'Call formatFillerRowOmnibus

'                End If

              Call Module1.treeSortExpandTree(all_evaluations, getNewEvaluationIndex( _

                  all_evaluations, oCurrentWorkingEvaluation _

                  ), oCurrentWorkingEvaluation)

            End If

            bFirstAgent = False

            sPreviousAgent = sCurrentAgent

            sCurrentAgent = comment

            this_agent_eval_qty = 0

            verification_comment_supplied = False

            first_eval = True

            first_metric = True

            business_expectation_na_resolved = False

            Call createNewCurrentEvaluation

        ElseIf TypeName(metric_type) = "String" And metric_type = "Evaluation Date" Then

            ' Do nothing

        ElseIf TypeName(metric_type) = "String" And metric_type = "Form:" Then

            ' Do nothing

        ElseIf TypeName(metric_type) = "String" And metric_type = "Report Period:" Then

            ' Do nothing

        ElseIf TypeName(metric_type) = "String" And Len(metric_type) = 0 Then

            'Exit For

        ElseIf Not bFirstAgent And Not duplicate_evaluation And TypeName(metric_type) = "String" Then

          Call oCurrentWorkingEvaluation.setRawComment(metric_type, comment)

          If Not oCurrentWorkingEvaluation.getMetadataFirst(comment, metric_type) = "" And _

              frmReportBuilderSubmit.SuppliedEvalType( _

              oCurrentWorkingEvaluation.getMetadataFirst(comment, metric_type)) = "" Then

            first_metric = False

          End If

        End If

End Sub

 

Private Sub processMissingEsat()

  Dim dates() As Date

  If frmReportBuilderSubmit.isKeyOfCollection(oAgentsMissingSecondaryScore, oCurrentWorkingEvaluation.getAgentName) Then

    dates = oAgentsMissingSecondaryScore(oCurrentWorkingEvaluation.getAgentName)

    ReDim Preserve dates(LBound(dates) To UBound(dates) + 1)

    oAgentsMissingSecondaryScore.Remove oCurrentWorkingEvaluation.getAgentName

  Else

    ReDim dates(0 To 0)

  End If

  dates(UBound(dates)) = oCurrentWorkingEvaluation.getCurrentTimeStamp()

  oAgentsMissingSecondaryScore.Add key:=oCurrentWorkingEvaluation.getAgentName(), Item:=dates

End Sub

 

Private Function generateEvaluationCreationDate(raw_input As Variant) As Date

  Dim oWorkingDate As Date

  'If InStr(1, raw_input, " ") = 0 Then

    generateEvaluationCreationDate = str(CDate(raw_input))

  'Else

  '  oWorkingDate = raw_input

  '  generateEvaluationCreationDate = Format(Mid(oWorkingDate, 1, InStr(1, oWorkingDate, " ") - 1), "Long Date")

  'End If

 

End Function

 

Private Function getLogicSwitchToken() As Boolean

    If Not first_metric Then

        getLogicSwitchToken = Not first_metric

    Else

        If (Not Not proc_scores) <> 0 Then ' proc_scores is  not empty, check to see if there is more than one entry

            getLogicSwitchToken = UBound(proc_scores) > LBound(proc_scores)

        Else

            getLogicSwitchToken = Not first_metric ' False

        End If

    End If

End Function

 

Public Function isAgentEvalDateKnown(sAgentName As String, oEvalDate As Date) As Boolean

    Dim bAnswer As Boolean

    If Not frmReportBuilderSubmit.isKeyOfCollection(sp_evaldate_collection, sAgentName) Then

        isAgentEvalDateKnown = False

    Else

        Dim aAgentEvals() As Date

        aAgentEvals = sp_evaldate_collection(sAgentName)

        Dim iAgentEvalIndex As Integer

        For iAgentEvalIndex = LBound(aAgentEvals) To UBound(aAgentEvals)

            If aAgentEvals(iAgentEvalIndex) = oEvalDate Then

                isAgentEvalDateKnown = True

                Exit Function

            End If

        Next iAgentEvalIndex

        isAgentEvalDateKnown = False

    End If

End Function

 

Public Sub resolveFinalEvaluation()

  Dim aDoubles() As Double

  If frmReportBuilderSubmit.isKeyOfCollection(agents_proc_scores, sCurrentAgent) And Not duplicate_evaluation _

        And Not first_metric Then

    If Not bClientSatisfaction Then

      Call oCurrentWorkingEvaluation.addCallHandlingMax

    End If

    Call oCurrentWorkingEvaluation.processRaw

    Call oCurrentWorkingEvaluation.setSortOrder

    If Not oCurrentWorkingEvaluation.hasSecondaryScore Then

      Call processMissingEsat

    End If

    'oTempEvaluation.isClientSatisfaction Then

    ' agents_esat_scores As Collection ' of arrays of EvalProcEsatTypeDate

    If oCurrentWorkingEvaluation.isSecondaryScoreValid Then

      If frmReportBuilderSubmit.isCollectionKey(sCurrentAgent, agents_esat_scores) Then

        aDoubles = agents_esat_scores(sCurrentAgent)

        ReDim Preserve aDoubles(0 To UBound(aDoubles) + 1)

        agents_esat_scores.Remove (sCurrentAgent)

      Else

        ReDim aDoubles(0 To 0)

      End If

      aDoubles(UBound(aDoubles)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

      agents_esat_scores.Add key:=sCurrentAgent, Item:=aDoubles

      If (Not Not esat_scores) <> 0 Then

        ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

      Else

        ReDim esat_scores(0 To 0)

      End If

      esat_scores(UBound(esat_scores)) = oCurrentWorkingEvaluation.getCurrentSecondaryScore

    End If

            

    If Not oCurrentWorkingEvaluation.isVerificationFound Then

      Call oCurrentWorkingEvaluation.processDefaultVerification

    End If

    Call Module1.treeSortExpandTree(all_evaluations, getNewEvaluationIndex( _

        all_evaluations, oCurrentWorkingEvaluation _

        ), oCurrentWorkingEvaluation)

  End If

End Sub

 

' move index forward

Public Sub moveIndexForward()

    If isIndexValid Then

        Dim oCurrentEval As Evaluation

        Set oCurrentEval = all_evaluations(iEvalCollectionIndex)

        If oCurrentEval.isIndexValid Then

            oCurrentEval.moveIndexForward

            If Not oCurrentEval.isIndexValid Then

              Me.moveIndexForward

            End If

        Else

            iEvalCollectionIndex = iEvalCollectionIndex + 1

       End If

    End If

End Sub

 

Public Function isIndexValid() As Boolean

    Dim oCurrentEval As Evaluation

    If iEvalCollectionIndex > (LBound(all_evaluations) - 1) And iEvalCollectionIndex < (UBound(all_evaluations) + 1) Then

      Set oCurrentEval = all_evaluations(iEvalCollectionIndex)

      If ((UBound(all_evaluations) - LBound(all_evaluations) + 1) > iEvalCollectionIndex And oCurrentEval.isIndexValid) _

          Or (UBound(all_evaluations) - LBound(all_evaluations) + 1) > iEvalCollectionIndex Then

        isIndexValid = True

      End If

    Else

      isIndexValid = False

    End If

End Function

 

Public Sub resetIndex()

    iEvalCollectionIndex = 0

    Dim entry As Variant

    Dim eval As Evaluation

    For Each entry In all_evaluations

      Set eval = entry

      Call eval.resetEvaluationIterator

    Next entry

End Sub

 

Public Function getCurrentEvalOriginalComment()

  Dim oThisEval As Evaluation

  Set oThisEval = all_evaluations(iEvalCollectionIndex)

  getCurrentEvalOriginalComment = oThisEval.getCurrentOriginalComment()

End Function

 

Public Function getCurrentEvalMetricType()

  Dim oThisEval As Evaluation

  Set oThisEval = all_evaluations(iEvalCollectionIndex)

  getCurrentEvalMetricType = oThisEval.getCurrentMetricType

End Function

 

Public Function getCurrentEvalComment()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalComment = oThisComment.getCurrentComment

End Function

 

Public Function getCurrentEvalEvalType()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalEvalType = oThisComment.getCurrentEvalType()

End Function

 

Public Function getCurrentEvalScoreLabel()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalScoreLabel = oThisComment.getCurrentScoreLabel

End Function

 

Public Function getCurrentEvalMetricScore()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalMetricScore = oThisComment.getCurrentMetricScore

End Function

 

Public Function getCurrentEvalMaxScore()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalMaxScore = oThisComment.getCurrentMaxScore()

End Function

 

Public Function getCurrentEvalMetricPercentage() As Variant

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalMetricPercentage = oThisComment.getCurrentMetricPercentage()

End Function

 

Public Function getCurrentEvalPrimaryScore()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalPrimaryScore = oThisComment.getCurrentPrimaryScore()

End Function

 

Public Function getCurrentEvalSecondaryScore()

    Dim oThisComment As Evaluation

    Set oThisComment = all_evaluations(iEvalCollectionIndex)

    getCurrentEvalSecondaryScore = oThisComment.getCurrentSecondaryScore()

End Function

 

Public Function getCurrentEvalAgentName() As String

  Dim oThisComment As Evaluation

  Set oThisComment = all_evaluations(iEvalCollectionIndex)

  getCurrentEvalAgentName = oThisComment.getAgentName

End Function

 

Public Function getCurrentEvalTimeStamp()

  Dim oThisComment As Evaluation

  Set oThisComment = all_evaluations(iEvalCollectionIndex)

  getCurrentEvalTimeStamp = oThisComment.getCurrentTimeStamp

End Function

 

Public Function getCurrentEvalGeneralStatistics() As EvalProcEsatTypeDate

  Dim oThisComment As Evaluation

  Set oThisComment = all_evaluations(iEvalCollectionIndex)

  Set getCurrentEvalGeneralStatistics = oThisComment.getGeneralStatistics()

End Function

 

Public Function isCurrentRowMetadataBadFormat(Optional sComment As String, Optional sMetricType As String) As Boolean

  Dim oThisComment As Evaluation

  Set oThisComment = all_evaluations(iEvalCollectionIndex)

  isCurrentRowMetadataBadFormat = oThisComment.isMetadataBadFormat(sComment, sMetricType)

End Function

 

Public Function isCurrentEvalSecondaryScoreValid() As Boolean

  Dim oThisComment As Evaluation

  Set oThisComment = all_evaluations(iEvalCollectionIndex)

  isCurrentEvalSecondaryScoreValid = oThisComment.isSecondaryScoreValid

End Function

 

' Returns the index integer in all_evaluations where the oCurrentWorkingEvaluation should be added

Private Function getNewEvaluationIndex(ByRef aEvaluations() As Evaluation, ByRef oNewEval As Evaluation) As Integer

  Dim iAllEvalsIndex As Integer

  Dim bFoundLesser As Boolean

  bFoundLesser = False

  On Error Resume Next

  If (Not Not aEvaluations) = 0 Or UBound(aEvaluations) < 0 Then

    getNewEvaluationIndex = 0

  Else

    For iAllEvalsIndex = LBound(aEvaluations) To UBound(aEvaluations)

      If fArrayFilterEvaluation.IArrayFilter_evaluate(aEvaluations(iAllEvalsIndex), oNewEval) = -1 Then

        Exit For

      End If

    Next iAllEvalsIndex

    getNewEvaluationIndex = iAllEvalsIndex

  End If

  On Error GoTo 0

End Function

' Creates a new instance of class Evaluation at global variable oCurrentWorkingEvaluation

Private Sub createNewCurrentEvaluation()

  Set oCurrentWorkingEvaluation = New Evaluation

  Call oCurrentWorkingEvaluation.setAgentName(sCurrentAgent)

  Call oCurrentWorkingEvaluation.setClientSatisfaction(bClientSatisfaction)

  Call oCurrentWorkingEvaluation.setIsCrd(bCRD)

  Call oCurrentWorkingEvaluation.setIsMssnMisc(bMSSN)

  Call oCurrentWorkingEvaluation.setIsNna(bNNA)

End Sub
