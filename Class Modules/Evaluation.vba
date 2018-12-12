Option Explicit

Private aRawComments() As Variant ' Array of String[] a = new String[2];

Private bNNA As Boolean

Private bCRD As Boolean

Private bMSSN As Boolean

Private aMetricTypes() As String

Private aEvaluationComments() As EvaluationComment

Private bClientSatisfaction As Boolean

Private oEvalStats As EvalProcEsatTypeDate

Private iRawCommentIterator As Integer

Private bBusinessExpectationsNaFound As Boolean

Private bVerificationFound As Boolean

Private bEsatFound As Boolean

Private sSmName As String

Private iEvaluationIndex As Integer

Private dExpectedSecondaryMax As Double

Private dMaxSecondaryScoreSubtotal As Double

Private bMissingBusinessExpectationsNaProcessed As Boolean

Private bSecondaryScoreValid As Boolean

''''''' Legacy variable to accommodate Sub addCallHandlingMax (This Class doesn't touch any Worksheet objects) '''''

Private iCallHandlingCt As Integer

Private hold_transfer_offset As Integer

Private current_agent_offset As Integer

Private output_row_offset As Integer

Private current_sm_offset As Integer

Private call_log_offset As Integer

Private added_updated_offset As Integer

Private survey_offset As Integer

Private callback_offset As Integer

Private primary_output_tab_n As String

Private current_agent As String

Private sm_known As Boolean

Private current_sm As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 

Public Enum EvaluationType

    Verbal = 0

    business = 1

    Survey100 = 2

    Negative = 3

    Written = 4

End Enum

 

Private Sub Class_Initialize()

  bClientSatisfaction = True

  bNNA = False

  bCRD = False

  bMSSN = False

  iRawCommentIterator = 0

  Set oEvalStats = New EvalProcEsatTypeDate

  bBusinessExpectationsNaFound = False

  bVerificationFound = False

  bEsatFound = False

  iEvaluationIndex = 0

  dMaxSecondaryScoreSubtotal = 0

  bMissingBusinessExpectationsNaProcessed = False

  bSecondaryScoreValid = True

  dExpectedSecondaryMax = 5

    '''''''' Sub addCallHandlingMax legacy -- IGNORE '''''''''''''''''''

    iCallHandlingCt = 0

    hold_transfer_offset = 0

    call_log_offset = 0

    added_updated_offset = 0

    survey_offset = 0

    callback_offset = 0

    current_agent_offset = 2

    output_row_offset = 2

    current_sm_offset = 2

    primary_output_tab_n = "Some Date"

    current_agent = "Chick or Dude"

    sm_known = True

    current_sm = "Napolitano, Jim"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   

End Sub

 

Public Sub setRawComment(left_comment As Variant, right_comment As Variant)

    Dim aNewComment() As Variant

    Dim isEvalTypeEmpty As Boolean

    Dim isPrimaryScoreEmpty As Boolean

    Dim isEvalCreationDateEmpty As Boolean

    ReDim aNewComment(0 To 1)

    aNewComment(0) = left_comment

    aNewComment(1) = right_comment

    If (Not Not aRawComments) = 0 Then

        ReDim aRawComments(0 To 0)

    Else

        ReDim Preserve aRawComments(LBound(aRawComments) To UBound(aRawComments) + 1)

    End If

    aRawComments(UBound(aRawComments)) = aNewComment

    isEvalTypeEmpty = (getCurrentEvalType = "")

    isPrimaryScoreEmpty = (getCurrentPrimaryScore() = 0)

    isEvalCreationDateEmpty = (getCurrentTimeStamp() = 0)

   

    If isEvalTypeEmpty Or isPrimaryScoreEmpty Or isEvalCreationDateEmpty Then

        Call findStatistics(left_comment, right_comment)

    End If

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

Public Sub setAgentName(ByVal sThisEvalAgentName As String)

    oEvalStats.Agent = sThisEvalAgentName

End Sub

 

Private Sub setEvalCreationDate(oCreationDate As Date)

    oEvalStats.edate = oCreationDate

End Sub

 

Private Sub setBusinessExpectationNaFound(bFound As Boolean)

    bBusinessExpectationsNaFound = bFound

End Sub

 

Private Sub setVerificationFound(bFound As Boolean)

    bVerificationFound = bFound

End Sub

 

Private Sub setEsatFound(bFound As Boolean)

    bEsatFound = bFound

End Sub

 

Public Sub setSmName(sSmName As String)

    sSmName = sSmName

End Sub

 

Public Function getSmName() As String

    getSmName = sSmName

End Function

 

Public Function isEsatFound()

    isEsatFound = bEsatFound

End Function

 

Public Function isVerificationFound() As Boolean

    isVerificationFound = bVerificationFound

End Function

 

Public Function isBusinessExpectationNaFound()

    isBusinessExpectationNaFound = bBusinessExpectationsNaFound

End Function

 

Public Function isClientSatisfaction() As Boolean

    isClientSatisfaction = bClientSatisfaction

End Function

 

Public Function isCrd()

    isCrd = bCRD

End Function

 

Public Function isNna()

    isNna = bNNA

End Function

 

Public Function isMssnMisc()

    isMssnMisc = bMSSN

End Function

 

' Call after any setClientSatisfaction, setIsCrd setIsNna setIsMssnMisc that is needed

Public Sub setSortOrder()

    If isClientSatisfaction Then

       ReDim aMetricTypes(0 To 14)

        If isCrd Then

            ReDim aMetricTypes(0 To 17)

        End If

        aMetricTypes(0) = "Comment"

        aMetricTypes(1) = "Verification"

        aMetricTypes(2) = "Accuracy / Completeness"

        aMetricTypes(3) = "Complete Expectations"

        aMetricTypes(4) = "Timely Resolution"

        aMetricTypes(5) = "World-Class Service"

        aMetricTypes(6) = "Hold Experience"

        aMetricTypes(7) = "Transfer Experience"

        aMetricTypes(8) = "Forward Thinking"

        aMetricTypes(9) = "Grammar Error Free"

        aMetricTypes(10) = "Appropriate Greeting"

        If isNna Then

            aMetricTypes(11) = "Promote Branch Self-Service"

        Else

            aMetricTypes(11) = "Correct Resources"

        End If

        aMetricTypes(12) = "Appropriate Closing"

        aMetricTypes(13) = "Business Processes"

        aMetricTypes(14) = "Challenge Processed Correctly"

        If isCrd Then

            aMetricTypes(15) = "Call Log Entered"

            aMetricTypes(16) = "Call Log Details"

            aMetricTypes(17) = "Survey"

        End If

    Else

        ReDim aMetricTypes(0 To 20)

        aMetricTypes(0) = "Comment"

        aMetricTypes(1) = "Evaluator Satisfaction"

        aMetricTypes(2) = "Verification"

        aMetricTypes(3) = "Accurate Information"

        aMetricTypes(4) = "Process / Procedures"

        aMetricTypes(5) = "Expectations"

        aMetricTypes(6) = "Hold / Transfer"

        aMetricTypes(7) = "Call Log"

        aMetricTypes(8) = "Added / Updated"

        aMetricTypes(9) = "Survey"

        aMetricTypes(10) = "Callback"

        aMetricTypes(11) = "Opening / Farewell"

        aMetricTypes(12) = "Actively Listened"

        aMetricTypes(13) = "Controlled Call"

        aMetricTypes(14) = "Clear / Confident"

        aMetricTypes(15) = "Hold Comment"

        aMetricTypes(16) = "Business Comment"

        aMetricTypes(17) = "UNKNOWN COMMENT1"

        aMetricTypes(18) = "UNKNOWN COMMENT2"

        aMetricTypes(19) = "UNKNOWN COMMENT3"

        aMetricTypes(20) = "UNKNOWN COMMENT4"

    End If

    If (Not Not aEvaluationComments) <> 0 Then

        Dim iTempNum As Integer

        Dim iSortedArrayIterator As Integer

        Dim oNextEvalCommentToBeAdded As EvaluationComment

        Dim aSortedEvals() As EvaluationComment

        Dim fSortFunc As ArrayFilterEvaluationComment

        Dim bGreaterFound As Boolean

        'ReDim aSortedEvals(LBound(aEvaluationComments) To UBound(aEvaluationComments))

        Set fSortFunc = New ArrayFilterEvaluationComment

        fSortFunc.initialize (aMetricTypes)

        For iTempNum = LBound(aEvaluationComments) To UBound(aEvaluationComments)

            bGreaterFound = False

            Set oNextEvalCommentToBeAdded = aEvaluationComments(iTempNum)

            If iTempNum = LBound(aEvaluationComments) Then

                ReDim aSortedEvals(0 To 0)

                Set aSortedEvals(0) = oNextEvalCommentToBeAdded

            Else

                For iSortedArrayIterator = LBound(aSortedEvals) To UBound(aSortedEvals)

                    If _

                            fSortFunc.IArrayFilter_evaluate(aSortedEvals(iSortedArrayIterator), oNextEvalCommentToBeAdded) = 1 _

                             Or fSortFunc.IArrayFilter_evaluate(aSortedEvals(iSortedArrayIterator), oNextEvalCommentToBeAdded) = 0 Then

                             ' do nothing

                    Else

                        Call Module1.treeSortExpandTree(aSortedEvals, iSortedArrayIterator, oNextEvalCommentToBeAdded)

                        bGreaterFound = True

                        Exit For

                    End If

                Next iSortedArrayIterator

                If Not bGreaterFound Then

                    Call Module1.treeSortExpandTree(aSortedEvals, UBound(aSortedEvals) + 1, oNextEvalCommentToBeAdded)

                End If

            End If

        Next iTempNum

    End If

End Sub

' Helper method to store an EvaluationComment

Private Sub insertProcessedRow(ByRef row As EvaluationComment)

    If (Not Not aEvaluationComments) = 0 Then

        ReDim aEvaluationComments(0 To 0)

    Else

        ReDim Preserve aEvaluationComments(LBound(aEvaluationComments) To UBound(aEvaluationComments) + 1)

    End If

    Set aEvaluationComments(UBound(aEvaluationComments)) = row

End Sub

' Helper function to find and store the Evaluation Type, the Primary score for the evaluation, and the Secondary score for same.

Private Sub findStatistics(raw_label As Variant, raw_comment As Variant)

    Dim meta_data As String

    'Module1.

    If (IsDate(raw_label) Or IsNumeric(raw_label)) And Not IsEmpty(raw_comment) And IsNumeric(raw_comment) Then

        oEvalStats.edate = raw_label

        oEvalStats.procedural = raw_comment

        If UBound(aRawComments) = 0 Then

            Dim aReplacementComments() As Variant

            aRawComments = aReplacementComments

        Else

            Call Module1.reduceArrayAndPreserve(aRawComments, LBound(aRawComments), UBound(aRawComments) - 1)

        End If

    Else

        If InStr(1, raw_comment, "||", vbBinaryCompare) > 0 And InStr(InStr(1, raw_comment, "||", vbBinaryCompare) + 2, raw_comment, "||", vbBinaryCompare) > 0 Then

            meta_data = LCase(LTrim(RTrim(Mid(raw_comment, InStr(1, raw_comment, "||", vbBinaryCompare) + 2, InStr(InStr(1, raw_comment, "||", vbBinaryCompare) + 2, raw_comment, "||", vbBinaryCompare) - (InStr(1, raw_comment, "||", vbBinaryCompare) + 2)))))

            If Len(frmReportBuilderSubmit.SuppliedEvalType(meta_data)) > 0 Then

                oEvalStats.etype = frmReportBuilderSubmit.SuppliedEvalType(meta_data)

            End If

        End If

    End If

End Sub

 

' To be called by Evaluation Collection when cycling through collection during output just before sort

Public Sub processRaw()

  Dim i As Integer

  Dim sWorkingText As String

  Dim aRawDuple() As Variant

  Dim oThisEval As EvaluationComment

  Dim sFirstMetadata As String

  Dim sSecondMetadata As String

  Dim sThisMetricType As String

  Dim sVerificationAnswer As String

  For i = LBound(aRawComments) To UBound(aRawComments)

    aRawDuple = aRawComments(i)

    Set oThisEval = New EvaluationComment

    oThisEval.setAgent (getAgentName)

    oThisEval.setEvalType (getCurrentEvalType)

    oThisEval.setTimeStamp (getCurrentTimeStamp)

    oThisEval.setPrimaryScore (getCurrentPrimaryScore())

    sWorkingText = getCommentPreMetadataProcessing(aRawDuple(1))

    If isMetadataBadFormat(sWorkingText, aRawDuple(0)) Then

      oThisEval.setGarbageTrue

      oThisEval.setOriginalComment (aRawDuple(1))

    Else

      sFirstMetadata = getMetadataFirst(sWorkingText, aRawDuple(0))

      sSecondMetadata = getMetadataSecond(sWorkingText, aRawDuple(0))

      sThisMetricType = getMetricTypeRevised(aRawDuple(0), sFirstMetadata)

      oThisEval.setComment (getStrippedComment(sWorkingText))

      If Not sThisMetricType = "Comment" And Not sThisMetricType = "Business Comment" And Not sThisMetricType = "Hold Comment" Then

        Call oThisEval.setMaxScore(getThisMaxScore(sThisMetricType, Me, sFirstMetadata))

        Call oThisEval.setMetricScore(frmReportBuilderSubmit.getRevisedMetricScore(sFirstMetadata, oThisEval.getMaxScore, Me, sThisMetricType, isClientSatisfaction, getCurrentEvalType))

        Call oThisEval.setScoreLabel(frmReportBuilderSubmit.getRevisedMetricLabel(oThisEval.getMetricScore, oThisEval.getMaxScore, sThisMetricType))

        If sThisMetricType = "Hold Experience" Or sThisMetricType = "Transfer Experience" And _

          frmReportBuilderSubmit.getRevisedMetricLabel(oThisEval.getMetricScore, oThisEval.getMaxScore, sThisMetricType) = "Yes" And getStrippedComment(sWorkingText) = "" Then

          Call oThisEval.setScoreLabel("N/A")

        End If

      End If

      oThisEval.setMetricType (sThisMetricType)

      If (getMetricSection(sThisMetricType) = "Business Expectations" Or sThisMetricType = "ESAT") And Not isMetadataBadFormat(sWorkingText, aRawDuple(0)) Then

        If getMetricSection(sThisMetricType) = "Business Expectations" Then

          dMaxSecondaryScoreSubtotal = dMaxSecondaryScoreSubtotal + _

              getThisMaxScore(sThisMetricType, Me, sFirstMetadata)

              'frmReportBuilderSubmit.getRevisedMetricScore(sFirstMetadata, getThisMaxScore(sThisMetricType, Me, sFirstMetadata), Me, sThisMetricType, isClientSatisfaction, getCurrentEvalType())

        Else

          dMaxSecondaryScoreSubtotal = 5

'frmReportBuilderSubmit.getRevisedEsatScore(sSecondMetadata)

          setEsatFound (True)

        End If

      End If

      If LCase(sFirstMetadata) = "n/a" And getMetricSection(sThisMetricType) = "Business Expectations" Then

        setBusinessExpectationNaFound (True)

      End If

      If getMetricSection(sThisMetricType) = "Verification" Or sThisMetricType = "Verification" Then

        setVerificationFound (True)

        If getMetricSection(sThisMetricType) = "Verification" Then

          sVerificationAnswer = sFirstMetadata

        Else

          sVerificationAnswer = sSecondMetadata

        End If

        If LCase(LTrim(RTrim(sVerificationAnswer))) = "yes" Then

          Call setVerificationStatus(True)

        Else

          Call setVerificationStatus(False)

        End If

      End If

    End If

    Call oThisEval.setSecondaryScore(getSecondaryScore())

    Call insertProcessedRow(oThisEval)

  Next i

End Sub

 

Public Sub processBusinessExpectationsNa()

    If Not bMissingBusinessExpectationsNaProcessed And Not bBusinessExpectationsNaFound Then

        If oEvalStats.etype = "Written" Then

            oEvalStats.esat = oEvalStats.esat + 1 ' Challenge Processed Correctly

            dMaxSecondaryScoreSubtotal = dMaxSecondaryScoreSubtotal + 1

        End If

    End If

    bMissingBusinessExpectationsNaProcessed = True

End Sub

 

Public Sub setVerificationStatus(verificationStatus As Boolean)

  oEvalStats.everification = verificationStatus

End Sub

 

Public Sub processDefaultVerification()

  Dim sFirstMetadata As String

  Dim sThisMetricType As String

  Dim sWorkingText As String

  Dim sRawLeft As String

  Dim sRawRight As String

  Dim oThisEval As EvaluationComment

  Set oThisEval = New EvaluationComment

  ' Was this interaction free of authentication errors?

  sRawLeft = "Was this interaction free of authentication errors?"

  sRawRight = "||Yes||DEFAULT"

  'Call setRawComment(sRawLeft, sRawRight)

  oThisEval.setAgent (getAgentName)

  oThisEval.setEvalType (getCurrentEvalType)

  oThisEval.setTimeStamp (getCurrentTimeStamp)

  oThisEval.setPrimaryScore (getCurrentPrimaryScore())

  sWorkingText = getCommentPreMetadataProcessing(sRawRight)

  sFirstMetadata = getMetadataFirst(sWorkingText, sRawLeft)

  sThisMetricType = getMetricTypeRevised(sRawLeft, sFirstMetadata)

  oThisEval.setComment (getStrippedComment(sWorkingText))

  Call oThisEval.setMaxScore(getThisMaxScore(sThisMetricType, Me, sFirstMetadata))

  Call oThisEval.setMetricScore(frmReportBuilderSubmit.getRevisedMetricScore(sFirstMetadata, oThisEval.getMaxScore, Me, sThisMetricType, isClientSatisfaction, getCurrentEvalType))

  Call oThisEval.setScoreLabel(frmReportBuilderSubmit.getRevisedMetricLabel(oThisEval.getMetricScore, oThisEval.getMaxScore, sThisMetricType))

  oThisEval.setMetricType (sThisMetricType)

  setVerificationFound (True)

  Call setVerificationStatus(True)

  Call oThisEval.setSecondaryScore(getSecondaryScore())

  Call insertProcessedRow(oThisEval)

' Call processRaw

' Call setSortOrder

End Sub

 

Public Sub processMissingSecondaryScore()

    bSecondaryScoreValid = False

End Sub

 

Public Function isSecondaryScoreValid() As Boolean

    isSecondaryScoreValid = bSecondaryScoreValid And hasSecondaryScore()

End Function

 

Public Function getSecondaryScore() As Double

    If oEvalStats.esat = 0# Then

        oEvalStats.esat = deriveSecondaryScore(aRawComments)

    End If

    getSecondaryScore = oEvalStats.esat

End Function

 

Private Function getCommentPreMetadataProcessing(ByVal sPreprocessedComment As String)

    If InStr(1, sPreprocessedComment, "||", vbBinaryCompare) = 0 And InStr(1, sPreprocessedComment, ":", vbBinaryCompare) < InStr(1, sPreprocessedComment, Chr(10), vbBinaryCompare) Then

        sPreprocessedComment = Replace(sPreprocessedComment, "Yes:", "||Yes||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Partial:", "||Partial||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "No:", "||No||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Not Likely:", "||Not Likely||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Not likely:", "||Not Likely||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Likely:", "||Likely||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Definitely:", "||Definitely||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly Disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly Agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "strongly disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "strongly agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Agree:", "||ESAT||3.75||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Neutral:", "||ESAT||2.5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Disagree:", "||ESAT||1.25||", 1, 1)

    ElseIf InStr(1, sPreprocessedComment, "||", vbBinaryCompare) = 0 And InStr(1, sPreprocessedComment, ":", vbBinaryCompare) < InStr(InStr(1, sPreprocessedComment, " ", vbBinaryCompare) + 1, sPreprocessedComment, " ", vbBinaryCompare) Then

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly Disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly Agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Strongly agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "strongly disagree:", "||ESAT||0||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "strongly agree:", "||ESAT||5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Agree:", "||ESAT||3.75||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Neutral:", "||ESAT||2.5||", 1, 1)

        sPreprocessedComment = Replace(sPreprocessedComment, "Disagree:", "||ESAT||1.25||", 1, 1)

    End If

    ' Resolve common error of more than two vertical pipes

    getCommentPreMetadataProcessing = Replace(sPreprocessedComment, "|||", "||")

End Function

' Get first argument

Public Function getMetadataFirst(ByVal sComment As String, ByVal sRawMetricType As String)

    Dim meta_data As String

    If InStr(1, sComment, "||", vbBinaryCompare) > 0 And InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||", vbBinaryCompare) > 0 Then

        meta_data = LCase(LTrim(RTrim(Mid(sComment, InStr(1, sComment, "||", vbBinaryCompare) + 2, InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||", vbBinaryCompare) - (InStr(1, sComment, "||", vbBinaryCompare) + 2)))))

    End If

    getMetadataFirst = meta_data

End Function

' Get second argument

Public Function getMetadataSecond(ByVal sComment As String, ByVal sRawMetricType As String) As String

    Dim revised_mtype As String

    Dim second_arg As String

    second_arg = ""

    revised_mtype = getMetricTypeRevised(sRawMetricType, getMetadataFirst(sComment, sRawMetricType))

    ' Verification sComment processing

    If revised_mtype = "Verification" Or revised_mtype = "ESAT" Then

        If InStr(InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||") + 2, sComment, "||") > 0 Then

            second_arg = StrConv(LTrim(RTrim(Mid(sComment, InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||") + 2, InStr(InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||") + 2, sComment, "||") - (InStr(InStr(1, sComment, "||", vbBinaryCompare) + 2, sComment, "||") + 2)))), vbProperCase)

        End If

    End If

    second_arg = Replace(second_arg, "  ", " ")

    getMetadataSecond = second_arg

End Function

 

Public Function isMetadataBadFormat(Optional ByVal sComment As String, Optional ByVal sRawMetricType As String) As Boolean

  Dim aInputDuple

  'Dim oRawElement As Variant

  If sComment = "" And sRawMetricType = "" And (Not Not aRawComments) <> 0 Then

    If iEvaluationIndex > UBound(aRawComments) Or iEvaluationIndex < LBound(aRawComments) Then

      isMetadataBadFormat = False

      Exit Function

    End If

    'oRawElement = aRawComments(iEvaluationIndex)

    aInputDuple = aRawComments(iEvaluationIndex) 'oRawElement

    If UBound(aInputDuple) < 0 Then

      isMetadataBadFormat = False

    Else

      isMetadataBadFormat = Len(getMetadataFirst(aInputDuple(UBound(aInputDuple)), aInputDuple(LBound(aInputDuple)))) < 1

    End If

  ElseIf sComment = "" And sRawMetricType = "" Then

    isMetadataBadFormat = False

  Else

    isMetadataBadFormat = Len(getMetadataFirst(sComment, sRawMetricType)) < 1

  End If

End Function

 

Public Sub addCallHandlingMax()

    Dim temp_agent_offset As Integer

    Dim temp_sm_offset As Integer

    Dim m_score As Double

    Dim m_max As Double

    If Not iCallHandlingCt = 0 Then

        m_max = (1 / iCallHandlingCt)

        If hold_transfer_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - hold_transfer_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - hold_transfer_offset)

            Call frmReportBuilderSubmit.addMetricMax(hold_transfer_offset, m_max, primary_output_tab_n)

            m_score = frmReportBuilderSubmit.getMetricScore(hold_transfer_offset, primary_output_tab_n)

            If frmReportBuilderSubmit.getMetricScore(hold_transfer_offset, primary_output_tab_n) = 1 Then

                Call frmReportBuilderSubmit.addMetricScore(hold_transfer_offset, m_score * m_max, primary_output_tab_n)

                Call frmReportBuilderSubmit.addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call frmReportBuilderSubmit.addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call frmReportBuilderSubmit.addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call frmReportBuilderSubmit.addMetricPercent(hold_transfer_offset, m_score, primary_output_tab_n)

            Call frmReportBuilderSubmit.addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call frmReportBuilderSubmit.addMetricMax(temp_sm_offset, m_max, current_sm)

                Call frmReportBuilderSubmit.addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If call_log_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - call_log_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - call_log_offset)

            Call frmReportBuilderSubmit.addMetricMax(call_log_offset, m_max, primary_output_tab_n)

            m_score = frmReportBuilderSubmit.getMetricScore(call_log_offset, primary_output_tab_n)

            If frmReportBuilderSubmit.getMetricScore(call_log_offset, primary_output_tab_n) = 1 Then

                Call frmReportBuilderSubmit.addMetricScore(call_log_offset, m_score * m_max, primary_output_tab_n)

                Call frmReportBuilderSubmit.addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call frmReportBuilderSubmit.addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call frmReportBuilderSubmit.addMetricMax(temp_agent_offset, m_max, current_agent)

            

            Call frmReportBuilderSubmit.addMetricPercent(call_log_offset, m_score, primary_output_tab_n)

            Call frmReportBuilderSubmit.addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call frmReportBuilderSubmit.addMetricMax(temp_sm_offset, m_max, current_sm)

                Call frmReportBuilderSubmit.addMetricPercent(temp_sm_offset, m_score / m_max, current_sm)

            End If

        End If

        If added_updated_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - added_updated_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - added_updated_offset)

            Call frmReportBuilderSubmit.addMetricMax(added_updated_offset, m_max, primary_output_tab_n)

            m_score = frmReportBuilderSubmit.getMetricScore(added_updated_offset, primary_output_tab_n)

            If frmReportBuilderSubmit.getMetricScore(added_updated_offset, primary_output_tab_n) = 1 Then

                Call frmReportBuilderSubmit.addMetricScore(added_updated_offset, m_score * m_max, primary_output_tab_n)

                Call frmReportBuilderSubmit.addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call frmReportBuilderSubmit.addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call frmReportBuilderSubmit.addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call frmReportBuilderSubmit.addMetricPercent(added_updated_offset, m_score, primary_output_tab_n)

            Call frmReportBuilderSubmit.addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call frmReportBuilderSubmit.addMetricMax(temp_sm_offset, m_max, current_sm)

                Call frmReportBuilderSubmit.addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If survey_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - survey_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - survey_offset)

            Call frmReportBuilderSubmit.addMetricMax(survey_offset, m_max, primary_output_tab_n)

            m_score = frmReportBuilderSubmit.getMetricScore(survey_offset, primary_output_tab_n)

            If frmReportBuilderSubmit.getMetricScore(survey_offset, primary_output_tab_n) = 1 Then

                Call frmReportBuilderSubmit.addMetricScore(survey_offset, m_score * m_max, primary_output_tab_n)

                Call frmReportBuilderSubmit.addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

            End If

            Call frmReportBuilderSubmit.addMetricMax(temp_agent_offset, m_max, current_agent)

            Call frmReportBuilderSubmit.addMetricPercent(survey_offset, m_score, primary_output_tab_n)

            Call frmReportBuilderSubmit.addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call frmReportBuilderSubmit.addMetricMax(temp_sm_offset, m_max, current_sm)

                Call frmReportBuilderSubmit.addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If callback_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - callback_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - callback_offset)

            Call frmReportBuilderSubmit.addMetricMax(callback_offset, m_max, primary_output_tab_n)

            m_score = frmReportBuilderSubmit.getMetricScore(callback_offset, primary_output_tab_n)

           If frmReportBuilderSubmit.getMetricScore(callback_offset, primary_output_tab_n) = 1 Then

                Call frmReportBuilderSubmit.addMetricScore(callback_offset, m_score * m_max, primary_output_tab_n)

                Call frmReportBuilderSubmit.addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call frmReportBuilderSubmit.addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call frmReportBuilderSubmit.addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call frmReportBuilderSubmit.addMetricPercent(callback_offset, frmReportBuilderSubmit.getMetricScore(callback_offset, primary_output_tab_n) / frmReportBuilderSubmit.getMetricMax(callback_offset, primary_output_tab_n), primary_output_tab_n)

            Call frmReportBuilderSubmit.addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call frmReportBuilderSubmit.addMetricMax(temp_sm_offset, m_max, current_sm)

                Call frmReportBuilderSubmit.addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

    End If

    iCallHandlingCt = 0

    callback_offset = 0

    survey_offset = 0

    added_updated_offset = 0

    hold_transfer_offset = 0

    call_log_offset = 0

 

End Sub

 

' aAllRaw is an Array of arrays with index zero (0) raw metric label, and one (1) raw comment

Private Function deriveSecondaryScore(ByRef aAllRaw() As Variant) As Double

    Dim iRawIndex As Integer

    Dim sEvalType As String

    Dim sMetricSection As String

    Dim dSecondarySubtotal As Double

    Dim dMetricScore As Double

    Dim sRawMetricLabel As String

    Dim sRawComment As String

    Dim sMetricType As String

    Dim oFirstMetadata As Variant

    For iRawIndex = LBound(aAllRaw) To UBound(aAllRaw)

        sEvalType = getCurrentEvalType()

        oFirstMetadata = getMetadataFirst(aAllRaw(iRawIndex)(1), aAllRaw(iRawIndex)(0))

        sMetricType = getMetricTypeRevised(aAllRaw(iRawIndex)(0), oFirstMetadata)

        sMetricSection = getMetricSection(sMetricType)

        If sMetricSection = "Business Expectations" Or sMetricSection = "ESAT" Then

            dMetricScore = dMetricScore + frmReportBuilderSubmit.getRevisedMetricScore(oFirstMetadata, getThisMaxScore(sMetricType, Me, oFirstMetadata), Me, sMetricType, isClientSatisfaction, getCurrentEvalType)

        End If

    Next iRawIndex

    deriveSecondaryScore = dMetricScore

End Function

 

Private Function getStrippedComment(ByVal sCommentWithMetadata As String) As String

    getStrippedComment = frmReportBuilderSubmit.stripLeadTrailNewline(LTrim(RTrim(Mid(sCommentWithMetadata, InStrRev(sCommentWithMetadata, "||", Compare:=vbBinaryCompare) + 2, Len(sCommentWithMetadata) - (InStrRev(sCommentWithMetadata, "||", Compare:=vbBinaryCompare) + 1)))))

End Function

 

Public Function hasSecondaryScore()

  Dim bAnswer As Boolean

  bAnswer = False

  If isEsatFound Then

    bAnswer = True

  ElseIf dMaxSecondaryScoreSubtotal = dExpectedSecondaryMax Then

        bAnswer = True

  ElseIf Not isBusinessExpectationNaFound Then

    Call processBusinessExpectationsNa

    If dMaxSecondaryScoreSubtotal = dExpectedSecondaryMax Then

      bAnswer = True

    Else

      bAnswer = False

    End If

  End If

  hasSecondaryScore = bAnswer

End Function

 

Public Function getThisMaxScore(ByVal metric_type As String, oCurrentEval As Evaluation, Optional ByVal meta_data As Variant) As Double

    If isCrd And getMetricSection(metric_type) = "Business Expectations" And Not getCurrentEvalType = "Written" Then ' CRD Non-Case

        Select Case metric_type

            Case "Appropriate Greeting", "Appropriate Closing", "Call Log Entered", "Survey", "Correct Resources", "Call Log Details"

                getThisMaxScore = 0.5

            Case "Business Processes"

                getThisMaxScore = 2

        End Select

    ElseIf (isNna Or isMssnMisc) And metric_type = "Business Processes" Then

        getThisMaxScore = 2

    ElseIf (isNna Or isMssnMisc) And metric_type = "Appropriate Closing" Then

        getThisMaxScore = 0.5

    ElseIf isCrd And getMetricSection(metric_type) = "Business Expectations" And getCurrentEvalType = "Written" Then 'CRD Case

      If metric_type = "Correct Resources" Or metric_type = "Service Level Met" Or metric_type = "Challenge Processed Correctly" Or metric_type = "Business Processes" Then

        getThisMaxScore = 1

      ElseIf metric_type = "Appropriate Greeting" Or metric_type = "Appropriate Closing" Then

        getThisMaxScore = 0.5

      End If

    'ElseIf 'getCurrentEvalType = "Written" And

    ElseIf (metric_type = "Grammar Error Free" Or metric_type = "Forward Thinking" Or metric_type = "Appropriate Closing" Or metric_type = "Challenge Processed Correctly" _

        Or metric_type = "Business Processes") And oCurrentEval.getCurrentEvalType = "Written" Then

      Select Case metric_type

        Case "Forward Thinking", "Grammar Error Free"

          getThisMaxScore = 0.25

        Case "Appropriate Closing"

          getThisMaxScore = 0.5

        Case "Challenge Processed Correctly"

          getThisMaxScore = 1

        Case "Business Processes"

          If isCrd Then

            getThisMaxScore = 1

          Else

            getThisMaxScore = 2

          End If

        Case "Service Level Met"

          getThisMaxScore = 1

      End Select

    Else

        Select Case metric_type

            Case "Comment"

                getThisMaxScore = 0

            Case "Evaluator Satisfaction"

                getThisMaxScore = 5

            Case "ESAT"

                getThisMaxScore = 5

            Case "Verification"

                getThisMaxScore = 1

            Case "Accurate Information", "Expectations", "Appropriate Greeting"

                getThisMaxScore = 0.5

            Case "Process / Procedures"

                getThisMaxScore = 2

            Case "Opening / Farewell"

                getThisMaxScore = 0.25

            Case "Actively Listened"

                getThisMaxScore = 0.25

           Case "Controlled Call"

                getThisMaxScore = 0.25

            Case "Clear / Confident"

                getThisMaxScore = 0.25

            Case "Business Comment"

                getThisMaxScore = 0

            Case "Hold Comment"

                getThisMaxScore = 0

            Case "Hold / Transfer"

                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

                    iCallHandlingCt = iCallHandlingCt + 1

                End If

                getThisMaxScore = 0

            Case "Call Log"

                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

                    iCallHandlingCt = iCallHandlingCt + 1

                End If

                getThisMaxScore = 0

            Case "Added / Updated"

                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

                    iCallHandlingCt = iCallHandlingCt + 1

                End If

                getThisMaxScore = 0

            Case "Survey"

                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

                    iCallHandlingCt = iCallHandlingCt + 1

                End If

                getThisMaxScore = 0

            Case "Callback"

                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

                    iCallHandlingCt = iCallHandlingCt + 1

                End If

                getThisMaxScore = 0

            Case "Accuracy / Completeness"

                getThisMaxScore = 2

            Case "Complete Expectations", "Timely Resolution"

                getThisMaxScore = 0.75

            Case "World-Class Service", "Correct Resources", "Appropriate Closing"

                getThisMaxScore = 1

            Case "Hold Experience", "Transfer Experience"

                getThisMaxScore = 0.25

            Case "Business Processes"

                getThisMaxScore = 2.5

            Case "Promote Branch Self-Service"

                getThisMaxScore = 2

        End Select

    End If

End Function

 

Public Function getMetricSection(metric_type_revised As String) As String

    Select Case LCase(LTrim(RTrim(metric_type_revised)))

        Case "accuracy / completeness"

            getMetricSection = "Meaningful Solutions"

        Case "complete expectations"

            getMetricSection = "Meaningful Solutions"

        Case "timely resolution"

            getMetricSection = "Meaningful Solutions"

        Case "world-class service"

            getMetricSection = "Servicing Skills"

        Case "hold experience"

            getMetricSection = "Servicing Skills"

        Case "transfer experience"

            getMetricSection = "Servicing Skills"

        Case "grammar error free"

            getMetricSection = "Servicing Skills"

        Case "forward thinking"

            getMetricSection = "Servicing Skills"

        Case "appropriate greeting"

            getMetricSection = "Business Expectations"

        Case "correct resources"

            getMetricSection = "Business Expectations"

        Case "appropriate closing"

            getMetricSection = "Business Expectations"

        Case "business processes"

            getMetricSection = "Business Expectations"

        Case "call log entered"

            getMetricSection = "Business Expectations"

        Case "call log details"

            getMetricSection = "Business Expectations"

        Case "service level met"

            getMetricSection = "Business Expectations"

        Case "survey"

            If isCrd Then

                getMetricSection = "Business Expectations"

            Else

                getMetricSection = "Business Expectations"

            End If

        Case "challenge processed correctly"

            getMetricSection = "Business Expectations"

        Case "promote branch self-service"

            getMetricSection = "Business Expectations"

        Case "esat"

            getMetricSection = "Evaluator Satisfaction"

        Case "hold comment"

            getMetricSection = "Comment"

        Case "business comment"

            getMetricSection = "Comment"

        Case "comment"

            getMetricSection = "Comment"

        Case "actively listened"

            getMetricSection = "Client Experience"

        Case "controlled call"

            getMetricSection = "Client Experience"

        Case "clear / confident"

            getMetricSection = "Client Experience"

        Case "process / procedures"

            getMetricSection = "Procedural Accuracy"

        Case "call log"

            getMetricSection = "Call Handling"

        Case "opening / farewell"

            getMetricSection = "Client Experience"

        Case "accurate information"

            getMetricSection = "Procedural Accuracy"

        Case "expectations"

            getMetricSection = "Procedural Accuracy"

        Case "callback"

            getMetricSection = "Call Handling"

        Case "added / updated"

            getMetricSection = "Call Handling"

        Case "survey"

            getMetricSection = "Call Handling"

        Case "hold / transfer"

            getMetricSection = "Call Handling"

        Case "verification"

            If isMssnMisc Then

                getMetricSection = "Business Expectations"

            Else

                getMetricSection = "Verification"

            End If

    End Select

End Function

 

Public Function getMetricTypeRevised(ByVal metric_name As Variant, ByVal meta_data As String) As String

    Select Case LCase(LTrim(RTrim(metric_name)))

        Case "comment"

            If LCase(LTrim(RTrim(meta_data))) = "evaluator satisfaction" Or LCase(LTrim(RTrim(meta_data))) = "esat" Then

                getMetricTypeRevised = "ESAT"

            ElseIf LCase(LTrim(RTrim(meta_data))) = "hold comment" Then

                getMetricTypeRevised = "Hold Comment"

            ElseIf LCase(LTrim(RTrim(meta_data))) = "verification" Then

                getMetricTypeRevised = "Verification"

            ElseIf LCase(LTrim(RTrim(meta_data))) = "business comment" Then

                getMetricTypeRevised = "Business Comment"

            ElseIf Len(frmReportBuilderSubmit.SuppliedEvalType(meta_data)) > 0 Then

                getMetricTypeRevised = "Comment"

            Else

                getMetricTypeRevised = ""

            End If

        Case "did the agent provide the complete and correct answer?"

            getMetricTypeRevised = "Accuracy / Completeness"

        Case "was the complete and correct answer provided?"

            getMetricTypeRevised = "Accuracy / Completeness"

        Case "were the next steps communicated completely?"

            getMetricTypeRevised = "Complete Expectations"

        Case "were next steps communicated completely?"

            getMetricTypeRevised = "Complete Expectations"

        Case "was the clients concern resolved in a timely manner?"

            getMetricTypeRevised = "Timely Resolution"

        Case "was the case resolved in a timely manner?"

            getMetricTypeRevised = "Timely Resolution"

        Case "was world class service demonstrated on this interaction?"

            getMetricTypeRevised = "World-Class Service"

        Case "was the case handled professionally?"

            getMetricTypeRevised = "World-Class Service"

        Case "was response free from grammatical errors?"

            getMetricTypeRevised = "Grammar Error Free"

        Case "did agent show forward thinking for any additional questions that may arise?"

            getMetricTypeRevised = "Forward Thinking"

        Case "did the agent create a satisfactory hold experience?"

            getMetricTypeRevised = "Hold Experience"

        Case "did the agent create a satisfactory transfer experience?"

            getMetricTypeRevised = "Transfer Experience"

        Case "was the appropriate greeting used?"

            getMetricTypeRevised = "Appropriate Greeting"

        Case "was this interaction free of authentication errors?"

            getMetricTypeRevised = "Verification"

        Case "were correct resources used?"

            getMetricTypeRevised = "Correct Resources"

        Case "was the appropriate closing used?"

            getMetricTypeRevised = "Appropriate Closing"

        Case "were all business guidelines and areas of focus addressed?"

            getMetricTypeRevised = "Business Processes"

        Case "were all business requirements accurately completed?"

            getMetricTypeRevised = "Business Processes"

        Case "actively listened to client and correctly identified the root cause of the call"

            getMetricTypeRevised = "Actively Listened"

        Case "appropriately controlled the call"

            getMetricTypeRevised = "Controlled Call"

        Case "communicated in a clear and confident manner"

            getMetricTypeRevised = "Clear / Confident"

        Case "followed correct processes and procedures"

            getMetricTypeRevised = "Process / Procedures"

        Case "logged call correctly and added necessary notes"

            getMetricTypeRevised = "Call Log"

        Case "provided a warm opening and a fond farewell"

            getMetricTypeRevised = "Opening / Farewell"

        Case "provided accurate information"

            getMetricTypeRevised = "Accurate Information"

        Case "set appropriate expectation with client"

            getMetricTypeRevised = "Expectations"

        Case "what is the likelihood that the caller will need to call again due to the agent's handling of the interaction?"

            getMetricTypeRevised = "Callback"

        Case "added or updated all required information"

            getMetricTypeRevised = "Added / Updated"

        Case "offered survey at the end of the call"

            getMetricTypeRevised = "Survey"

        Case "followed appropriate hold / dial procedure"

            getMetricTypeRevised = "Hold / Transfer"

        Case "was call log entered correctly?"

            getMetricTypeRevised = "Call Log Entered"

        Case "were all details provided in call log?"

            getMetricTypeRevised = "Call Log Details"

        Case "was a survey offered?"

            getMetricTypeRevised = "Survey"

        Case "did the agent promote branch self-service?"

            getMetricTypeRevised = "Promote Branch Self-Service"

        Case "if challenged, had the processing been done correctly?"

            If LCase(LTrim(RTrim(meta_data))) = "yes" Or LCase(LTrim(RTrim(meta_data))) = "no" Or LCase(LTrim(RTrim(meta_data))) = "n/a" Then

                setBusinessExpectationNaFound (True)

            End If

            getMetricTypeRevised = "Challenge Processed Correctly"

        Case "did the agents response meet sl?"

            getMetricTypeRevised = "Service Level Met"

        Case Else

            getMetricTypeRevised = ""

    End Select

   

End Function

 

' move index forward

Public Sub moveIndexForward()

    iEvaluationIndex = iEvaluationIndex + 1

End Sub

 

Public Function isIndexValid()

  If (Not Not aEvaluationComments) <> 0 Then

    isIndexValid = (UBound(aEvaluationComments) - LBound(aEvaluationComments) + 1) > iEvaluationIndex

  Else

    isIndexValid = False

  End If

End Function

 

Public Sub resetEvaluationIterator()

    iEvaluationIndex = 0

End Sub

 

Public Function getRawCommentCount() As Integer

    getRawCommentCount = (UBound(aRawComments) - LBound(aRawComments) + 1)

End Function

 

 

Public Function isRawCommentIndexPositionValid() As Boolean

    isRawCommentIndexPositionValid = iRawCommentIterator < getRawCommentCount

End Function

 

 

Public Sub resetRawCommentIndexPosition()

    iRawCommentIterator = 0

End Sub

 

Public Sub moveRawCommentIndexPosition()

  iRawCommentIterator = iRawCommentIterator + 1

End Sub

 

Public Function getAgentName()

    getAgentName = oEvalStats.Agent

End Function

 

Public Function getCurrentMetricType()

  Dim oThisComment As EvaluationComment

  Set oThisComment = aEvaluationComments(iEvaluationIndex)

  getCurrentMetricType = oThisComment.getMetricType()

End Function

 

Public Function getCurrentOriginalComment()

  Dim oThisComment As EvaluationComment

  Set oThisComment = aEvaluationComments(iEvaluationIndex)

  getCurrentOriginalComment = oThisComment.getOriginalComment()

End Function

 

Public Function getCurrentComment()

  Dim oThisComment As EvaluationComment

  If (Not Not aEvaluationComments) = 0 Then

    getCurrentComment = ""

  ElseIf UBound(aEvaluationComments) < 0 Then

    getCurrentComment = ""

  Else

    Set oThisComment = aEvaluationComments(iEvaluationIndex)

    getCurrentComment = oThisComment.getComment()

  End If

End Function

 

Public Function getCurrentEvalType()

    getCurrentEvalType = oEvalStats.etype

End Function

 

Public Function getCurrentScoreLabel()

    Dim oThisComment As EvaluationComment

    Set oThisComment = aEvaluationComments(iEvaluationIndex)

    getCurrentScoreLabel = oThisComment.getScoreLabel()

End Function

 

Public Function getCurrentMetricScore()

    Dim oThisComment As EvaluationComment

    Set oThisComment = aEvaluationComments(iEvaluationIndex)

    getCurrentMetricScore = oThisComment.getMetricScore()

End Function

 

Public Function getCurrentMaxScore()

    Dim oThisComment As EvaluationComment

    Set oThisComment = aEvaluationComments(iEvaluationIndex)

    getCurrentMaxScore = oThisComment.getMaxScore()

End Function

 

Public Function getCurrentMetricPercentage() As Variant

    Dim oThisComment As EvaluationComment

    Set oThisComment = aEvaluationComments(iEvaluationIndex)

    getCurrentMetricPercentage = oThisComment.getMetricPercentage()

End Function

 

Public Function getCurrentPrimaryScore()

    getCurrentPrimaryScore = oEvalStats.procedural

End Function

 

Public Function getCurrentSecondaryScore()

    getCurrentSecondaryScore = oEvalStats.esat

End Function

 

Public Function getCurrentTimeStamp()

  getCurrentTimeStamp = oEvalStats.edate

End Function

 

Public Function getGeneralStatistics() As EvalProcEsatTypeDate

  Set getGeneralStatistics = oEvalStats

End Function
