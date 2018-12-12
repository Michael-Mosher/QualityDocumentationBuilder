Option Explicit

 

    ' current agent dates

    Private temp_eval_arr() As Date

    ' current agent metrics

    Private temp_arr_clct() As Collection

    Private error_lock As Boolean

    ' metric_name => comment

    Private this_clct As Collection

    Private sm_proc_scores As Collection ' of arrays of Double

    Private sm_esat_scores As Collection ' of arrays of Double

    Private sp_sm_collection As Collection ' SP name is key pointing to SM name

    Private first_eval As Boolean

    Private oAgentsEvalScores As Collection ' agent names point to array() EvalProcEsatTypeDate

Private section_metric_label_qty As MetricSctnMetricLblScoreLblQty

Private agents_section_scores As AgentsMetricSectionScores

   

    Private temp_coll As Collection

    Private metric_types() As String

   

    Private this_metric_name As String

    Private this_eval_metric_qty As Integer

    Private this_agent_eval_qty As Integer

    Private current_agent As String

    Private previous_agent As String

    Private eval_date_i As Integer

    Private eval_clct_i As Integer

    Private proc_score_clct_i As Integer

    Private esat_score_clct_i As Integer

    Private metric_score_i As Integer

    Private metric_max_i As Integer

    Private metric_pct_i As Integer

    Private eval_comment_i As Integer

    Private esat_qty_deficiency As Integer

    Private eval_qty_max As Integer

    Private agent_qty As Integer

   

    Private output_book As Workbook

    Private original_data_wb As Workbook

    Private output_row_offset As Integer

    Private first_clmn_label As String

    Private last_clmn_label As String

    Private sm_sp_sheet_name As String

    Private qs_nr_evaluation_scores As String

    Private qs_nr_verification_summary As String

    Private qs_nr_verbal_evals As String

    Private qs_nr_business_evals As String

    Private qs_nr_certification_evals As String

    Private qs_nr_survey_evals As String

    Private qs_nr_negative_evals As String

    Private qs_nr_written_evals As String

    Private qs_nr_metric_lbl_section_qty As String

    Private qs_nr_metric_lbl_section_pct As String

    Private qs_nr_metric_lbl_metric_qty As String

    Private qs_nr_section_avg As String

    Private qs_nr_esat_qty As String

   

    Private garbage_row_offset As Integer

    Private current_agent_offset As Integer

    Private current_sm_offset As Integer

    Private first_agent As Boolean

    Private first_time As Boolean

    Private first_metric As Boolean

    Private verification_comment_supplied As Boolean

    Private duplicate_evaluation As Boolean

    Private is_agent As Boolean

    Private primary_output_tab_n As String

    Private call_handling_count As Integer

    Private hold_transfer_offset As Integer

    Private call_log_offset As Integer

    Private added_updated_offset As Integer

    Private survey_offset As Integer

    Private callback_offset As Integer

    Private sm_known As Boolean

    Private is_qp As Boolean

    Private certification_eval As Boolean

    Private business_expectation_na_resolved As Boolean

    Private has_esat As Boolean

    Private current_sm As String

    Private first_comment_row As Integer

    Private first_header_row As Integer

    Private avg_row_offset As Integer

    ' Quality Ranking variables

    Private eval_month As String

    Private eval_year As Integer

    Private qr_proc_clmn_n As String

    Private qr_proc_named_range_n As String

    Private qr_esat_named_range_n As String

    Private qr_esat_range_n As String

    Private update_qr As Boolean

    Private created_new_month As Boolean

    Private truncate_numbers As Boolean

    Private leftmost_rightblockclmn_qr As String

    Private qr_primarykey_clmn_n As String

    Private month_proc_score_clmn_n As String

    Private month_esat_score_clmn_n As String

    Private qr_proc_last_clmn_n As String

    Private qr_esat_last_clmn_n As String

   

    Private is_client_experience As Boolean

 

Private collection_ag As Collection

Private garbage_report As Collection

Private garbage_notes As Collection

 

 

 

Private Sub btnSubmit_Click()

    Dim r As Range

    Dim sort_key1 As Range

    Dim sort_key2 As Range

    Dim metric_type As Variant

    Dim comment As Variant

    Dim temp_cell_value As Variant

    Dim evals As Variant

    Dim arr_of_clcts As Variant

    ' k:metric_type_revised, v:array [0:comment, 1:metric_score, 2:metric_max, 3:metric_pct]

    Dim metric_notesnstats As Collection

    Dim current_agent_dates() As Date

    Dim this_sm_proc() As Double

    Dim this_sm_esat() As Double

    Dim current_eval As Date

    Dim sheets_i As Integer

    Dim rows_i2 As Long

    Dim proc_sub_total As Double

    Dim esat_sub_total As Double

    Dim i As Integer

    Dim eval_year_string As String

    Dim temp_eval As EvalProcEsatTypeDate

    Dim temp_adate() As Date

    Dim temp_bdate() As Date

    Dim aMetricLabelCount() As Integer

    Dim temp_evalarray() As EvalProcEsatTypeDate

    Dim oTheEvaluations As EvaluationCollection

   

    Dim logic_switch_token As Boolean

   

    Dim temp_coll2 As Collection

    Dim temp_name As String

    Dim oTempEvaluation As Evaluation

   

    Const eval_date_i = 0

    Const eval_clct_i = 3

    Const proc_score_clct_i = 1

    Const esat_score_clct_i = 2

    Const metric_score_i = 1

    Const metric_max_i = 2

    Const metric_pct_i = 3

    Const eval_comment_i = 0

    this_eval_metric_qty = 0

    this_agent_eval_qty = 0

    first_header_row = 1

    first_comment_row = first_header_row + 1

    output_row_offset = first_comment_row

    garbage_row_offset = first_comment_row

    current_sm_offset = first_comment_row

    current_agent_offset = first_comment_row

    hold_transfer_offset = 0

    call_log_offset = 0

    added_updated_offset = 0

    survey_offset = 0

    callback_offset = 0

    avg_row_offset = 3

    esat_qty_deficiency = 0

    eval_qty_max = 0

    agent_qty = 0

    first_agent = True

    first_metric = True

    first_eval = True

    sm_known = False

    certification_eval = False

    logic_switch_token = False

    business_expectation_na_resolved = False

    Set oTempEvaluation = New Evaluation

    oTempEvaluation.setClientSatisfaction (cbxIsClientExperience)

    oTempEvaluation.setIsCrd (cbxCRD)

    oTempEvaluation.setIsMssnMisc (cbxMSSN)

    oTempEvaluation.setIsNna (cbxNNA)

   

    Set output_book = ActiveWorkbook 'Workbooks.Add

    Application.ScreenUpdating = False

    verification_comment_supplied = False

    update_qr = cbxQualityRanking.Value

    truncate_numbers = cbxTruncatedNumbers

    is_client_experience = cbxIsClientExperience.Value

    If cbxFullMonth.Value Then

        primary_output_tab_n = sltEvalMonth.Value & ", " & sltEvalYear.Value

        eval_month = sltEvalMonth.Value

        eval_year = sltEvalYear.Value

    ElseIf Not cbxFullMonth.Value Then

        primary_output_tab_n = Replace(dtpBeginDate.Value, "/", "-") & " - " & Replace(dtpEndDate.Value, "/", "-")

    End If

    If update_qr Then

        qr_proc_named_range_n = "tblQualityRanking"

        qr_esat_named_range_n = "tblESATRanking"

        qr_primarykey_clmn_n = "SP"

        eval_year_string = eval_year & ""

        leftmost_rightblockclmn_qr = "Previous Month"

        qr_proc_last_clmn_n = "Procedural Pct. Chg."

        qr_esat_last_clmn_n = "ESAT Pct. Chg."

        If Len(eval_month) < 5 Then

            qr_proc_clmn_n = eval_month & " " & Right(eval_year_string, 2) & " Procedural"

        Else

            qr_proc_clmn_n = Mid(eval_month, 1, 3) & ". " & Right(eval_year_string, 2) & " Procedural"

        End If

        If isTableColumnName(qr_proc_clmn_n, qr_proc_named_range_n, "Quality Ranking", output_book) Then

            update_qr = False

        End If

       

        With output_book.Worksheets("Quality Ranking")

            With .ListObjects(qr_proc_named_range_n)

                Set sort_key1 = .ListColumns("SP").Range

                With .Sort

                    .SortFields.Clear

                    .Header = xlYes

                    '.Orientation = XlTopBottom

                    .Apply

                    .SortFields.Clear

                    .SortFields.Add key:=sort_key1, SortOn:=xlSortOnValues, ORDER:=xlAscending, DataOption:=xlSortNormal

                    .Header = xlYes

                    .Apply

                End With

            End With

            With .ListObjects(qr_esat_named_range_n)

                Set sort_key1 = .ListColumns("SP").Range

                With .Sort

                    .SortFields.Clear

                    .Header = xlYes

                    '.Orientation = XlTopBottom

                    .Apply

                    .SortFields.Clear

                    .SortFields.Add key:=sort_key1, SortOn:=xlSortOnValues, ORDER:=xlAscending, DataOption:=xlSortNormal

                    .Header = xlYes

                    .Apply

                End With

            End With

        End With

    End If

    first_clmn_label = "Agent"

    last_clmn_label = "Evaluator Satisfaction"

    current_agent = ""

    previous_agent = ""

    sm_sp_sheet_name = "SM-SP"

    duplicate_evaluation = False

    is_agent = True

    has_esat = False

    Set sm_proc_scores = New Collection

    Set sm_esat_scores = New Collection

    Set oAgentsEvalScores = New Collection

    Set sp_sm_collection = New Collection

    'Set sp_evaldate_collection = New Collection

    'Set evals_missing_esat = New Collection

    If is_client_experience Then

        ReDim metric_types(0 To 14)

        If cbxCRD.Value Then

            ReDim metric_types(0 To 17)

        End If

        metric_types(0) = "Comment"

        metric_types(1) = "Verification"

        metric_types(2) = "Accuracy / Completeness"

        metric_types(3) = "Complete Expectations"

        metric_types(4) = "Timely Resolution"

        metric_types(5) = "World-Class Service"

        metric_types(6) = "Hold Experience"

        metric_types(7) = "Transfer Experience"

        metric_types(8) = "Forward Thinking"

        metric_types(9) = "Grammar Error Free"

        metric_types(10) = "Appropriate Greeting"

        If cbxNNA.Value Then

            metric_types(11) = "Promote Branch Self-Service"

        Else

            metric_types(11) = "Correct Resources"

        End If

        metric_types(12) = "Appropriate Closing"

        metric_types(13) = "Business Processes"

        metric_types(14) = "Challenge Processed Correctly"

        If cbxCRD.Value Then

            metric_types(15) = "Call Log Entered"

            metric_types(16) = "Call Log Details"

            metric_types(17) = "Survey"

        End If

        Set section_metric_label_qty = New MetricSctnMetricLblScoreLblQty

        Set agents_section_scores = New AgentsMetricSectionScores

        Set temp_coll = New Collection

       'section_metric_label_qty.Add key:="Meaningful Solutions", Item:=temp_coll

        Set temp_coll = New Collection

       'section_metric_label_qty.Add key:="Servicing Skills", Item:=temp_coll

        Set temp_coll = New Collection

       'section_metric_label_qty.Add key:="Business Expectations", Item:=temp_coll

        Set temp_coll = New Collection

'        For sheets_i = 2 To 11

'            temp_name = ""

'            Set temp_coll = section_metric_label_qty(oTempEvaluation.getMetricSection(metric_types(sheets_i)))

'            section_metric_label_qty.Remove (oTempEvaluation.getMetricSection(metric_types(sheets_i)))

'            If isCollectionKey(metric_types(sheets_i), temp_coll) Then

'                Set temp_coll2 = temp_coll(metric_types(sheets_i))

'                temp_coll.Remove (metric_types(sheets_i))

'            Else

'                Set temp_coll2 = New Collection

'            End If

'            temp_coll2.Add key:="Yes", Item:=0

'            temp_coll2.Add key:="Partial", Item:=0

'            temp_coll2.Add key:="No", Item:=0

'            temp_coll2.Add key:="N/A", Item:=0

'            temp_coll.Add key:=metric_types(sheets_i), Item:=temp_coll2

'            section_metric_label_qty.Add key:=oTempEvaluation.getMetricSection(metric_types(sheets_i)), Item:=temp_coll

'        Next sheets_i

    Else

        ReDim metric_types(0 To 20)

        metric_types(0) = "Comment"

        metric_types(1) = "Evaluator Satisfaction"

        metric_types(2) = "Verification"

        metric_types(3) = "Accurate Information"

        metric_types(4) = "Process / Procedures"

        metric_types(5) = "Expectations"

        metric_types(6) = "Hold / Transfer"

        metric_types(7) = "Call Log"

        metric_types(8) = "Added / Updated"

        metric_types(9) = "Survey"

        metric_types(10) = "Callback"

        metric_types(11) = "Opening / Farewell"

        metric_types(12) = "Actively Listened"

        metric_types(13) = "Controlled Call"

        metric_types(14) = "Clear / Confident"

        metric_types(15) = "Hold Comment"

        metric_types(16) = "Business Comment"

        metric_types(17) = "UNKNOWN COMMENT1"

        metric_types(18) = "UNKNOWN COMMENT2"

        metric_types(19) = "UNKNOWN COMMENT3"

        metric_types(20) = "UNKNOWN COMMENT4"

       

        Set section_metric_label_qty = New MetricSctnMetricLblScoreLblQty

        Set agents_section_scores = New AgentsMetricSectionScores

       

'        Set temp_coll = New Collection

'        section_metric_label_qty.Add key:="Procedural Accuracy", Item:=temp_coll

'        Set temp_coll = New Collection

'        section_metric_label_qty.Add key:="Call Handling", Item:=temp_coll

'        Set temp_coll = New Collection

'        section_metric_label_qty.Add key:="Client Experience", Item:=temp_coll

'        Set temp_coll = New Collection

'        For sheets_i = 3 To 14

'            temp_name = ""

'            Set temp_coll = section_metric_label_qty(oTempEvaluation.getMetricSection(metric_types(sheets_i)))

'            section_metric_label_qty.Remove (oTempEvaluation.getMetricSection(metric_types(sheets_i)))

'            If isCollectionKey(metric_types(sheets_i), temp_coll) Then

'                Set temp_coll2 = temp_coll(metric_types(sheets_i))

'                temp_coll.Remove (metric_types(sheets_i))

'            Else

'                Set temp_coll2 = New Collection

'            End If

'            temp_coll2.Add key:="Yes", Item:=0

'            temp_coll2.Add key:="Partial", Item:=0

'            temp_coll2.Add key:="No", Item:=0

'            temp_coll2.Add key:="Not Likely", Item:=0

'            temp_coll2.Add key:="Likely", Item:=0

'            temp_coll2.Add key:="Definitely", Item:=0

'            temp_coll.Add key:=metric_types(sheets_i), Item:=temp_coll2

'            section_metric_label_qty.Add key:=oTempEvaluation.getMetricSection(metric_types(sheets_i)), Item:=temp_coll

'        Next sheets_i

'        Set temp_coll2 = New Collection

'        temp_coll2.Add key:="Strongly Agree", Item:=0

'        temp_coll2.Add key:="Agree", Item:=0

'        temp_coll2.Add key:="Neutral", Item:=0

'        temp_coll2.Add key:="Disagree", Item:=0

'        temp_coll2.Add key:="Strongly Disagree", Item:=0

'        Set temp_coll = New Collection

'        temp_coll.Add key:="ESAT", Item:=temp_coll2

'        section_metric_label_qty.Add key:="Evaluator Satisfaction", Item:=temp_coll

'        Set temp_coll2 = New Collection

'        Set temp_coll = New Collection

    End If

    ' section_metric_label_qty As Collection ' Collection with keys for the five sections, Evaluator Satisfaction, Procedural Accuracy, Call Handling, Client Experience, Verification pointing to collection with names of the metrics as keys, pointing to collections with the possible score labels as keys, pointing to integers that represent the count

'    Set temp_coll2 = New Collection

'    temp_coll2.Add key:="Yes", Item:=0

'    temp_coll2.Add key:="No", Item:=0

'    temp_coll.Add key:="Verification", Item:=temp_coll2

'    section_metric_label_qty.Add key:="Verification", Item:=temp_coll

   

 

    'Begin Test

 

    Dim output_cell_addr As String

    Dim original_data_ws_name As String

    Dim temp_sheet As Worksheet

    Dim current_name_exists As Boolean

    Dim temp_arr() As Double

    Dim temp_arr2() As Double

    Dim temp_date As Date

    Dim ws As Worksheet

    Dim found_later_name As Boolean

    Dim temp_esatscore_arr() As EvalProcEsatTypeDate

                    Dim temp_esateval As EvalProcEsatTypeDate

   

    Set original_data_wb = ActiveWorkbook

    original_data_ws_name = ActiveSheet.name

    Set oTheEvaluations = New EvaluationCollection

    Set temp_sheet = output_book.Sheets.Add

    temp_sheet.name = primary_output_tab_n

    first_time = True

    Set temp_sheet = output_book.Sheets(primary_output_tab_n)

   

    ' Start fresh, clear sheet content

    Call initializeCommentTab(temp_sheet)

    oTheEvaluations.setClientSatisfaction (cbxIsClientExperience.Value)

    oTheEvaluations.setIsCrd (cbxCRD.Value)

    oTheEvaluations.setIsMssnMisc (cbxMSSN.Value)

    oTheEvaluations.setIsNna (cbxNNA.Value)

    With original_data_wb.Worksheets(original_data_ws_name)

        .Activate

    End With

   

    'Core evaluation parsing code below

    On Error Resume Next

    For Each r In Range("A1", Range("A1").End(xlDown))

        metric_type = r.Value

        comment = r.offset(0, 1).Value

        Call oTheEvaluations.insertData(metric_type, comment)

'        'If IsNumeric(r.Value) Then

'        '    metric_type = Format(r.Value, "Long Date") & " " & Format(r.Value, "Long Time")

'        'Else

'        metric_type = r.Value

'        'End If

'        comment = r.offset(0, 1).Value

'        ' Begin scenarios

'        If is_agent And (IsDate(metric_type) Or IsNumeric(r.Value)) And Not IsEmpty(r.offset(0, 1)) And IsNumeric(comment) Then

'            logic_switch_token = getLogicSwitchToken

'            If isKeyOfCollection(oAgentsEvalScores, current_agent) And Not duplicate_evaluation And Not first_metric Then

'                Set oTempEvaluation = all_evaluations(UBound(all_evaluations))

'                'Call addCallHandlingMax

'

'                If Not has_esat Then

'                    Call processMissingEsat

'                ElseIf is_client_experience Then

'                    ' oAgentsEvalScores As Collection ' of arrays of EvalProcEsatTypeDate

'                    If oTempEvaluation.getAgentName = current_agent Then

'                        'temp_esatscore_arr = oAgentsEvalScores(current_agent)

'                        'Set temp_esateval = temp_esatscore_arr(UBound(temp_esatscore_arr))

'                        If Not oTempEvaluation.isBusinessExpectationNaFound Then

'                            Call oTempEvaluation.processBusinessExpectationsNa

'                        End If

'                        If (Not Not esat_scores) <> 0 Then

'                            ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

'                        Else

'                            ReDim esat_scores(0 To 0)

'                        End If

'                        esat_scores(UBound(esat_scores)) = oTempEvaluation.getSecondaryScore

'                        If sm_known Then

'                            If isKeyOfCollection(sm_esat_scores, current_sm) Then

'                                this_sm_proc = sm_esat_scores(current_sm)

'                                sm_esat_scores.Remove current_sm

'                                ReDim Preserve this_sm_proc(LBound(this_sm_proc) To UBound(this_sm_proc) + 1)

'                            Else

'                                ReDim this_sm_proc(0 To 0)

'                            End If

'                            this_sm_proc(UBound(this_sm_proc)) = oTempEvaluation.getSecondaryScore

'                            sm_esat_scores.Add key:=current_sm, Item:=this_sm_proc

'                        End If

'                    End If

'                End If

'

'                If Not oTempEvaluation.isVerificationFound Then

'                    Call oTempEvaluation.processDefaultVerification

'                End If

'            End If

''            If isAgentsFirstEvalPopulated And Not duplicate_evaluation Then

''                Call incrementOffsetOmnibus

''                temp_adate = sp_evaldate_collection(current_agent)

''                Call addAgentNameOmnibus(current_agent)

''                temp_date = temp_adate(UBound(temp_adate))

''                Call addTimeStampOmnibus(temp_date)

''                Call addMetricTypeOmnibus("FILLER-IGNORE ME")

''                Call formatFillerRowOmnibus

''            End If

'            current_eval = Format(Mid(metric_type, 1, InStr(1, metric_type, " ") - 1), "Long Date")

'            duplicate_evaluation = False

''            has_esat = False

''            business_expectation_na_resolved = False

'            If isCollectionKey(current_agent, sp_evaldate_collection) Then

'                temp_adate = sp_evaldate_collection(current_agent)

'                For i = UBound(temp_adate) To LBound(temp_adate) Step -1

'                    If temp_adate(i) = metric_type Then

'                        duplicate_evaluation = True

'                        Exit For

'                    End If

'                Next i

'            End If

'

'            If Not duplicate_evaluation Then

'                Set oTempEvaluation = New Evaluation

'                Call oTempEvaluation.setRawComment(metric_type, comment)

'                'Call handleDateProcScore(CDate(metric_type), comment)

'            End If

'        ElseIf is_agent And TypeName(metric_type) = "String" And metric_type = "Group:" Then

'           ' Do nothing

'        ElseIf TypeName(metric_type) = "String" And metric_type = "Agent:" Then

'            logic_switch_token = getLogicSwitchToken()

'            If is_agent And logic_switch_token And Not duplicate_evaluation Then

'                Set oTempEvaluation = all_evaluations(UBound(all_evaluations))

'                'Call addCallHandlingMax

'                If Not has_esat Then

'                    Call processMissingEsat

'                ElseIf is_client_experience Then

'                    ' oAgentsEvalScores As Collection ' of arrays of EvalProcEsatTypeDate

'

'                    If isCollectionKey(current_agent, oAgentsEvalScores) Then

'                        temp_esatscore_arr = oAgentsEvalScores(current_agent)

'                        Set temp_esateval = temp_esatscore_arr(UBound(temp_esatscore_arr))

'                        If Not business_expectation_na_resolved Then

'                            Call processBusinessExpectationsNa(temp_esateval)

'                        End If

'

'                        If (Not Not esat_scores) <> 0 Then

'                            ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

'                        Else

'                            ReDim esat_scores(0 To 0)

'                        End If

'                        esat_scores(UBound(esat_scores)) = temp_esateval.esat

'

'                        If sm_known Then

'                            If isKeyOfCollection(sm_esat_scores, current_sm) Then

'                                this_sm_proc = sm_esat_scores(current_sm)

'                                sm_esat_scores.Remove current_sm

'                                ReDim Preserve this_sm_proc(LBound(this_sm_proc) To UBound(this_sm_proc) + 1)

'                            Else

'                                ReDim this_sm_proc(0 To 0)

'                            End If

'                            this_sm_proc(UBound(this_sm_proc)) = temp_esateval.esat

'                            sm_esat_scores.Add key:=current_sm, Item:=this_sm_proc

'                        End If

'                        ' Put the values on the row

'                        Call addEsatScoreOmnibus(temp_esateval.esat)

'                        'garbage_text = getEsatScore(output_row_offset - 1, primary_output_tab_n)

'                        With output_book.Worksheets(primary_output_tab_n)

'                            If IsEmpty(.Range(getColumnLetter("Evaluator Satisfaction") & (output_row_offset - 1))) Then

'                                Call fillEsatScoreUp(output_row_offset, temp_esateval.esat, primary_output_tab_n)

'                                Call fillEsatScoreUp(current_agent_offset, temp_esateval.esat, current_agent)

'                                If sm_known Then

'                                    Call fillEsatScoreUp(current_sm_offset, temp_esateval.esat, current_sm)

'                                End If

'                            End If

'                        End With

'

'                    End If

'                End If

'

'                If Not verification_comment_supplied Then

'                    Call processDefaultVerification

'                End If

'                If SheetExists(current_agent) Then 'And Not first_eval Then

'                    Call incrementOffsetOmnibus

'                    temp_adate = sp_evaldate_collection(current_agent)

'                    Call addAgentNameOmnibus(current_agent)

'                    temp_date = temp_adate(UBound(temp_adate))

'                    Call addTimeStampOmnibus(temp_date)

'                    Call addMetricTypeOmnibus("FILLER-IGNORE ME")

'                    Call formatFillerRowOmnibus

'                End If

'            End If

'            previous_agent = current_agent

'            current_agent = comment

'            this_agent_eval_qty = 0

'            current_agent_offset = first_comment_row

'            verification_comment_supplied = False

'            first_eval = True

'            found_later_name = False

'            has_esat = False

'            current_sm = getSmName()

'            first_metric = True

'            business_expectation_na_resolved = False

'            If Len(current_sm) = 0 Then ' Or current_sm = "--" Then

'                is_agent = False

'                sm_known = False

'            Else

'                is_agent = True

'                output_row_offset = getLastContentRow(primary_output_tab_n)

'                If Not output_row_offset = first_comment_row Then

'                    output_row_offset = output_row_offset + 1

'                End If

'                ' Looking for SM-SP match

'                sm_known = True

'                If InStr(1, current_sm, "(SM) ") = 0 Then

'                    current_sm = "(SM) " & current_sm

'                End If

'                If Not isKeyOfCollection(sm_proc_scores, current_sm) Then

'                    ReDim this_sm_proc(0 To 0)

'                End If

'

'                If Not SheetExists(current_sm, output_book) Then

'                    addSmTab (current_sm)

'                End If

'                current_sm_offset = getLastContentRow(current_sm)

'                If Not current_sm_offset = first_comment_row Then

'                    current_sm_offset = current_sm_offset + 1

'                End If

'            End If

'            If is_agent Then

'                current_name_exists = SheetExists(current_agent)

'                If current_name_exists Then

'                    duplicate_evaluation = True

'                End If

'                If Len(previous_agent) = 0 Then

'                    Set temp_sheet = output_book.Sheets.Add(after:=output_book.Worksheets(primary_output_tab_n))

'                Else

'                    If Not current_name_exists Then

'                        For sheets_i = 2 To output_book.Worksheets.Count

'                            With output_book.Worksheets(sheets_i)

'                                If StrComp(current_agent, .name) = -1 Or Not InStr(1, .name, "(SM) ") = 0 Then

'                                    Set temp_sheet = output_book.Sheets.Add(Before:=output_book.Worksheets(sheets_i))

'                                    found_later_name = True

'                                    Exit For

'                                End If

'                            End With

'                        Next sheets_i

'                        If Not found_later_name Then

'                            Set temp_sheet = output_book.Sheets.Add(after:=output_book.Worksheets(sheets_i - 1))

'                        End If

'                    End If

'                End If

'                If Not current_name_exists Then

'                    Call initializeCommentTab(temp_sheet)

'                    temp_sheet.name = comment

'                    first_metric = True

'                End If

'            End If

'        ElseIf is_agent And TypeName(metric_type) = "String" And metric_type = "Evaluation Date" Then

'            ' Do nothing

'        ElseIf is_agent And TypeName(metric_type) = "String" And metric_type = "Form:" Then

'            ' Do nothing

'        ElseIf is_agent And TypeName(metric_type) = "String" And metric_type = "Report Period:" Then

'            ' Do nothing

'        ElseIf TypeName(metric_type) = "String" And Len(metric_type) = 0 Then

'            Exit For

'        ElseIf is_agent And Not duplicate_evaluation And TypeName(metric_type) = "String" Then

'            Call handleDefaultText(metric_type, comment)

'        End If

    Next r

    Err.Clear

    On Error GoTo 0

' ****************************************************************************************

  ' Process metadata

  Call oTheEvaluations.resolveFinalEvaluation

  ' Go to the beginning of sort order

  oTheEvaluations.resetIndex

  Dim sCurrentSmName As String

  Dim aSmNameList() As String

  Dim aDoubles() As Double

  Dim aSmPrimaryScores() As Double

  Dim aSmSecondaryScores() As Double

  Dim oSmSecondaryValid As Collection

  Set oSmSecondaryValid = New Collection

  Dim dPreviousAgentPrimaryAvg As Double

  Dim dPreviousAgentSecondaryAvg As Double

  Dim bPreviousAgentSecondaryValid As Boolean

  Dim bTempBoole As Boolean

  Dim aEvalStatsArray() As EvalProcEsatTypeDate

  ' Loop through all evaluations

  While oTheEvaluations.isIndexValid

    If Len(previous_agent) = 0 And Len(current_agent) = 0 Then

     ReDim aEvalStatsArray(0 To 0)

      Set aEvalStatsArray(UBound(aEvalStatsArray)) = oTheEvaluations.getCurrentEvalGeneralStatistics()

      oAgentsEvalScores.Add key:=oTheEvaluations.getCurrentEvalAgentName(), Item:=aEvalStatsArray

      this_agent_eval_qty = 1

      eval_qty_max = 1

    End If

    ' Filler between timestamps after the first

    If Not current_eval = oTheEvaluations.getCurrentEvalTimeStamp And Len(current_agent) > 0 Then

      Call frmReportBuilderSubmit.addAgentNameOmnibus(current_agent)

      Call frmReportBuilderSubmit.addTimeStampOmnibus(current_eval)

      Call frmReportBuilderSubmit.addMetricTypeOmnibus("FILLER-IGNORE ME")

      Call formatFillerRowOmnibus

      Call incrementOffsetOmnibus

      ' Document secondary score deficit

     If Not oTheEvaluations.isSecondaryAvgAvailableForEval(oTheEvaluations.getCurrentEvalAgentName, oTheEvaluations.getCurrentEvalTimeStamp) Then

        Call addEsatDeficitEntry(oTheEvaluations.getCurrentEvalAgentName, oTheEvaluations.getCurrentEvalTimeStamp)

        esat_qty_deficiency = esat_qty_deficiency + 1

      End If

      ' Create EvalProcEsatTypeDate object entry for this evaluation. Will need this for Evaluation Type-based statistics and Verification

      If frmReportBuilderSubmit.isKeyOfCollection(oAgentsEvalScores, oTheEvaluations.getCurrentEvalAgentName) Then

        aEvalStatsArray = oAgentsEvalScores.Item(oTheEvaluations.getCurrentEvalAgentName)

        oAgentsEvalScores.Remove (oTheEvaluations.getCurrentEvalAgentName)

        ReDim Preserve aEvalStatsArray(LBound(aEvalStatsArray) To UBound(aEvalStatsArray) + 1)

      Else

        ReDim aEvalStatsArray(0 To 0)

      End If

      Set aEvalStatsArray(UBound(aEvalStatsArray)) = oTheEvaluations.getCurrentEvalGeneralStatistics()

      oAgentsEvalScores.Add key:=oTheEvaluations.getCurrentEvalAgentName, Item:=aEvalStatsArray

    End If

    ' Define SM status

    sCurrentSmName = frmReportBuilderSubmit.getSmName(oTheEvaluations.getCurrentEvalAgentName)

    If Len(frmReportBuilderSubmit.stripLeadTrailNewline(sCurrentSmName)) > 0 And Not sCurrentSmName = "--" Then

      sm_known = True

      sCurrentSmName = formatSmName(sCurrentSmName)

    Else

      sm_known = False

    End If

    ' Action when new evaluation, but different if also different SP

    If Len(current_agent) > 0 And Not current_eval = oTheEvaluations.getCurrentEvalTimeStamp And current_agent = oTheEvaluations.getCurrentEvalAgentName Then

      ' Searching for largest number of evaluations for any individual SP

      this_agent_eval_qty = this_agent_eval_qty + 1

      If this_agent_eval_qty > eval_qty_max Then

        eval_qty_max = this_agent_eval_qty

      End If

    ElseIf Len(current_agent) > 0 And Not current_agent = oTheEvaluations.getCurrentEvalAgentName Then

      this_agent_eval_qty = 1

    End If

    ' Set SM tab

    If sm_known Then

      If Not SheetExists(sCurrentSmName, output_book) Then

        If (Not Not aSmNameList) = 0 Then

          ReDim aSmNameList(0 To 0)

        Else

          ReDim Preserve aSmNameList(LBound(aSmNameList) To UBound(aSmNameList) + 1)

        End If

        aSmNameList(UBound(aSmNameList)) = sCurrentSmName

        Call addSmTab(sCurrentSmName)

        current_sm_offset = first_comment_row

      Else

        current_sm_offset = frmReportBuilderSubmit.getLastContentRow(sCurrentSmName) + 1

      End If

    End If

    ' Evaluation scores for SM averages, if first evaluation, or new evaluation, and SM relevant

    If sm_known And (Len(current_agent) = 0 Or Not current_eval = oTheEvaluations.getCurrentEvalTimeStamp) Then

      ' Primary

      If frmReportBuilderSubmit.isKeyOfCollection(sm_proc_scores, sCurrentSmName) Then

        aSmPrimaryScores = sm_proc_scores(sCurrentSmName)

        sm_proc_scores.Remove (sCurrentSmName)

        Call frmReportBuilderSubmit.incrementDoubleArrayLength(aSmPrimaryScores)

      Else

        ReDim aSmPrimaryScores(0 To 0)

      End If

      aSmPrimaryScores(UBound(aSmPrimaryScores)) = oTheEvaluations.getCurrentEvalPrimaryScore

      sm_proc_scores.Add key:=sCurrentSmName, Item:=aSmPrimaryScores

      ' Secondary

      If oTheEvaluations.isCurrentEvalSecondaryScoreValid Then

        If frmReportBuilderSubmit.isKeyOfCollection(sm_esat_scores, sCurrentSmName) Then

          aSmSecondaryScores = sm_esat_scores(sCurrentSmName)

          sm_esat_scores.Remove (sCurrentSmName)

          Call frmReportBuilderSubmit.incrementDoubleArrayLength(aSmSecondaryScores)

        Else

          ReDim aSmSecondaryScores(0 To 0)

        End If

        aSmSecondaryScores(UBound(aSmSecondaryScores)) = oTheEvaluations.getCurrentEvalSecondaryScore

        sm_esat_scores.Add key:=sCurrentSmName, Item:=aSmSecondaryScores

      End If

      ' Secondary validity

      If Not frmReportBuilderSubmit.isKeyOfCollection(oSmSecondaryValid, sCurrentSmName) Then

        oSmSecondaryValid.Add key:=sCurrentSmName, Item:=oTheEvaluations.isSecondaryAvgAvailableForEval(oTheEvaluations.getCurrentEvalAgentName, oTheEvaluations.getCurrentEvalTimeStamp)

      ElseIf oSmSecondaryValid(sCurrentSmName) Then

        If frmReportBuilderSubmit.isKeyOfCollection(oSmSecondaryValid, sCurrentSmName) Then

          oSmSecondaryValid.Remove (sCurrentSmName)

        End If

        oSmSecondaryValid.Add key:=sCurrentSmName, Item:=oTheEvaluations.isSecondaryAvgAvailableForEval(oTheEvaluations.getCurrentEvalAgentName, oTheEvaluations.getCurrentEvalTimeStamp)

      End If

    End If

    ' Add average for previous agent

    If Not Len(current_agent) = 0 And Not current_agent = oTheEvaluations.getCurrentEvalAgentName Then

      Call addAgentAverages(current_agent, "agent", dPreviousAgentPrimaryAvg, dPreviousAgentSecondaryAvg, bPreviousAgentSecondaryValid)

      ' Add the averages for the respective agent to Quality Ranking, if applicable

      Set temp_sheet = getWorkSheet(current_agent)

      With temp_sheet

        If update_qr Then

          Call addQualityRankingProcedural(current_agent, .Range(getColumnLetter("Procedural Score") & getLastContentRow(.name)).offset(avg_row_offset, 0).Value, eval_month, eval_year)

          If oTheEvaluations.isSecondaryAvgAvailableForEval(current_agent, current_eval) Then

            Call addQualityRankingEsat(.name, .Range(getColumnLetter("Evaluator Satisfaction") & getLastContentRow(.name)).offset(avg_row_offset, 0).Value, eval_month, eval_year)

          End If

        End If

      End With

    ElseIf Len(current_agent) > 0 Then

      dPreviousAgentPrimaryAvg = oTheEvaluations.getPrimaryAvgForAgent(current_agent)

      bPreviousAgentSecondaryValid = oTheEvaluations.isSecondaryAvgAvailableForEval(current_agent, current_eval)

      If bPreviousAgentSecondaryValid Then

        dPreviousAgentSecondaryAvg = oTheEvaluations.getSecondaryAvgForAgent(current_agent)

      End If

    End If

    ' New agent, change variable values

    If Not current_agent = oTheEvaluations.getCurrentEvalAgentName Then

      previous_agent = current_agent

      current_agent = oTheEvaluations.getCurrentEvalAgentName

      If frmReportBuilderSubmit.SheetExists(current_agent, output_book) Then

        current_agent_offset = frmReportBuilderSubmit.getLastContentRow(current_agent)

        If frmReportBuilderSubmit.isKeyOfCollection(oAgentsEvalScores, current_agent) Then

          aEvalStatsArray = oAgentsEvalScores(current_agent)

          this_agent_eval_qty = UBound(aEvalStatsArray) - LBound(aEvalStatsArray) + 1

        End If

      Else

        current_agent_offset = first_comment_row

        this_agent_eval_qty = 1

      End If

    End If

    ' New evaluation, change variable value

    If Not current_eval = oTheEvaluations.getCurrentEvalTimeStamp Then

      current_eval = oTheEvaluations.getCurrentEvalTimeStamp

    End If

    ' Add values to sheets

    Call frmReportBuilderSubmit.addAgentNameOmnibus(oTheEvaluations.getCurrentEvalAgentName)

    Call frmReportBuilderSubmit.addCommentOmnibus(oTheEvaluations.getCurrentEvalComment)

    Call frmReportBuilderSubmit.addEvalTypeOmnibus(oTheEvaluations.getCurrentEvalEvalType)

    Call frmReportBuilderSubmit.addMetricMaxOmnibus(oTheEvaluations.getCurrentEvalMaxScore)

    Call frmReportBuilderSubmit.addMetricPercentageOmnibus(oTheEvaluations.getCurrentEvalMetricPercentage)

    Call frmReportBuilderSubmit.addMetricScoreLabelOmnibus(oTheEvaluations.getCurrentEvalScoreLabel)

    Call frmReportBuilderSubmit.addMetricTypeOmnibus(oTheEvaluations.getCurrentEvalMetricType)

    Call frmReportBuilderSubmit.addTimeStampOmnibus(oTheEvaluations.getCurrentEvalTimeStamp)

    Call frmReportBuilderSubmit.addProceduralScoreOmnibus(oTheEvaluations.getCurrentEvalPrimaryScore)

    Call frmReportBuilderSubmit.addMetricScoreOmnibus(oTheEvaluations.getCurrentEvalMetricScore)

    Call addSmPrimaryScore(oTheEvaluations.getCurrentEvalPrimaryScore, sCurrentSmName)

    ' Add to Bad Format tab, if relevant

    If oTheEvaluations.isCurrentRowMetadataBadFormat() Then

      Call addGarbageMetadataRow(oTheEvaluations.getCurrentEvalAgentName, oTheEvaluations.getCurrentEvalTimeStamp, oTheEvaluations.getCurrentEvalEvalType, _

          oTheEvaluations.getCurrentEvalOriginalComment, oTheEvaluations.getCurrentEvalMetricType, oTheEvaluations.getCurrentEvalScoreLabel, _

          oTheEvaluations.getCurrentEvalMaxScore, oTheEvaluations.getCurrentEvalPrimaryScore, oTheEvaluations.getCurrentEvalSecondaryScore)

    Else

      ' Good metadata, so track metric label

      Call agents_section_scores.addScoreToAgentMetricSection( _

          oTheEvaluations.getCurrentEvalAgentName, frmReportBuilderSubmit.getMetricSection( _

              oTheEvaluations.getCurrentEvalMetricType), _

          oTheEvaluations.getCurrentEvalMetricScore)

      Call section_metric_label_qty.incrementQuantity(frmReportBuilderSubmit.getMetricSection( _

              oTheEvaluations.getCurrentEvalMetricType), oTheEvaluations.getCurrentEvalMetricType, oTheEvaluations.getCurrentEvalScoreLabel)

    End If

    ' Add secondary score, if valid

    If oTheEvaluations.isCurrentEvalSecondaryScoreValid Then

      Call frmReportBuilderSubmit.addEsatScoreOmnibus(oTheEvaluations.getCurrentEvalSecondaryScore)

      Call addSmSecondaryScore(oTheEvaluations.getCurrentEvalSecondaryScore, sCurrentSmName)

    End If

    ' Next!

    Call incrementOffsetOmnibus

    Call oTheEvaluations.moveIndexForward

  Wend

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Call oTheEvaluations.resetIndex

  If Not oTheEvaluations.isSecondaryAvgAvailableForEval(current_agent, current_eval) Then

    Call addEsatDeficitEntry(current_agent, current_eval)

    esat_qty_deficiency = esat_qty_deficiency + 1

  End If

  Call addAgentAverages(current_agent, "agent", dPreviousAgentPrimaryAvg, dPreviousAgentSecondaryAvg, bPreviousAgentSecondaryValid)

  If (Not Not aSmNameList) <> 0 Then

    For i = LBound(aSmNameList) To UBound(aSmNameList)

      aDoubles = sm_proc_scores(aSmNameList(i))

      Dim aDoublesAgain() As Double

      If frmReportBuilderSubmit.isKeyOfCollection(sm_esat_scores, aSmNameList(i)) Then

        aDoublesAgain = sm_esat_scores(aSmNameList(i))

      End If

      Call addAgentAverages(aSmNameList(i), "SM", frmReportBuilderSubmit.getDoubleArrayAverage(aDoubles), _

          frmReportBuilderSubmit.getDoubleArrayAverage(aDoublesAgain), oSmSecondaryValid(aSmNameList(i)))

    Next i

  End If

  Call addAgentAverages(primary_output_tab_n, "all", oTheEvaluations.getPrimaryAvgForAll, oTheEvaluations.getSecondaryAvgForAll, oTheEvaluations.isSecondaryAvgAvailableForAll)

  With output_book.Worksheets(primary_output_tab_n)

    If Not IsEmpty(.Range(getColumnLetter(first_clmn_label) & first_comment_row).Value) Then

      Set r = .Range(getColumnLetter("Procedural Score") & getLastContentRow(.name)).offset(avg_row_offset, 0)

      r.offset(1, 0).WrapText = True

      Call applyHeaderColor(.Range(r.offset(1, 0), r.offset(1, 1)))

      Call applyHeaderFontStyle(.Range(r.offset(1, 0), r.offset(1, 1)))

      r.offset(1, 1).WrapText = True

      .Rows(r.offset(1, 0).row).RowHeight = 35

      If truncate_numbers Then

        r.NumberFormat = "0.00"

        r.offset(0, 1).NumberFormat = "0.00"

      End If

    End If

  End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  Set ws = output_book.Worksheets(primary_output_tab_n)

  If Len(ws.Range(getColumnLetter(first_clmn_label) & first_comment_row)) > 0 Then

'        ' Resolve last evaluation

'        If SheetExists(current_agent, output_book) Then

'            ' Resolve calculations for last agent

'            Call addCallHandlingMax

'            ' If ESAT less than Procedural

'            If Not has_esat And Not duplicate_evaluation Then

'                Call processMissingEsat

'            ElseIf is_client_experience Then

'                ' oAgentsEvalScores As Collection ' of arrays of EvalProcEsatTypeDate

'                If isCollectionKey(current_agent, oAgentsEvalScores) Then

'                    temp_esatscore_arr = oAgentsEvalScores(current_agent)

'                    Set temp_esateval = temp_esatscore_arr(UBound(temp_esatscore_arr))

'                    If Not business_expectation_na_resolved Then

'                        Call processBusinessExpectationsNa(temp_esateval)

'                    End If

'                    If (Not Not esat_scores) <> 0 Then

'                        ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

'                    Else

'                        ReDim esat_scores(0 To 0)

'                    End If

'                    esat_scores(UBound(esat_scores)) = temp_esateval.esat

'                    If sm_known Then

'                        If isKeyOfCollection(sm_esat_scores, current_sm) Then

'                            this_sm_proc = sm_esat_scores(current_sm)

'                            sm_esat_scores.Remove current_sm

'                            ReDim Preserve this_sm_proc(LBound(this_sm_proc) To UBound(this_sm_proc) + 1)

'                        Else

'                            ReDim this_sm_proc(0 To 0)

'                        End If

'                        this_sm_proc(UBound(this_sm_proc)) = temp_esateval.esat

'                        sm_esat_scores.Add key:=current_sm, Item:=this_sm_proc

'                    End If

'                    ' Put the values on the row

'                    Call addEsatScoreOmnibus(temp_esateval.esat)

'                    'garbage_text = getEsatScore(output_row_offset - 1, primary_output_tab_n)

'                    With output_book.Worksheets(primary_output_tab_n)

'                        If IsEmpty(.Range(getColumnLetter("Evaluator Satisfaction") & (output_row_offset - 1))) Then

'                            Call fillEsatScoreUp(output_row_offset, temp_esateval.esat, primary_output_tab_n)

'                            Call fillEsatScoreUp(current_agent_offset, temp_esateval.esat, current_agent)

'                            If sm_known Then

'                                Call fillEsatScoreUp(current_sm_offset, temp_esateval.esat, current_sm)

'                            End If

'                        End If

'                    End With

'

'                End If

'            End If

'            With output_book.Worksheets(current_agent)

'                Application.Goto .Range("A1"), True

'                With .Range(getColumnLetter(first_clmn_label) & first_comment_row)

'                    .Select

'                    ActiveWindow.FreezePanes = True

'                End With

'            End With

'            If Not verification_comment_supplied And Not duplicate_evaluation Then

'                Call processDefaultVerification

'            End If

'            If Not duplicate_evaluation Then

'                Call incrementOffsetOmnibus

'                temp_adate = sp_evaldate_collection(current_agent)

'                Call addAgentNameOmnibus(current_agent)

'                temp_date = temp_adate(UBound(temp_adate))

'                Call addTimeStampOmnibus(temp_date)

'                Call addMetricTypeOmnibus("FILLER-IGNORE ME")

'                Call formatFillerRowOmnibus

'            End If

'        End If

'

'        ' Primary comment tab averages for all evaluations

'        ' Apply calculations to All Agents Notes primary_output_tab_n

'        For i = UBound(proc_scores) To LBound(proc_scores) Step -1

'            proc_sub_total = proc_sub_total + proc_scores(i)

'        Next i

'        If (Not Not esat_scores) <> 0 Then

'            For i = UBound(esat_scores) To LBound(esat_scores) Step -1

'                esat_sub_total = esat_sub_total + esat_scores(i)

'            Next i

'        End If

'        With output_book.Worksheets(primary_output_tab_n)

'            Set r = .Range(getColumnLetter("Procedural Score") & .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).row).offset(avg_row_offset, 0)

'            r.Value = proc_sub_total / (UBound(proc_scores) + 1)

'            r.offset(1, 0).Value = "Procedural Score -- Team Average"

'            r.offset(1, 0).WrapText = True

'            If esat_qty_deficiency = 0 Then

'                r.offset(0, 1).Value = esat_sub_total / (UBound(esat_scores) + 1)

'            Else

'                r.offset(0, 1).Value = "Unable to Calculate. Please see Bad Format Rows tab."

'            End If

'            r.offset(1, 1).Value = "Evaluator Satisfaction -- Team Average"

'            r.offset(1, 1).WrapText = True

'            .Rows(r.offset(1, 0).row).RowHeight = 30

'            Set r = .Range(r.offset(1, 0), r.offset(1, 1))

'            Call applyHeaderColor(r)

'            Call applyHeaderFontStyle(r)

'            If truncate_numbers Then

'                Range(r, r.offset(0, 1)).NumberFormat = "0.00"

'            End If

'        End With

           

        ' Make appearance of sheets uniform, hide raw data

        Dim miscellaneous_num As Integer

        Dim temp_clmn_letter As String

        Dim custom_sort_key As Range

       

        If SheetExists("Quality Ranking") Then

            With output_book.Worksheets("Quality Ranking")

                .Move after:=output_book.Worksheets(Worksheets.Count)

            End With

        End If

        If SheetExists(sm_sp_sheet_name) Then

            With output_book.Worksheets(sm_sp_sheet_name)

                .Move after:=output_book.Worksheets(Worksheets.Count)

            End With

        End If

        If esat_qty_deficiency > 0 And SheetExists("Bad Format Rows") Then

            With output_book.Worksheets("Bad Format Rows")

                With .Range(getColumnLetter(last_clmn_label) & first_header_row).offset(0, 3)

                    .Value = esat_qty_deficiency

                End With

                With .Tab

                    .Color = 255

                    .TintAndShade = 0

                End With

                .Columns(13).AutoFit

                .Columns(14).AutoFit

            End With

        End If

        ' ****************************************************************************

        For sheets_i = 1 To output_book.Worksheets.Count

            With output_book.Worksheets(sheets_i)

                If .name = "Raw" Then

                    .Visible = xlSheetHidden

                End If

                If .name = "Quality Ranking" Then

                    GoTo NextIteration

                End If

                If .name = "Form Constants" Then

                    .Visible = xlSheetVeryHidden

                End If

                If Not .Visible = xlSheetHidden And Not .Visible = xlSheetVeryHidden And Not .name = sm_sp_sheet_name And Not IsEmpty(.Range(getColumnLetter(first_clmn_label) & first_comment_row).Value) Then

                    comment = .name

                    If Not .name = primary_output_tab_n And Not .name = "Bad Format Rows" And (Not InStr(1, .name, ",") = 0 Or Not InStr(1, .name, "(SM) ") = 0) Then

                        ' Resolve calculation for SM

                        If Not InStr(1, .name, "(SM) ") = 0 Then

                            proc_sub_total = 0

                            esat_sub_total = 0

                            this_sm_proc = sm_proc_scores.Item(comment)

                            If isCollectionKey(.name, sm_esat_scores) Then

                                this_sm_esat = sm_esat_scores.Item(.name)

                            End If

                           

                            Application.Goto .Range("A1"), True

                            With .Range(getColumnLetter(first_clmn_label) & first_comment_row)

                                .Select

                                ActiveWindow.FreezePanes = True

                            End With

                            ' The average scores block

                            Set r = .Range(getColumnLetter("Procedural Score") & .Range(getColumnLetter(first_clmn_label) & first_comment_row).End(xlDown).offset(avg_row_offset, 0).row)

'                            r.Value = frmReportBuilderSubmit.getDoubleArrayAverage(this_sm_proc)

'                            r.offset(1, 0).Value = "Procedural Score -- " & .name & " Team Average"

'                            r.offset(1, 0).WrapText = True

'                            If esat_sub_total > 0 Then

'                                r.offset(0, 1).Value = frmReportBuilderSubmit.getDoubleArrayAverage(this_sm_esat)

'                            Else

'                                r.offset(0, 1).Value = "Unable to Calculate. Please see Bad Format Rows tab."

'                            End If

                            r.offset(1, 1).Value = "Evaluator Satisfaction -- " & .name & " Team Average"

                            r.offset(1, 1).WrapText = True

                            .Rows(r.offset(1, 0).row).RowHeight = 35

                            r.offset(1, 0).WrapText = True

                            If truncate_numbers Then

                                r.NumberFormat = "0.00"

                                r.offset(0, 1).NumberFormat = "0.00"

                            End If

                            Call applyHeaderColor(.Range(r.offset(1, 0), r.offset(1, 1)))

                            Call applyHeaderFontStyle(.Range(r.offset(1, 0), r.offset(1, 1)))

                        ElseIf Not .name = primary_output_tab_n Then

                            ' Agent tabs

                            If Not IsEmpty(.Range(getColumnLetter(first_clmn_label) & first_comment_row).Value) Then

                                Set r = .Range(getColumnLetter("Procedural Score") & getLastContentRow(.name)).offset(avg_row_offset, 0)

                                r.offset(1, 1).WrapText = True

                                .Rows(r.offset(1, 0).row).RowHeight = 30

                                r.offset(1, 0).WrapText = True

                                Call applyHeaderColor(.Range(r.offset(1, 0), r.offset(1, 1)))

                                Call applyHeaderFontStyle(.Range(r.offset(1, 0), r.offset(1, 1)))

                                'Call addAgentAverages(.name, "agent")

                               

                            End If

                        End If

                    End If

 

                    ' Remove gridlines and enable freezing of top row

                    .Activate

                    ActiveWindow.DisplayGridlines = False

                    Application.Goto .Range("A1"), True

                    With .Range(getColumnLetter(first_clmn_label) & first_comment_row)

                        .Select

                        ActiveWindow.FreezePanes = True

                    End With

                   

'                    If .name = primary_output_tab_n Then

'                        Call addAgentAverages(.name, "all")

'                    End If

                    ' Adjust format of comment sheeets for uniform look

                    Set temp_sheet = output_book.Worksheets(sheets_i)

                    .Columns(getColumnLetter("Comment")).WrapText = True

                    .Columns(getColumnLetter("Comment")).ColumnWidth = 50

                    .Columns(getColumnLetter("Agent")).AutoFit

                    .Columns(getColumnLetter("Metric Type")).AutoFit

                    .Columns(getColumnLetter("Evaluation Type")).AutoFit

                    .Columns(getColumnLetter("Time Stamp")).AutoFit

                    .Columns(getColumnLetter("Metric Score")).AutoFit

                    .Columns(getColumnLetter("Maximum Metric Score")).AutoFit

                    temp_clmn_letter = getColumnLetter("Metric Percentage")

                    .Columns(temp_clmn_letter).AutoFit

                    .Range(temp_clmn_letter & "2:" & temp_clmn_letter & .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).row).NumberFormat = "0.00%"

                    .Columns(getColumnLetter("Procedural Score")).ColumnWidth = 13

                    .Columns(getColumnLetter("Evaluator Satisfaction")).ColumnWidth = 17

                    .Rows("1:" & .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).row).AutoFit

                    If .ListObjects.Count = 0 Then

                        rows_i2 = .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).row

                        .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(getColumnLetter(first_clmn_label) & "1:" & getColumnLetter(last_clmn_label) & rows_i2), _

                                XlListObjectHasheaders:=xlYes).name = "Table " & sheets_i

                    End If

                    ' Sort comment sheets

'                    With .ListObjects("Table " & sheets_i)

'                        .TableStyle = "TableStyleMedium2"

'                        Set custom_sort_key = .ListColumns("Metric Type").Range

'                        Set sort_key1 = .ListColumns("Agent").Range

'                        Set sort_key2 = .ListColumns("Time Stamp").Range

'                        With .Sort

'                            .SortFields.Clear

'                            .SortFields.Add key:=custom_sort_key, SortOn:=xlSortOnValues, ORDER:=xlAscending, CustomOrder:= _

'                            Join(metric_types, ",") '& "FILLER-IGNORE ME"

'                            '"Comment,Verification,Evaluator Satisfaction,ESAT,Accurate Information,Process / Procedures,Expectations,Hold / Transfer,Call Log,Added / Updated,Survey,Call Back,Callback,Opening / Farewell,Actively Listened,Controlled Call,Clear / Confident,Hold Comment,Business Comment,FILLER-IGNORE ME" _

'                            , DataOption:=xlSortNormal

'                            .Header = xlYes

'                            '.Orientation = XlTopBottom

'                            .Apply

'                            .SortFields.Clear

'                            .SortFields.Add key:=sort_key1, SortOn:=xlSortOnValues, ORDER:=xlAscending, DataOption:=xlSortNormal

'                            .SortFields.Add key:=sort_key2, SortOn:=xlSortOnValues, ORDER:=xlAscending, DataOption:=xlSortNormal

'                            .Header = xlYes

'                            .Apply

'                        End With

'                    End With

                End If

            End With

NextIteration:

        Next sheets_i

        ' ************************************************************************************

        ' Resolve Quality Ranking sort and add new month's ranks

        If update_qr Then

            With output_book.Worksheets("Quality Ranking")

                With .ListObjects(qr_proc_named_range_n)

                    Set sort_key1 = .ListColumns(month_proc_score_clmn_n).Range

                    With .Sort

                        .SortFields.Clear

                        .SortFields.Add key:=sort_key1, SortOn:=xlSortOnValues, ORDER:=xlDescending, DataOption:=xlSortNormal

                        .Header = xlYes

                        .Apply

                    End With

                End With

                With .ListObjects(qr_esat_named_range_n)

                    Set sort_key1 = .ListColumns(month_esat_score_clmn_n).Range

                    With .Sort

                        .SortFields.Clear

                        .SortFields.Add key:=sort_key1, SortOn:=xlSortOnValues, ORDER:=xlDescending, DataOption:=xlSortNormal

                        .Header = xlYes

                        .Apply

                    End With

                End With

                If created_new_month Then

                    .ListObjects(qr_proc_named_range_n).ListColumns("Latest Month").DataBodyRange.Copy (.ListObjects(qr_proc_named_range_n).ListColumns("Previous Month").DataBodyRange)

                    .ListObjects(qr_proc_named_range_n).ListColumns("Latest Month").DataBodyRange(1, 1).Value = 1

                    .ListObjects(qr_esat_named_range_n).ListColumns("Latest Month").DataBodyRange.Copy (.ListObjects(qr_esat_named_range_n).ListColumns("Previous Month").DataBodyRange)

                    .ListObjects(qr_esat_named_range_n).ListColumns("Latest Month").DataBodyRange(1, 1).Value = 1

                    With .ListObjects(qr_proc_named_range_n).ListColumns("Latest Month").DataBodyRange

                        .DataSeries Rowcol:=xlColumns, Step:=1

                    End With

                    With .ListObjects(qr_esat_named_range_n).ListColumns("Latest Month").DataBodyRange

                        .DataSeries Rowcol:=xlColumns, Step:=1

                    End With

                    For i = 1 To .Range("A1").End(xlToRight).Column

                        .Columns("A").AutoFit

                    Next i

                   

                    For Each r In .ListObjects(qr_proc_named_range_n).ListColumns(getQrCurrentMonthProc(eval_month, eval_year)).DataBodyRange

                        If Len(r.Value) = 0 Then

                            .Rows(r.row).EntireRow.Hidden = True

                        End If

                    Next r

                    For Each r In .ListObjects(qr_esat_named_range_n).ListColumns(getQrCurrentMonthEsat(eval_month, eval_year)).DataBodyRange

                        If Len(r.Value) = 0 Then

                            .Rows(r.row).EntireRow.Hidden = True

                        End If

                    Next r

                End If

                If truncate_numbers Then

                    .ListObjects(qr_proc_named_range_n).ListColumns("Latest Month").DataBodyRange.NumberFormat = "0.00"

                    .ListObjects(qr_proc_named_range_n).ListColumns("Previous Month").DataBodyRange.NumberFormat = "0.00"

                Else

                    .ListObjects(qr_proc_named_range_n).ListColumns("Latest Month").DataBodyRange.NumberFormat = "0.00000000"

                    .ListObjects(qr_proc_named_range_n).ListColumns("Previous Month").DataBodyRange.NumberFormat = "0.00000000"

                End If

            End With

        End If

        '***********************************************************************************************

       

    ' Metric Perentages PivotTable

        Dim pt_sheet As Worksheet

        Dim metric_percentages_pv As PivotTable

        Dim temp_pt_field As PivotField

        Set pt_sheet = output_book.Worksheets.Add(after:=output_book.Worksheets(primary_output_tab_n))

        pt_sheet.name = "Delete Me"

        output_book.Worksheets("Delete Me").Activate

        With output_book.Worksheets("Delete Me")

            .Select

            .Range("A1").Select

        End With

        ActiveWindow.DisplayGridlines = False

''''''''''''''' Though primary_output_tab_n Worksheet has Table over header and all comment rows, and the PivotTable below has the correct target data, the PT is not populating with that data _

thus, to prevent a bad user experience, this operation is rendered as "comment" lines, for now ''''''''''''''''''''''''''

        With output_book.Worksheets(primary_output_tab_n)

            Set metric_percentages_pv = .PivotTableWizard

        End With

        metric_percentages_pv.name = "Metric Percentages"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        With output_book.Worksheets("Delete Me")

            .Visible = False

        End With

        For Each pt_sheet In output_book.Worksheets

          If pt_sheet.name = primary_output_tab_n Then

            With output_book.Worksheets(pt_sheet.index + 1)

              .name = "Metric Percentages"

            End With

          End If

        Next pt_sheet

       

'        Set temp_pt_field = metric_percentages_pv.PivotFields("Agent")

'        temp_pt_field.Orientation = xlPageField

'        Set temp_pt_field = metric_percentages_pv.PivotFields("Evaluation Type")

'        temp_pt_field.Orientation = xlPageField

'        Set temp_pt_field = metric_percentages_pv.PivotFields("Time Stamp")

'        temp_pt_field.Orientation = xlPageField

'        Set temp_pt_field = metric_percentages_pv.PivotFields("Metric Percentage")

'        temp_pt_field.Orientation = xlDataField

'        temp_pt_field.Function = xlAverage

'        temp_pt_field.NumberFormat = "0.00%"

'        temp_pt_field.name = "Metric Percentage Avg"

'        Set temp_pt_field = metric_percentages_pv.PivotFields("Metric Type")

'        temp_pt_field.Orientation = xlColumnField

'        temp_pt_field.name = "Metrics"

'        Call clearPivotFilter(temp_pt_field)

'        If is_client_experience Then

'            Call filterPivotField(temp_pt_field, "Accuracy / Completeness")

'            Call clearPivotFilter(temp_pt_field, "FILLER-IGNORE ME")

'            Call filterPivotField(temp_pt_field, "Complete Expectations")

'            Call filterPivotField(temp_pt_field, "Timely Resolution")

'            Call filterPivotField(temp_pt_field, "World-Class Service")

'            Call filterPivotField(temp_pt_field, "Hold Experience")

'            Call filterPivotField(temp_pt_field, "Transfer Experience")

'            Call filterPivotField(temp_pt_field, "Appropriate Greeting")

'            Call filterPivotField(temp_pt_field, "Correct Resources")

'            Call filterPivotField(temp_pt_field, "Appropriate Closing")

'            Call filterPivotField(temp_pt_field, "Business Processes")

'            Call filterPivotField(temp_pt_field, "Verification")

'        Else

'            Call filterPivotField(temp_pt_field, "ESAT")

'            Call clearPivotFilter(temp_pt_field, "FILLER-IGNORE ME")

'            Call filterPivotField(temp_pt_field, "Accurate Information")

'            Call filterPivotField(temp_pt_field, "Expectations")

'            Call filterPivotField(temp_pt_field, "Process / Procedures")

'            Call filterPivotField(temp_pt_field, "Opening / Farewell")

'            Call filterPivotField(temp_pt_field, "Actively Listened")

'            Call filterPivotField(temp_pt_field, "Controlled Call")

'            Call filterPivotField(temp_pt_field, "Clear / Confident")

'            Call filterPivotField(temp_pt_field, "Hold / Transfer")

'            Call filterPivotField(temp_pt_field, "Call Log")

'            Call filterPivotField(temp_pt_field, "Added / Updated")

'            Call filterPivotField(temp_pt_field, "Survey")

'            Call filterPivotField(temp_pt_field, "Callback")

'            Call filterPivotField(temp_pt_field, "Verification")

'        End If

'        metric_percentages_pv.TableRange2.HorizontalAlignment = xlRight

        ' Put Bad Format Rows tab at beginning

        If SheetExists("Bad Format Rows", output_book) Then

            output_book.Worksheets("Bad Format Rows").Move before:=output_book.Worksheets(1)

        End If

       

        ' *********************************************************************************

        ' * Quality Summary

        ' *********************************************************************************

        Dim qs_row_offset As Integer

        Dim qs_ul_row_eval_scores As Integer

        Dim qs_ul_clm_eval_scores As Integer

        Dim qs_ul_row_verification As Integer

        Dim qs_ul_clm_verification As Integer

        Dim qs_ul_row_verbal As Integer

        Dim qs_ul_clm_verbal As Integer

        Dim qs_ul_row_business As Integer

        Dim qs_ul_clm_business As Integer

        Dim qs_ul_row_certification As Integer

        Dim qs_ul_clm_certification As Integer

        Dim qs_ul_row_survey As Integer

        Dim qs_ul_clm_survey As Integer

        Dim qs_ul_row_negative As Integer

        Dim qs_ul_clm_negative As Integer

        Dim qs_ul_row_written As Integer

        Dim qs_ul_clm_written As Integer

        Dim qs_ul_row_metric_lbl_metric_qty As Integer

        Dim qs_ul_clm_metric_lbl_metric_qty As Integer

        Dim qs_ul_row_metric_lbl_section_pct As Integer

        Dim qs_ul_clm_metric_lbl_section_pct As Integer

        Dim qs_ul_row_metric_lbl_section_qty As Integer

        Dim qs_ul_clm_metric_lbl_section_qty As Integer

        Dim qs_ul_row_section_avg As Integer

        Dim qs_ul_clm_section_avg As Integer

        Dim qs_ul_row_esat_qty As Integer

        Dim qs_ul_clm_esat_qty As Integer

        Dim temp_int As Integer

        Dim esat_strongly_agree_ct As Integer

        Dim esat_agree_ct As Integer

        Dim esat_neutral_ct As Integer

        Dim esat_disagree_ct As Integer

        Dim esat_strongly_disagree_ct As Integer

        Dim rngTmp As Range

        Dim tmp_num As Integer

        Dim max_column As Integer

        Dim max_row As Integer

        Dim tmp_num2 As Integer

        Dim section_1_column_offset As Integer

        Dim section_2_column_offset As Integer

        Dim section_3_column_offset As Integer

       

        esat_strongly_agree_ct = 0

        esat_agree_ct = 0

        esat_neutral_ct = 0

        esat_disagree_ct = 0

        esat_strongly_disagree_ct = 0

        qs_row_offset = 0

        qs_row_offset = qs_row_offset + 8

        With getWorkSheet("Metric Percentages")

            With .cells(1, 4)

                .Value = "Team:"

                .offset(1, 0).Value = "Time Period:"

                .offset(0, 1).Value = tbxDepartmentName.Value

                .offset(1, 1).Value = primary_output_tab_n

                Call setHeaderFormatting(Range(.offset(0, 0), .offset(1, 0)), getWorkSheet("Metric Percentages"))

 

            End With

            With .Range(.cells(1, 5), .cells(2, 5)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent1

                .TintAndShade = 0.599993896298105

                .PatternTintAndShade = 0

            End With

            With .Range(.cells(1, 4), .cells(2, 5))

                With .Borders

                    .LineStyle = xlContinuous

                End With

            End With

        End With

       

        

        qs_ul_row_metric_lbl_section_qty = 9 + qs_row_offset

        qs_ul_clm_metric_lbl_section_qty = 1

        qs_ul_row_metric_lbl_section_pct = 17 + qs_row_offset

        qs_ul_clm_metric_lbl_section_pct = 1

        qs_ul_row_metric_lbl_metric_qty = 24 + qs_row_offset

        qs_ul_clm_metric_lbl_metric_qty = 1

       

        qs_ul_row_verification = 9 + qs_row_offset

        qs_ul_clm_verification = 16

        qs_ul_row_section_avg = 0

        qs_ul_clm_section_avg = 16

       

        qs_ul_row_eval_scores = 9 + qs_row_offset

        qs_ul_clm_eval_scores = 21

       

        qs_ul_row_esat_qty = 9 + qs_row_offset

        qs_ul_clm_esat_qty = 10

       

        qs_ul_row_verbal = 33 + qs_row_offset

        qs_ul_clm_verbal = 1

        qs_ul_row_business = 33 + qs_row_offset

        qs_ul_clm_business = 6

        qs_ul_row_certification = 33 + qs_row_offset

        qs_ul_clm_certification = 11

        qs_ul_row_survey = 0

        qs_ul_clm_survey = 1

        qs_ul_row_negative = 0

        qs_ul_clm_negative = 6

        qs_ul_row_written = 0

        qs_ul_clm_written = 11

        qs_nr_evaluation_scores = "Evaluation Scores"

        qs_nr_verification_summary = "Verification Summary"

        qs_nr_verbal_evals = "Verbal Evals"

        qs_nr_business_evals = "Business Evals"

        qs_nr_certification_evals = "Certification Evals"

        qs_nr_survey_evals = "Survey Evals"

        qs_nr_negative_evals = "Negative Evals"

        qs_nr_written_evals = "Written Evals"

        qs_nr_metric_lbl_metric_qty = "Quantity of Metric Label per Metric"

        qs_nr_metric_lbl_section_pct = "Percentage of Metric Label per Section"

        qs_nr_metric_lbl_section_qty = "Quantity of Metric Label per Section"

        qs_nr_section_avg = "Agent Average per Section"

        qs_nr_esat_qty = "Evaluator Satisfaction Score Quantity"

        Set temp_sheet = getWorkSheet("Metric Percentages")

       

        temp_sheet.Select

        temp_sheet.Range("A1").Select

       

        ' disable gridlines

        output_book.Activate

        ActiveWindow.DisplayGridlines = False

        ' create tables

        With temp_sheet

            ' section_metric_label_qty As Collection ' Collection with keys for the five sections, Evaluator Satisfaction, Procedural Accuracy, Call Handling, Client Experience, Verification pointing to collection with names of the metrics as keys, pointing to collections with the possible score labels as keys, pointing to integers that represent the count

            ' agents_section_scores As Collection ' with keys for agents pointing to Collection with keys of section names pointing to arrays of metric scores

            Dim sMetricSection As String

            Dim sScoreLabel As String

            Dim sMetricType As String

            If is_client_experience Then

                max_column = 4

            Else

                max_column = 7

            End If

            Set rngTmp = .cells(qs_ul_row_metric_lbl_section_qty, qs_ul_clm_metric_lbl_section_qty)

            With rngTmp

                .Value = qs_nr_metric_lbl_section_qty

                If is_client_experience Then

                    .offset(2, 0).Value = "Meaningful Solutions"

                    .offset(3, 0).Value = "Servicing Skills"

                    .offset(4, 0).Value = "Business Expectations"

                    .offset(5, 0).Value = "Totals"

                    .offset(1, 1).Value = "Yes"

                    .offset(1, 2).Value = "Partial"

                    .offset(1, 3).Value = "No"

                    .offset(1, 4).Value = "N/A"

                Else

                    .offset(2, 0).Value = "Procedural Accuracy"

                    .offset(3, 0).Value = "Call Handling"

                    .offset(4, 0).Value = "Client Experience"

                    .offset(5, 0).Value = "Totals"

                    .offset(1, 1).Value = "Yes"

                    .offset(1, 2).Value = "Partial"

                    .offset(1, 3).Value = "No"

                    .offset(1, 4).Value = "Not Likely"

                    .offset(1, 5).Value = "Likely"

                    .offset(1, 6).Value = "Definitely"

                    .offset(1, 7).Value = "N/A"

                End If

                For tmp_num = 1 To max_column

                    Dim qty_total As Double

                    If .offset(1, tmp_num).Value = "N/A" Then

                   

'                        .offset(2, tmp_num).Value = (UBound(proc_scores) + (1 - LBound(proc_scores))) * 3 - (.offset(2, 1).Value + .offset(2, 2).Value + .offset(2, 3).Value + .offset(2, 4).Value + .offset(2, 5).Value + .offset(2, 6).Value)

'                        .offset(3, tmp_num).Value = (UBound(proc_scores) + (1 - LBound(proc_scores))) * 5 - (.offset(3, 1).Value + .offset(3, 2).Value + .offset(3, 3).Value + .offset(3, 4).Value + .offset(3, 5).Value + .offset(3, 6).Value)

'                        .offset(4, tmp_num).Value = (UBound(proc_scores) + (1 - LBound(proc_scores))) * 4 - (.offset(4, 1).Value + .offset(4, 2).Value + .offset(4, 3).Value + .offset(4, 4).Value + .offset(4, 5).Value + .offset(4, 6).Value)

                    Else

                        sMetricSection = .offset(2, 0).Value

                        qty_total = 0

                        sScoreLabel = .offset(1, tmp_num)

                        For Each metric_type In metric_types

                          sMetricType = metric_type

                          If frmReportBuilderSubmit.getMetricSection(sMetricType) = sMetricSection Then

                            qty_total = qty_total + section_metric_label_qty.getSpecificQty(sMetricSection, sMetricType, sScoreLabel)

                          End If

                        Next metric_type

                        .offset(2, tmp_num).Value = qty_total

                       

                        'Set temp_coll = section_metric_label_qty.Item(.offset(3, 0).Value)

                        sMetricSection = .offset(3, 0).Value

                        qty_total = 0

                        sScoreLabel = .offset(1, tmp_num)

                        For Each metric_type In metric_types

                          sMetricType = metric_type

                          If frmReportBuilderSubmit.getMetricSection(sMetricType) = sMetricSection Then

                            qty_total = qty_total + section_metric_label_qty.getSpecificQty(sMetricSection, sMetricType, sScoreLabel)

                          End If

                        Next metric_type

                        .offset(3, tmp_num).Value = qty_total

                       

                        sMetricSection = .offset(4, 0).Value

                        qty_total = 0

                        sScoreLabel = .offset(1, tmp_num)

                        For Each metric_type In metric_types

                          sMetricType = metric_type

                          If frmReportBuilderSubmit.getMetricSection(sMetricType) = sMetricSection Then

                            qty_total = qty_total + section_metric_label_qty.getSpecificQty(sMetricSection, sMetricType, sScoreLabel)

                          End If

                        Next metric_type

                        .offset(4, tmp_num).Value = qty_total

                    End If

                    .offset(5, tmp_num).Formula = "=SUM(" & getColumnLetterFromNum(.offset(2, tmp_num).Column) & .offset(2, tmp_num).row & ":" & getColumnLetterFromNum(.offset(4, tmp_num).Column) & .offset(4, tmp_num).row & ")"

                Next tmp_num

               

                ' Make totals value from formula

                .Range(.offset(5, 1), .offset(5, max_column)).Copy

                .Range(.offset(5, 1), .offset(5, max_column)).PasteSpecial xlPasteValues

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With ' End rngTmp use

            'output_book.Names.Add NAME:=Replace(qs_nr_metric_lbl_section_qty, " ", ""), RefersToR1C1:=.Range(rngTmp.offset(1, 0), rngTmp.offset(5, 7))

            ' Format table labels

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(5, max_column))

                .HorizontalAlignment = xlLeft

                With .Borders

                    .LineStyle = xlContinuous

                End With

            End With

            With .Range(rngTmp, rngTmp.offset(5, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlTop)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlLeft)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp, rngTmp.offset(5, 0))

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(5, 0), rngTmp.offset(5, max_column))

                With .Borders(xlTop)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(1, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            Call setHeaderFormatting(rngTmp, temp_sheet)

            Call setHeaderFormatting(.Range(rngTmp.offset(1, 0), rngTmp.offset(1, max_column)), temp_sheet)

            Set r = .cells(qs_ul_row_metric_lbl_section_qty + 1, qs_ul_clm_metric_lbl_section_qty)

            .Range(rngTmp, rngTmp.offset(0, max_column)).Merge

            .Range(rngTmp, rngTmp.offset(0, max_column)).HorizontalAlignment = xlCenter

 

            If Not is_client_experience Then

                'ESAT Summary

                Set rngTmp = .cells(qs_ul_row_esat_qty, qs_ul_clm_esat_qty)

                rngTmp.Value = qs_nr_esat_qty

                Call setHeaderFormatting(rngTmp, temp_sheet)

                With rngTmp

                    .offset(1, 0).Value = "Metric"

                    .offset(1, 1).Value = "Qty."

'                    If (Not Not esat_scores) <> 0 Then

'                        For i = UBound(esat_scores) To LBound(esat_scores) Step -1

'                            Select Case esat_scores(i)

'                                Case 5

'                                    esat_strongly_agree_ct = esat_strongly_agree_ct + 1

'                                Case 3.75

'                                    esat_agree_ct = esat_agree_ct + 1

'                                Case 2.5

'                                    esat_neutral_ct = esat_neutral_ct + 1

'                                Case 1.25

'                                    esat_disagree_ct = esat_disagree_ct + 1

'                                Case 0

'                                    esat_strongly_disagree_ct = esat_strongly_disagree_ct + 1

'                            End Select

'                        Next i

'                    End If

                    .offset(2, 0).Value = "Strongly Agree"

                    .offset(2, 1).Value = esat_strongly_agree_ct

                    .offset(3, 0).Value = "Agree"

                    .offset(3, 1).Value = esat_agree_ct

                    .offset(4, 0).Value = "Neutral"

                    .offset(4, 1).Value = esat_neutral_ct

                    .offset(5, 0).Value = "Disagree"

                    .offset(5, 1).Value = esat_disagree_ct

                    .offset(6, 0).Value = "Strongly Disagree"

                    .offset(6, 1).Value = esat_strongly_disagree_ct

                End With

                .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(rngTmp.offset(1, 0), .cells(rngTmp.End(xlDown).row, rngTmp.Column + 1)), _

                                XlListObjectHasheaders:=xlYes).name = Replace(qs_nr_esat_qty, " ", "")

                With .Range(rngTmp, rngTmp.offset(0, 1))

                    .Merge

                    .HorizontalAlignment = xlCenter

                End With

            End If

           

            ' Quality Summary - Metric Label Percentage per Section

            Set rngTmp = .cells(qs_ul_row_metric_lbl_section_pct, qs_ul_clm_metric_lbl_section_pct)

           

            With rngTmp

                .Value = qs_nr_metric_lbl_section_pct

                If is_client_experience Then

                    max_row = 4

                    max_column = 4

                    .offset(2, 0).Value = "Meaningful Solutions"

                    .offset(3, 0).Value = "Servicing Skills"

                    .offset(4, 0).Value = "Business Expectations"

                    .offset(1, 1).Value = "Yes"

                    .offset(1, 2).Value = "Partial"

                    .offset(1, 3).Value = "No"

                    .offset(1, 4).Value = "N/A"

                Else

                    max_row = 4

                    max_column = 7

                    .offset(2, 0).Value = "Procedural Accuracy"

                    .offset(3, 0).Value = "Call Handling"

                    .offset(4, 0).Value = "Client Experience"

                    .offset(1, 1).Value = "Yes"

                    .offset(1, 2).Value = "Partial"

                    .offset(1, 3).Value = "No"

                    .offset(1, 4).Value = "Not Likely"

                    .offset(1, 5).Value = "Likely"

                    .offset(1, 6).Value = "Definitely"

                    .offset(1, 7).Value = "N/A"

                End If

               

                For tmp_num = 1 To max_column

                    For tmp_num2 = 2 To max_row

                        If r.offset(4, tmp_num).Value = 0 Then

                            .offset(tmp_num2, tmp_num).Value = 0

                        Else

                            .offset(tmp_num2, tmp_num).Value = r.offset(tmp_num2 - 1, tmp_num).Value / r.offset(4, tmp_num).Value

                        End If

                    Next tmp_num2

                Next tmp_num

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            .Range(rngTmp.offset(2, 1), rngTmp.offset(max_row, max_column)).NumberFormat = "0.00%"

            .Range(rngTmp.offset(2, 1), rngTmp.offset(max_row, max_column)).HorizontalAlignment = xlLeft

            output_book.Names.Add name:=Replace(qs_nr_metric_lbl_section_pct, " ", ""), RefersToR1C1:=.Range(rngTmp.offset(1, 0), rngTmp.offset(max_row, max_column))

            ' Format table labels

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(max_row, max_column))

                With .Borders

                    .LineStyle = xlContinuous

                End With

            End With

            With .Range(rngTmp, rngTmp.offset(max_row, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlTop)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlLeft)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp, rngTmp.offset(max_row, 0))

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(1, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            Call setHeaderFormatting(rngTmp, temp_sheet)

            Call setHeaderFormatting(.Range(rngTmp.offset(1, 0), rngTmp.offset(1, max_column)), temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, max_column)).Merge

            .Range(rngTmp, rngTmp.offset(0, max_column)).HorizontalAlignment = xlCenter

       

            ' Quality Summary - Metric Label Quantity per Metric

            Set r = .cells(qs_ul_row_metric_lbl_section_qty + 1, qs_ul_clm_metric_lbl_section_qty)

            Set rngTmp = .cells(qs_ul_row_metric_lbl_metric_qty, qs_ul_clm_metric_lbl_metric_qty)

            With rngTmp

                .Value = qs_nr_metric_lbl_metric_qty

                If is_client_experience Then

                    .offset(2, 0).Value = "Yes"

                    .offset(3, 0).Value = "Partial"

                    .offset(4, 0).Value = "No"

                    max_row = 4

                    .offset(1, 1).Value = "Accuracy / Completeness"

                    .offset(1, 2).Value = "Complete Expectations"

                    .offset(1, 3).Value = "Timely Resolution"

                    .offset(1, 4).Value = "World-Class Service"

                    .offset(1, 5).Value = "Hold Experience"

                    .offset(1, 6).Value = "Transfer Experience"

                    .offset(1, 7).Value = "Appropriate Greeting"

                    .offset(1, 8).Value = "Correct Resources"

                    .offset(1, 9).Value = "Appropriate Closing"

                    .offset(1, 10).Value = "Business Processes"

                   max_column = 10

                    section_1_column_offset = 3

                    section_2_column_offset = 6

                    section_3_column_offset = 10

                Else

                    .offset(2, 0).Value = "Yes"

                    .offset(3, 0).Value = "Partial"

                    .offset(4, 0).Value = "No"

                    .offset(5, 0).Value = "Not Likely"

                    .offset(6, 0).Value = "Likely"

                    .offset(7, 0).Value = "Definitely"

                    max_row = 7

                    .offset(1, 1).Value = "Accurate Information"

                    .offset(1, 2).Value = "Process / Procedures"

                    .offset(1, 3).Value = "Expectations"

                    .offset(1, 4).Value = "Hold / Transfer"

                    .offset(1, 5).Value = "Call Log"

                    .offset(1, 6).Value = "Added / Updated"

                    .offset(1, 7).Value = "Survey"

                    .offset(1, 8).Value = "Call Back"

                    .offset(1, 9).Value = "Opening / Farewell"

                    .offset(1, 10).Value = "Actively Listened"

                    .offset(1, 11).Value = "Controlled Call"

                    .offset(1, 12).Value = "Clear / Confident"

                    max_column = 12

                    section_1_column_offset = 3

                    section_2_column_offset = 8

                    section_3_column_offset = 12

                End If

               

                For tmp_num = 1 To max_column

                    qty_total = 0

                    If .offset(1, tmp_num).Value = "Call Back" Then

                        .offset(5, tmp_num).Value = r.offset(2, 4).Value

                        .offset(6, tmp_num).Value = r.offset(2, 5).Value

                        .offset(7, tmp_num).Value = r.offset(2, 6).Value

                    Else

                      For tmp_num2 = 2 To max_row

                        sMetricSection = frmReportBuilderSubmit.getMetricSection(.offset(1, tmp_num).Value)

                        sMetricType = .offset(1, tmp_num).Value

                        sScoreLabel = .offset(tmp_num2, 0).Value

                        qty_total = section_metric_label_qty.getSpecificQty(sMetricSection, sMetricType, sScoreLabel)

                        .offset(tmp_num2, tmp_num).Value = qty_total

                      Next tmp_num2

                    End If

                Next tmp_num

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            output_book.Names.Add name:=Replace(qs_nr_metric_lbl_metric_qty, " ", ""), RefersToR1C1:=.Range(rngTmp.offset(1, 0), rngTmp.offset(max_row, max_column))

           

            ' Format table labels

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(max_row, max_column))

                .HorizontalAlignment = xlLeft

                With .Borders

                    .LineStyle = xlContinuous

                End With

            End With

            With .Range(rngTmp, rngTmp.offset(max_row, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlTop)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

                With .Borders(xlLeft)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(max_row, 0))

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(1, max_column))

                With .Borders(xlBottom)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, section_1_column_offset), rngTmp.offset(max_row, section_1_column_offset))

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            With .Range(rngTmp.offset(1, section_2_column_offset), rngTmp.offset(max_row, section_2_column_offset))

                With .Borders(xlRight)

                    .LineStyle = xlContinuous

                    .Weight = xlThick

                End With

            End With

            If ((max_row - 1) / 2) > 2 Then

                With .Range(rngTmp.offset(((max_row - 1) / 2) + 1, 0), rngTmp.offset(4, max_column))

                    With .Borders(xlBottom)

                        .LineStyle = xlContinuous

                        .Weight = xlThick

                    End With

                End With

            End If

            Call setHeaderFormatting(rngTmp, temp_sheet)

            Call setHeaderFormatting(.Range(rngTmp.offset(1, 0), rngTmp.offset(1, section_3_column_offset)), temp_sheet)

            With .Range(rngTmp.offset(1, 0), rngTmp.offset(1, section_1_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent1

                .TintAndShade = 0.399975585192419

                .PatternTintAndShade = 0

            End With

            With .Range(rngTmp.offset(2, 1), rngTmp.offset(max_row, section_1_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent1

                .TintAndShade = 0.599993896298105

                .PatternTintAndShade = 0

            End With

            With .Range(rngTmp.offset(1, section_1_column_offset + 1), rngTmp.offset(1, section_2_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent5

                .TintAndShade = 0.399975585192419

                .PatternTintAndShade = 0

            End With

            With .Range(rngTmp.offset(2, section_1_column_offset + 1), rngTmp.offset(max_row, section_2_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent5

                .TintAndShade = 0.599993896298105

                .PatternTintAndShade = 0

            End With

            With .Range(rngTmp.offset(1, section_2_column_offset + 1), rngTmp.offset(1, section_3_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent4

                .TintAndShade = 0.399975585192419

                .PatternTintAndShade = 0

            End With

            With .Range(rngTmp.offset(2, section_2_column_offset + 1), rngTmp.offset(max_row, section_3_column_offset)).Interior

                .Pattern = xlSolid

                .PatternColorIndex = xlAutomatic

                .ThemeColor = xlThemeColorAccent4

                .TintAndShade = 0.599993896298105

                .PatternTintAndShade = 0

            End With

            .Range(rngTmp, rngTmp.offset(0, section_3_column_offset)).Merge

            .Range(rngTmp, rngTmp.offset(0, section_3_column_offset)).HorizontalAlignment = xlCenter

           

            ' Verbal evaluation average

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_verbal, qs_ul_clm_verbal, qs_nr_verbal_evals, "verbal", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_verbal, qs_ul_clm_verbal)

            On Error GoTo NoVerbalBod

            If Not IsEmpty(.ListObjects.Item(Replace(qs_nr_verbal_evals, " ", "")).ListColumns.Item(1).DataBodyRange.offset(1, 0).Value) Then

                qs_ul_row_survey = .ListObjects.Item(Replace(qs_nr_verbal_evals, " ", "")).ListColumns.Item(1).DataBodyRange.End(xlDown).row + 2

            Else

NoVerbalBod:

                Err.Clear

                qs_ul_row_survey = rngTmp.row + 4

            End If

            On Error GoTo 0

           

            ' Format table labels

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

           

            ' Survey evaluation averages

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_survey, qs_ul_clm_survey, qs_nr_survey_evals, "survey", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_survey, qs_ul_clm_survey)

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

       

            ' Business evaluation averages

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_business, qs_ul_clm_business, qs_nr_business_evals, "business", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_business, qs_ul_clm_business)

            On Error GoTo NoBusBod

            If getEvalTypeCount("Business") > 0 Then

                qs_ul_row_negative = .ListObjects.Item(Replace(qs_nr_business_evals, " ", "")).ListColumns(1).DataBodyRange.End(xlDown).row + 2

            Else

NoBusBod:

                Err.Clear

                qs_ul_row_negative = rngTmp.row + 4

            End If

            On Error GoTo 0

            ' Format table labels

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

           

            ' Negative evaluation averages

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_negative, qs_ul_clm_negative, qs_nr_negative_evals, "negative", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_negative, qs_ul_clm_negative)

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

           

            ' Certification evaluation averages

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_certification, qs_ul_clm_certification, qs_nr_certification_evals, "certification", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_certification, qs_ul_clm_certification)

            On Error GoTo NoCertBod

            If getEvalTypeCount("Certification") > 0 Then

                qs_ul_row_written = .ListObjects.Item(Replace(qs_nr_certification_evals, " ", "")).ListColumns(1).DataBodyRange.End(xlDown).row + 2

            Else

NoCertBod:

                Err.Clear

                qs_ul_row_written = rngTmp.row + 4

            End If

            On Error GoTo 0

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

           

            ' Written evaluation averages

            Call fillEvalAvgTable(output_book.Worksheets("Metric Percentages"), qs_ul_row_written, qs_ul_clm_written, qs_nr_written_evals, "written", oTheEvaluations)

            Set rngTmp = .cells(qs_ul_row_written, qs_ul_clm_written)

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 3)).Merge

            .Range(rngTmp, rngTmp.offset(0, 3)).HorizontalAlignment = xlCenter

 

            ' Verification table

            With .cells(qs_ul_row_verification, qs_ul_clm_verification)

                .Value = qs_nr_verification_summary

                .offset(1, 0).Value = "Agent"

                .offset(1, 1).Value = "Evaluations"

                .offset(1, 2).Value = "Failed Verification"

                Dim temp_text As String

                Dim qty_othertotal As Double

                current_agent_offset = 1

                For sheets_i = 1 To output_book.Worksheets.Count

                    With output_book.Worksheets(sheets_i)

                        temp_text = .name

                    End With

                    If isCollectionKey(temp_text, oAgentsEvalScores) Then

                        qty_total = 0

                        qty_othertotal = 0

                        current_agent_offset = current_agent_offset + 1

                        temp_evalarray = oAgentsEvalScores(temp_text)

                        For tmp_num = UBound(temp_evalarray) To LBound(temp_evalarray) Step -1

                            With temp_evalarray(tmp_num)

                                If Not .everification Then

                                    qty_othertotal = qty_othertotal + 1

                                End If

                                qty_total = qty_total + 1

                            End With

                        Next tmp_num

                        .offset(current_agent_offset, 0).Value = temp_text

                        .offset(current_agent_offset, 1).Value = qty_total

                        .offset(current_agent_offset, 2).Value = qty_othertotal

                    End If

                Next sheets_i

            End With

            With .Range(.cells(qs_ul_row_verification, qs_ul_clm_verification).offset(2, 2), .cells(qs_ul_row_verification, qs_ul_clm_verification).offset(current_agent_offset, 2))

                .FormatConditions.Add Type:=xlExpression, Formula1:= _

                    "=" & getColumnLetterFromNum(qs_ul_clm_verification + 2) & (qs_ul_row_verification + 2) & ">0"

                .FormatConditions(.FormatConditions.Count).SetFirstPriority

                With .FormatConditions(1).Interior

                    .PatternColorIndex = xlAutomatic

                    .Color = 255

                    .TintAndShade = 0

                End With

                .FormatConditions(1).StopIfTrue = False

            End With

            Set rngTmp = .cells(qs_ul_row_verification, qs_ul_clm_verification)

            .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(rngTmp.offset(1, 0), cells(rngTmp.End(xlDown).row, rngTmp.Column + 2)), _

                                XlListObjectHasheaders:=xlYes).name = Replace(qs_nr_verification_summary, " ", "")

            On Error GoTo NoVerification

            If Not IsEmpty(.ListObjects.Item(Replace(qs_nr_verification_summary, " ", "")).ListColumns(1).DataBodyRange.offset(1, 0).Value) Then

                qs_ul_row_section_avg = .ListObjects.Item(Replace(qs_nr_verification_summary, " ", "")).ListColumns(1).DataBodyRange.End(xlDown).row + 2

                .ListObjects.Item(Replace(qs_nr_verification_summary, " ", "")).DataBodyRange.HorizontalAlignment = xlLeft

            Else

NoVerification:

                Err.Clear

                qs_ul_row_section_avg = rngTmp.row + 4

            End If

            On Error GoTo 0

            ' Format table labels

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, 2)).Merge

            .Range(rngTmp, rngTmp.offset(0, 2)).HorizontalAlignment = xlCenter

           

            ' Agent averages per evaluation card section

            With .cells(qs_ul_row_section_avg, qs_ul_clm_section_avg)

                .Value = qs_nr_section_avg

                If is_client_experience Then

                    .offset(1, 0).Value = "Agent"

                    .offset(1, 1).Value = "Meaningful Solutions"

                    .offset(1, 2).Value = "Servicing Skills"

                    .offset(1, 3).Value = "Business Expectations"

                    max_column = 3

                Else

                    .offset(1, 0).Value = "Agent"

                    .offset(1, 1).Value = "Procedural Accuracy"

                    .offset(1, 2).Value = "Call Handling"

                    .offset(1, 3).Value = "Client Experience"

                    max_column = 3

                End If

                current_agent_offset = 1

                For sheets_i = 1 To output_book.Worksheets.Count

                    With output_book.Worksheets(sheets_i)

                        temp_text = .name

                    End With

                    If isCollectionKey(temp_text, oAgentsEvalScores) Then

                        current_agent_offset = current_agent_offset + 1

                        For tmp_num2 = 1 To max_column

                            .offset(current_agent_offset, 0).Value = temp_text

                            ' Procedural Accuracy average

                            .offset(current_agent_offset, tmp_num2).Value = agents_section_scores.getSpecificScoreAverage(temp_text, .offset(1, tmp_num2).Value)

                        Next tmp_num2

                    End If

                Next sheets_i

            End With

            For sheets_i = qs_ul_clm_section_avg To (qs_ul_clm_section_avg + max_column)

                .Columns(getColumnLetterFromNum(sheets_i)).AutoFit

            Next sheets_i

            Set rngTmp = .cells(qs_ul_row_section_avg, qs_ul_clm_section_avg)

            .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(rngTmp.offset(1, 0), cells(rngTmp.End(xlDown).row, rngTmp.Column + max_column)), _

                                XlListObjectHasheaders:=xlYes).name = Replace(qs_nr_section_avg, " ", "")

            On Error GoTo NoSectionAvgTbl

            If oAgentsEvalScores.Count > 0 Then

                If .ListObjects.Item(Replace(qs_nr_section_avg, " ", "")).DataBodyRange.Count > 0 Then

                    .ListObjects.Item(Replace(qs_nr_section_avg, " ", "")).DataBodyRange.HorizontalAlignment = xlLeft

                End If

            End If

NoSectionAvgTbl:

            Err.Clear

            On Error GoTo 0

            ' Format table labels

            Call setHeaderFormatting(rngTmp, temp_sheet)

            .Range(rngTmp, rngTmp.offset(0, max_column)).Merge

            .Range(rngTmp, rngTmp.offset(0, max_column)).HorizontalAlignment = xlCenter

 

            ' Evaluation table

            With .cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores)

                Dim temp_clmn_offset As Integer

                Dim temp_esat_eval() As EvalProcEsatTypeDate

                Dim proc_all_total As Double

               Dim esat_all_total As Double

                Dim avgs_to_avg As Variant

                Dim unique_eval_ct As Integer

                .Value = qs_nr_evaluation_scores

                .offset(1, 0).Value = "Agent"

                For sheets_i = 1 To eval_qty_max

                    .offset(1, sheets_i).Value = "Evaluation " & sheets_i

                Next sheets_i

                .offset(1, 0 + sheets_i).Value = getPrimaryScoreName()

                .offset(1, 1 + sheets_i).Value = getSecondaryScoreName()

                current_agent_offset = 1

                proc_all_total = 0

                esat_all_total = 0

                unique_eval_ct = 0

                For sheets_i = 1 To output_book.Worksheets.Count

                    With output_book.Worksheets(sheets_i)

                        temp_text = .name

                    End With

                    If isCollectionKey(temp_text, oAgentsEvalScores) Then

                        current_agent_offset = current_agent_offset + 1

                        temp_clmn_offset = 0

                        proc_sub_total = 0

                        esat_sub_total = 0

                        .offset(current_agent_offset, 0).Value = temp_text

                        temp_evalarray = oAgentsEvalScores(temp_text)

                        temp_clmn_offset = LBound(temp_evalarray) - 1

                        For tmp_num = LBound(temp_evalarray) To UBound(temp_evalarray)

                            Set temp_eval = temp_evalarray(tmp_num)

                            .offset(current_agent_offset, tmp_num - temp_clmn_offset).Value = temp_eval.procedural

                            proc_sub_total = proc_sub_total + temp_eval.procedural

                            proc_all_total = proc_all_total + temp_eval.procedural

                            unique_eval_ct = unique_eval_ct + 1

                            If temp_eval.SecondaryValid Then

                                esat_sub_total = esat_sub_total + temp_eval.esat

                                esat_all_total = esat_all_total + temp_eval.esat

                            End If

                        Next tmp_num

                        .offset(current_agent_offset, eval_qty_max + 1).Value = proc_sub_total / CDbl(UBound(temp_evalarray) + (1 - LBound(temp_evalarray)))

                        If truncate_numbers Then

                            .offset(current_agent_offset, eval_qty_max + 1).NumberFormat = "0.00"

                        End If

                        If oTheEvaluations.isSecondaryAvgAvailableForAgent(temp_text) Then

                            .offset(current_agent_offset, eval_qty_max + 2).Value = esat_sub_total / CDbl(UBound(temp_evalarray) + (1 - LBound(temp_evalarray)))

                            If truncate_numbers Then

                                .offset(current_agent_offset, eval_qty_max + 2).NumberFormat = "0.00"

                            End If

                        End If

                    End If

                Next sheets_i

               

                .offset(current_agent_offset + 1, 0).Value = "Total Average"

                'avgs_to_avg

                .offset(current_agent_offset + 1, eval_qty_max + 1).Value = proc_all_total / unique_eval_ct

                If truncate_numbers Then

                    .offset(current_agent_offset + 1, eval_qty_max + 1).NumberFormat = "0.00"

                End If

                If esat_qty_deficiency = 0 Then

                    .offset(current_agent_offset + 1, eval_qty_max + 2).Value = esat_all_total / unique_eval_ct

                    If truncate_numbers Then

                        .offset(current_agent_offset + 1, eval_qty_max + 2).NumberFormat = "0.00"

                    End If

                End If

           End With

           

            ' conditional formatting

            Set rngTmp = .Range(.cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores).offset(2, eval_qty_max + 1), .cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores).offset(current_agent_offset, eval_qty_max + 1))

            temp_text = getColumnLetterFromNum(qs_ul_clm_eval_scores + eval_qty_max + 1) & (qs_ul_row_eval_scores + 2)

            Call processPrimaryScoreQuintileFormatting(rngTmp, temp_text)

           

            ' ESAT conditional formatting

            Call processSecondaryScoreQuintileFormatting(.Range(.cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores).offset(2, eval_qty_max + 2), .cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores).offset(current_agent_offset, eval_qty_max + 2)), getColumnLetterFromNum(qs_ul_clm_eval_scores + eval_qty_max + 2) & (qs_ul_row_eval_scores + 2))

           

            

            Set rngTmp = .cells(qs_ul_row_eval_scores, qs_ul_clm_eval_scores)

            ' Use tmp_num for horizontal offset: number of evaluations for the most evaluated SP plus Procedural column plus ESAT column

            tmp_num = eval_qty_max + 2

            .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(rngTmp.offset(1, 0), cells(rngTmp.End(xlDown).row, rngTmp.Column + tmp_num)), _

                                XlListObjectHasheaders:=xlYes).name = Replace(qs_nr_evaluation_scores, " ", "")

            .ListObjects.Item(Replace(qs_nr_evaluation_scores, " ", "")).DataBodyRange.HorizontalAlignment = xlLeft

            ' Format table labels

            Call setHeaderFormatting(rngTmp, temp_sheet)

            Range(rngTmp, rngTmp.offset(0, tmp_num)).Merge

            Range(rngTmp, rngTmp.offset(0, tmp_num)).HorizontalAlignment = xlCenter

            .Columns(getColumnLetterFromNum(qs_ul_clm_eval_scores)).AutoFit

           

        End With

    End If

    ' Close process

    Application.ScreenUpdating = False

    Unload frmReportBuilderSubmit

End Sub

 

Public Function formatSmName(sRawSmName As String) As String

  If Not InStr(1, sRawSmName, "(SM) ") = 1 Then

    formatSmName = "(SM) " & sRawSmName

  End If

End Function

 

Private Function getPrimaryAvgForSm(sSmName As String) As Double

  Dim aTempDoubles() As Double

  aTempDoubles = sm_proc_scores(sSmName)

  getPrimaryAvgForSm = frmReportBuilderSubmit.getDoubleArrayAverage(aTempDoubles)

End Function

 

Private Function getSecondaryAvgForSm(sSmName As String) As Double

  Dim aTempDoubles() As Double

  aTempDoubles = sm_proc_scores(sSmName)

  If isSecondaryAvgAvailableForSm(sSmName) Then

    getSecondaryAvgForSm = frmReportBuilderSubmit.getDoubleArrayAverage(aTempDoubles)

  Else

    getSecondaryAvgForSm = 0

  End If

End Function

 

Private Function isSecondaryAvgAvailableForSm(sSmName As String) As Boolean

  isSecondaryAvgAvailableForSm = frmReportBuilderSubmit.getArrayLength(sm_proc_scores(sSmName)) = frmReportBuilderSubmit.getArrayLength(sm_esat_scores(sSmName))

End Function

 

Private Sub addSmPrimaryScore(ByVal score As Double, ByVal sSmName As String)

  Dim aTempScores() As Double

  If frmReportBuilderSubmit.isKeyOfCollection(sm_proc_scores, sSmName) Then

    aTempScores = sm_proc_scores(sSmName)

    sm_proc_scores.Remove (sSmName)

    ReDim Preserve aTempScores(LBound(aTempScores) To UBound(aTempScores) + 1)

  Else

    ReDim aTempScores(0 To 0)

  End If

  aTempScores(UBound(aTempScores)) = score

  sm_proc_scores.Add key:=sSmName, Item:=aTempScores

End Sub

 

Private Sub addSmSecondaryScore(ByVal score As Double, ByVal sSmName As String)

  Dim aTempScores() As Double

  If frmReportBuilderSubmit.isKeyOfCollection(sm_esat_scores, sSmName) Then

    aTempScores = sm_esat_scores(sSmName)

    sm_esat_scores.Remove (sSmName)

    ReDim Preserve aTempScores(LBound(aTempScores) To UBound(aTempScores) + 1)

  Else

    ReDim aTempScores(0 To 0)

  End If

  aTempScores(UBound(aTempScores)) = score

  sm_esat_scores.Add key:=sSmName, Item:=aTempScores

End Sub

 

Public Function getDoubleArrayAverage(aDoubles() As Double) As Double

  Dim i As Integer

  Dim subtotal As Double

  On Error Resume Next

  If (Not Not aDoubles) = 0 Or UBound(aDoubles) = -1 Then

    getDoubleArrayAverage = 0

  Else

    For i = UBound(aDoubles) To LBound(aDoubles) Step -1

      subtotal = subtotal + aDoubles(i)

    Next i

    getDoubleArrayAverage = subtotal / (UBound(aDoubles) - LBound(aDoubles) + 1)

  End If

  On Error GoTo 0

End Function

 

Public Function getArrayLength(aArrayToBeMeasured As Variant) As Integer

  On Error Resume Next

  If (Not Not aArrayToBeMeasured) = 0 Or UBound(aArrayToBeMeasured) = -1 Then

    getArrayLength = 0

  Else

    getArrayLength = (UBound(aArrayToBeMeasured) - LBound(aArrayToBeMeasured) + 1)

  End If

  On Error GoTo 0

End Function

 

Public Sub incrementDoubleArrayLength(ByRef aOneDimensionalArray() As Double)

  If (Not Not aOneDimensionalArray) = 0 Then

    ReDim aOneDimensionalArray(0 To 0)

  Else

    ReDim Preserve aOneDimensionalArray(LBound(aOneDimensionalArray) To UBound(aOneDimensionalArray) + 1)

  End If

End Sub

 

Private Sub addGarbageMetadataRow(ByVal sAgentName As String, ByVal dTimeStamp As Date, ByVal sEvalType As String, ByVal sEvalComment As String, _

    ByVal sMetricType As String, ByVal sScoreLabel As String, ByVal dScoreMax As Double, _

    ByVal dPrimaryScore As Double, Optional ByVal dSecondaryScore As Double)

  Dim revised_mtype As String

  Dim sGarbageMetadataTabName As String

  If Len(sMetricType) = 0 Then

    revised_mtype = "OTHER-UNKNOWN"

  Else

    revised_mtype = sMetricType

  End If

  sGarbageMetadataTabName = "Bad Format Rows"

  Call addAgentName(garbage_row_offset, sAgentName, sGarbageMetadataTabName)

  Call addEvalType(garbage_row_offset, sEvalType, sGarbageMetadataTabName)

  Call addTimeStamp(garbage_row_offset, dTimeStamp, sGarbageMetadataTabName)

  Call addMetricScoreLabel(garbage_row_offset, sScoreLabel, sGarbageMetadataTabName)

  Call addMetricMax(garbage_row_offset, dScoreMax, sGarbageMetadataTabName)

  Call addProcScore(garbage_row_offset, dPrimaryScore, sGarbageMetadataTabName)

  Call addEsatScore(garbage_row_offset, dSecondaryScore, sGarbageMetadataTabName)

  Call addComment(garbage_row_offset, sEvalComment, sGarbageMetadataTabName)

  Call addMetricType(garbage_row_offset, revised_mtype, sGarbageMetadataTabName)

  garbage_row_offset = garbage_row_offset + 1

End Sub

 

Private Sub processPrimaryScoreQuintileFormatting(ByRef formatting_range As Range, ByVal formula_starting_target_cell As String)

    ' Primary Score Quintile thresholds

    Dim five_bottom_prime As String

    Dim four_top_prime As String

    Dim four_bottom_prime As String

    Dim three_top_prime As String

    Dim three_bottom_prime As String

    Dim two_top_prime As String

    Dim two_bottom_prime As String

    Dim one_top_prime As String

   

    If is_client_experience Then

        five_bottom_prime = "4.99999999"

        four_bottom_prime = "4.66999999"

        four_top_prime = "5"

        three_top_prime = "4.67"

        three_bottom_prime = "3.44999999"

        two_bottom_prime = "2.48999999"

        two_top_prime = "3.45"

        one_top_prime = "2.49"

    Else

        five_bottom_prime = "4.74999999"

        four_bottom_prime = "4.39999999"

        four_top_prime = "4.75"

        three_bottom_prime = "4.4"

        three_top_prime = "3.89999999"

        two_bottom_prime = "2.79999999"

        two_top_prime = "3.9"

        one_top_prime = "2.49"

    End If

 

    ' Procedural score formatting

    With formatting_range

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=" & formula_starting_target_cell & ">" & five_bottom_prime

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Font

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

        End With

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent4

            .TintAndShade = -0.249946592608417

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" & four_bottom_prime & "," & formula_starting_target_cell & "<" & four_top_prime & ")"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent4

            .TintAndShade = 0.599963377788629

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" & three_bottom_prime & "," & formula_starting_target_cell & "<" & three_top_prime & ")"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent5

            .TintAndShade = 0.599963377788629

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" & two_bottom_prime & "," & formula_starting_target_cell & "<" & two_top_prime & ")"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .Color = 255

            .TintAndShade = 0

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & formula_starting_target_cell & "<" & one_top_prime

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Font

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

       End With

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .Color = 192

            .TintAndShade = 0

        End With

        .FormatConditions(1).StopIfTrue = False

    End With

End Sub

 

Private Sub processSecondaryScoreQuintileFormatting(ByRef formatting_range As Range, ByVal formula_starting_target_cell As String)

    ' Secondary Score Quintile thresholds

    Dim five_bottom_second As String

    Dim four_top_second As String

    Dim four_bottom_second As String

    Dim three_top_second As String

    Dim three_bottom_second As String

    Dim two_top_second As String

    Dim two_bottom_second As String

    Dim one_top_second As String

   

    If is_client_experience Then

        five_bottom_second = "4.99999999"

        four_bottom_second = "4.66999999"

        four_top_second = "5"

        three_bottom_second = "3.59999999"

        three_top_second = "4.67"

        two_bottom_second = "2.99999999"

        two_top_second = "3.6"

        one_top_second = "3"

    Else

        five_bottom_second = "4.74999999"

        four_bottom_second = "4.24999999"

        four_top_second = "4.75"

        three_top_second = "4.25"

        three_bottom_second = "3.49999999"

        two_bottom_second = "2.99999999"

        two_top_second = "3.49"

        one_top_second = "3"

    End If

 

    With formatting_range

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" & five_bottom_second & ",LEN(" & formula_starting_target_cell & ")>0)"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Font

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

        End With

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent4

            .TintAndShade = -0.249946592608417

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" & four_bottom_second & "," & formula_starting_target_cell & "<" & four_top_second & ",LEN(" & formula_starting_target_cell & ")>0)"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent4

            .TintAndShade = 0.599963377788629

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & "<" & three_top_second & "," & formula_starting_target_cell & ">" & three_bottom_second & ",LEN(" & formula_starting_target_cell & ")>0)"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent5

            .TintAndShade = 0.599963377788629

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:= _

            "=AND(" & formula_starting_target_cell & ">" + two_bottom_second + "," & formula_starting_target_cell & "<" & two_top_second & ",LEN(" & formula_starting_target_cell & ")>0)"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .Color = 255

            .TintAndShade = 0

        End With

        .FormatConditions(1).StopIfTrue = False

        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & formula_starting_target_cell & "<" + one_top_second + ",LEN(" & formula_starting_target_cell & ")>0)"

        .FormatConditions(.FormatConditions.Count).SetFirstPriority

        With .FormatConditions(1).Font

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

        End With

        With .FormatConditions(1).Interior

            .PatternColorIndex = xlAutomatic

            .Color = 192

            .TintAndShade = 0

        End With

        .FormatConditions(1).StopIfTrue = False

    End With

End Sub

 

Private Function getWorkSheet(ByVal ws_name As String) As Worksheet

  Dim this_ws As Worksheet

  If Len(ws_name) > 0 Then

    If Not SheetExists(ws_name, output_book) Then

      Set this_ws = output_book.Worksheets.Add

      this_ws.name = ws_name

      Call arrangeWorksheetOrder(ws_name)

    End If

    Set getWorkSheet = output_book.Worksheets(ws_name)

  End If

End Function

 

Public Sub arrangeWorksheetOrder(sWsName As String)

  With output_book.Worksheets(sWsName)

    If sWsName = "Bad Format Rows" Then

      .Move before:=output_book.Worksheets(1)

    ElseIf InStr(1, sWsName, "(SM) ") = 0 Then

      .Move before:=output_book.Worksheets(getFirstSmSheetIndex())

    ElseIf Not frmReportBuilderSubmit.containsNonNumericCharacters(sWsName) Then

      .Move before:=output_book.Worksheets(output_book.Worksheets.Count)

    Else

      .Move before:=output_book.Worksheets(1)

    End If

  End With

End Sub

 

Public Function getFirstSmSheetIndex() As Integer

  Dim i As Integer

  Dim oWorkingWs As Worksheet

  For i = 1 To output_book.Worksheets.Count

    Set oWorkingWs = output_book.Worksheets(i)

    If InStr(1, oWorkingWs.name, "(SM) ") > 0 Then

      getFirstSmSheetIndex = i

      Exit Function

    End If

  Next i

  getFirstSmSheetIndex = output_book.Worksheets.Count

End Function

 

Public Sub addCollectionStringKeyAndIncrementIntegerValue(ByRef oTheCollection As Collection, sTheKey As String)

  Dim iCounter As Integer

  If frmReportBuilderSubmit.isKeyOfCollection(oTheCollection, sTheKey) Then

    iCounter = oTheCollection(sTheKey)

    oTheCollection.Remove (sTheKey)

    oTheCollection.Add key:=sTheKey, Item:=(iCounter + 1)

  Else

    oTheCollection.Add key:=sTheKey, Item:=1

  End If

End Sub

 

'

'Private Function getLogicSwitchToken() As Boolean

'    If Not first_metric Then

'        getLogicSwitchToken = Not first_metric

'    Else

'        If (Not Not proc_scores) <> 0 Then ' proc_scores is  not empty, check to see if there is more than one entry

'            getLogicSwitchToken = UBound(proc_scores) > LBound(proc_scores)

'        Else

'            getLogicSwitchToken = Not first_metric ' False

'        End If

'    End If

'End Function

 

'Private Sub processDefaultVerification()

'    Dim temp_cell_value As Variant

'    Dim temp_evalarray() As EvalProcEsatTypeDate

'    Dim temp_eval As EvalProcEsatTypeDate

'    Dim temp_coll As Collection

'    Dim temp_coll2 As Collection

'    Dim i As Integer

'    If Not verification_comment_supplied Then

'        output_row_offset = output_row_offset + 1

'        current_agent_offset = current_agent_offset + 1

'        If sm_known Then

'            current_sm_offset = current_sm_offset + 1

'        End If

'        temp_cell_value = getAgentName(output_row_offset - 1, primary_output_tab_n)

'        Call addAgentNameOmnibus(temp_cell_value)

'        temp_cell_value = "Verification"

'        Call addMetricTypeOmnibus(temp_cell_value)

'        temp_cell_value = getEvalType(output_row_offset - 1, primary_output_tab_n)

'        Call addEvalTypeOmnibus(temp_cell_value)

'        temp_cell_value = getTimeStamp(output_row_offset - 1, primary_output_tab_n)

'        Call addTimeStampOmnibus(temp_cell_value)

'        temp_cell_value = "Yes"

'        Call addMetricScoreLabelOmnibus(temp_cell_value)

'        temp_cell_value = 1

'        Call addMetricScoreOmnibus(temp_cell_value)

'        Call addMetricMaxOmnibus(temp_cell_value)

'        Call addMetricPercentageOmnibus(temp_cell_value)

'        temp_evalarray = oAgentsEvalScores(current_agent)

'        Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'        temp_cell_value = temp_eval.procedural

'        oAgentsEvalScores.Remove (current_agent)

'        temp_eval.everification = True

'        Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'        oAgentsEvalScores.Add key:=current_agent, Item:=temp_evalarray

'        Call addProceduralScoreOmnibus(temp_cell_value)

'        If oAgentsEvalScores.Count = oAgentsEvalScores.Count Then

'            temp_evalarray = oAgentsEvalScores(current_agent)

'            Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'            temp_cell_value = temp_eval.esat

'            Call addEsatScoreOmnibus(temp_cell_value)

'        ElseIf Not IsEmpty(getEsatScore(output_row_offset - 1, primary_output_tab_n)) Then

'            temp_cell_value = getEsatScore(output_row_offset - 1, primary_output_tab_n)

'            Call addEsatScoreOmnibus(temp_cell_value)

'        End If

'

'        If isCollectionKey("Verification", section_metric_label_qty) Then

'            Set temp_coll = section_metric_label_qty("Verification")

'            section_metric_label_qty.Remove ("Verification")

'        Else

'            Set temp_coll = New Collection

'        End If

'        If isCollectionKey("Verification", temp_coll) Then

'            Set temp_coll2 = temp_coll("Verification")

'            temp_coll.Remove ("Verification")

'        Else

'            Set temp_coll2 = New Collection

'        End If

'        If isCollectionKey("Yes", temp_coll2) Then

'            i = temp_coll2("Yes")

'            temp_coll2.Remove ("Yes")

'        Else

'            i = 0

'        End If

'        temp_coll2.Add key:="Yes", Item:=i + 1

'        temp_coll.Add key:="Verification", Item:=temp_coll2

'        section_metric_label_qty.Add key:="Verification", Item:=temp_coll

'    End If

'End Sub

 

'Private Sub processMissingEsat()

'    Dim temp_bdate As Variant ' using as array

'    Dim temp_adate As Variant ' using as array

'

'    If (Not Not proc_scores) <> 0 Then

'        If (Not Not esat_scores) = 0 Then

'            If isCollectionKey(current_agent, evals_missing_esat) Then

'                temp_bdate = evals_missing_esat(current_agent)

'                evals_missing_esat.Remove (current_agent)

'                ReDim Preserve temp_bdate(LBound(temp_bdate) To UBound(temp_bdate) + 1)

'            Else

'                ReDim temp_bdate(0 To 0)

'            End If

'            esat_qty_deficiency = esat_qty_deficiency + 1

'            temp_adate = sp_evaldate_collection(current_agent)

'            temp_bdate(UBound(temp_bdate)) = temp_adate(UBound(temp_adate))

'            evals_missing_esat.Add key:=current_agent, Item:=temp_bdate

'            If Not is_client_experience Then

'                Call addEsatDeficitEntry

'            End If

'        ElseIf isCollectionKey(current_agent, sp_evaldate_collection) Then

'            If isCollectionKey(current_agent, evals_missing_esat) Then

'                temp_bdate = evals_missing_esat(current_agent)

'                evals_missing_esat.Remove (current_agent)

'                ReDim Preserve temp_bdate(LBound(temp_bdate) To UBound(temp_bdate) + 1)

'            Else

'                ReDim temp_bdate(0 To 0)

'            End If

'            esat_qty_deficiency = esat_qty_deficiency + 1

'            temp_adate = sp_evaldate_collection(current_agent)

'            temp_bdate(UBound(temp_bdate)) = temp_adate(UBound(temp_adate))

'            evals_missing_esat.Add key:=current_agent, Item:=temp_bdate

'            If Not is_client_experience Then

'                Call addEsatDeficitEntry

'            End If

'        End If

'    End If

 

Private Sub incrementOffsetOmnibus()

    output_row_offset = output_row_offset + 1

    current_agent_offset = current_agent_offset + 1

    If sm_known Then

        current_sm_offset = current_sm_offset + 1

    End If

End Sub

 

Private Sub applyHeaderColor(ByRef r As Range)

    With r.Interior

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

End Sub

 

Private Sub applyHeaderFontStyle(ByRef r As Range)

    With r.Font

        .ThemeColor = xlThemeColorDark1

        .TintAndShade = 0

    End With

    With r.Font

        .name = "Tahoma"

        .Size = 10

        .Strikethrough = False

        .Superscript = False

        .Subscript = False

        .OutlineFont = False

        .Shadow = False

        .Underline = xlUnderlineStyleNone

        .ThemeColor = xlThemeColorDark1

        .TintAndShade = 0

        .ThemeFont = xlThemeFontNone

    End With

    r.Font.Bold = True

End Sub

 

Private Function getPrimaryScoreName()

    If is_client_experience Then

        getPrimaryScoreName = "Client Satisfaction"

    Else

        getPrimaryScoreName = "Procedural"

    End If

End Function

 

Private Function getSecondaryScoreName()

    If is_client_experience Then

        getSecondaryScoreName = "Business Expectations"

    Else

        getSecondaryScoreName = "ESAT"

    End If

End Function

 

Private Sub addQSProceduralEval(ByVal agent_name As String, ByVal score As Double, eval_num As Integer)

    Dim i As Integer

    Dim latest_month As Integer

    Dim raw_date As String

    Dim prev_month_clmn_n As String

    Dim qr_cell As Range

    Dim found_later As Boolean

    Dim eval_clmn_n As String

    eval_clmn_n = "Evaluation " & eval_num

    i = 1

    found_later = False

    With output_book.Worksheets("Quality Ranking")

        If Not isTableColumnName(eval_clmn_n, qs_nr_evaluation_scores, "Metric Percentages") Then

            latest_month = getTableClmnNumber(getPrimaryScoreName(), qr_proc_named_range_n)

            .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns.Add Position:=(latest_month)

            .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns(latest_month).Range(1, 1).Value = eval_clmn_n

        End If

        If isTableRowName(agent_name, qs_nr_evaluation_scores) Then

            i = getTableRowNumber(agent_name, qs_nr_evaluation_scores)

            .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns(eval_clmn_n).DataBodyRange(i - 1, 1).Value = score

        Else

            For Each qr_cell In .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns("Agent").DataBodyRange

                If StrComp(agent_name, qr_cell.Value) = -1 Then

                    .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListRows.Add Position:=i

                    .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListRows(i).Range(1, 1).Value = agent_name

                    With .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListRows(i).Range(1, 1)

                        .Font.Color = rgbBlack

                    End With

                    .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns(eval_clmn_n).DataBodyRange(i, 1).Value = score

                    found_later = True

                    Exit For

                End If

                i = i + 1

            Next qr_cell

            If Not found_later Then

                .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListRows.Add AlwaysInsert:=True

                .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListRows(i).Range(1, 1).Value = agent_name

                .ListObjects(Replace(qs_nr_evaluation_scores, " ", "")).ListColumns(eval_clmn_n).DataBodyRange(i, 1).Value = score

            End If

        End If

    End With

End Sub

 

Private Sub addQSAverages(ByVal agent_name As String, ByVal score As Double, ByVal esat As Double, table_n As String)

    Dim i As Integer

    Dim latest_month As Integer

    Dim raw_date As String

    Dim prev_month_clmn_n As String

    Dim qr_cell As Range

    Dim found_later As Boolean

    Dim primary_label As String

    Dim secondary_label As String

    'month_esat_score_clmn_n = "Evaluation " & eval_num

    i = 1

    found_later = False

    primary_label = getPrimaryScoreName()

    secondary_label = getSecondaryScoreName()

    With output_book.Worksheets("Metric Percentages")

        If isTableRowName(agent_name, table_n) Then

            i = getTableRowNumber(agent_name, table_n)

            .ListObjects(Replace(table_n, " ", "")).ListColumns(primary_label).DataBodyRange(i - 1, 1).Value = score

            .ListObjects(Replace(table_n, " ", "")).ListColumns(secondary_label).DataBodyRange(i - 1, 1).Value = esat

        Else

            For Each qr_cell In .ListObjects(Replace(table_n, " ", "")).ListColumns("Agent").DataBodyRange

                If StrComp(agent_name, qr_cell.Value) = -1 Then

                    .ListObjects(Replace(table_n, " ", "")).ListRows.Add Position:=i

                    .ListObjects(Replace(table_n, " ", "")).ListRows(i).Range(1, 1).Value = agent_name

                    With .ListObjects(Replace(table_n, " ", "")).ListRows(i).Range(1, 1)

                        .Font.Color = rgbBlack

                    End With

                    .ListObjects(Replace(table_n, " ", "")).ListColumns(primary_label).DataBodyRange(i, 1).Value = score

                    .ListObjects(Replace(table_n, " ", "")).ListColumns(secondary_label).DataBodyRange(i, 1).Value = esat

                    found_later = True

                    Exit For

                End If

                i = i + 1

            Next qr_cell

            If Not found_later Then

                .ListObjects(Replace(table_n, " ", "")).ListRows.Add AlwaysInsert:=True

                .ListObjects(Replace(table_n, " ", "")).ListRows(i).Range(1, 1).Value = agent_name

                .ListObjects(Replace(table_n, " ", "")).ListColumns(primary_label).DataBodyRange(i, 1).Value = score

                .ListObjects(Replace(table_n, " ", "")).ListColumns(secondary_label).DataBodyRange(i, 1).Value = esat

            End If

        End If

    End With

End Sub

 

Public Sub expandTableDown(ByVal table_n As String, ByVal sheet_n As String, ByRef wb As Workbook)

    Dim ws As Worksheet

    Set ws = wb.Worksheets(sheet_n)

    Dim last_clmn As Integer

    Dim last_row As Integer

    Dim first_clmn As Integer

    Dim first_row As Integer

    With wb.Worksheets(sheet_n)

        With .ListObjects(Replace(table_n, " ", ""))

            first_clmn = .ListObjects(Replace(table_n, " ", "")).HeaderRowRange.Column

            first_row = .ListObjects(Replace(table_n, " ", "")).HeaderRowRange.row

           last_clmn = .HeaderRowRange.End(xlToRight).Column

            If IsEmpty(.HeaderRowRange.offset(1, 0).Value) Then

                last_row = first_row + 1

            Else

                last_row = .ListObjects(Replace(table_n, " ", "")).DataBodyRange.End(xlDown).row + 1

            End If

            .Resize Range("$" & getColumnLetterFromNum(first_clmn) & "$" & first_row & ":$" & getColumnLetterFromNum(last_clmn) & "$" & last_row)

        End With

    End With

End Sub

 

 

Private Function isTableColumnName(ByVal clmn_n As Variant, ByVal table_name As String, ByVal sheet_n As String, Optional ByRef wb As Workbook) As Boolean

    Dim header_r As Range

    Dim found_ans As Boolean

    Dim counter As Integer

    counter = 1

    found_ans = False

    If IsEmpty(sheet_n) Or Len(sheet_n) < 1 Then

        sheet_n = "Quality Ranking"

    End If

    If IsMissing(wb) Then

        Set wb = output_book

    End If

    With getWorkSheet(sheet_n)

        For Each header_r In .ListObjects(Replace(table_name, " ", "")).HeaderRowRange

            If IsNumeric(clmn_n) And counter = clmn_n Then

                isTableColumnName = True

                found_ans = True

                Exit For

            ElseIf header_r.Value = clmn_n Then

                isTableColumnName = True

                found_ans = True

                Exit For

            End If

            counter = counter + 1

        Next header_r

    End With

    If Not found_ans Then

        isTableColumnName = False

    End If

End Function

 

Private Function isTableRowName(ByVal row_n As String, ByVal table_name As String, Optional sheet_n As String = "", Optional wb As Workbook) As Boolean

    Dim cell_r As Range

    Dim found As Boolean

    found = False

    If IsEmpty(sheet_n) Or Len(sheet_n) < 1 Then

        sheet_n = "Quality Ranking"

    End If

    If IsMissing(wb) Then

        Set wb = output_book

    End If

    With getWorkSheet(sheet_n)

        For Each cell_r In .ListObjects(Replace(table_name, " ", "")).ListColumns(1).DataBodyRange

            If cell_r.Value = row_n Then

                found = True

                Exit For

            End If

        Next cell_r

    End With

    isTableRowName = found

End Function

 

Private Function getTableClmnNumber(ByVal clmn_n As String, ByVal table_name As String, Optional sheet_n As String = "", Optional wb As Workbook) As Integer

    Dim header_r As Range

    Dim leftmost_rightblock_clmn_letter As String

    Dim i As Integer

    If IsEmpty(sheet_n) Or Len(sheet_n) < 1 Then

        sheet_n = "Quality Ranking"

    End If

    If IsMissing(wb) Then

        Set wb = output_book

    End If

    With getWorkSheet(sheet_n)

        i = 1

        For Each header_r In .ListObjects(Replace(table_name, " ", "")).HeaderRowRange

            If header_r.Value = clmn_n Then

                getTableClmnNumber = i

                Exit For

            End If

            i = i + 1

        Next header_r

    End With

End Function

 

Private Function getTableRowNumber(ByVal agent_name As String, ByVal table_name As String, Optional sheet_n As String = "", Optional primary_key_clmn As String = "SP", Optional wb As Workbook) As Integer

    Dim row_found As Boolean

    Dim temp_int As Integer

    Dim i As Integer

    Dim primarykey_r As Range

    row_found = False

    If IsEmpty(sheet_n) Or Len(sheet_n) < 1 Then

        sheet_n = "Quality Ranking"

    End If

    On Error Resume Next

    If wb = Null Or IsEmpty(wb) Then

        Set wb = output_book

    End If

    Err.Clear

    On Error GoTo 0

    With wb.Worksheets(sheet_n)

        i = 1

        For Each primarykey_r In .ListObjects(Replace(table_name, " ", "")).ListColumns(primary_key_clmn).Range

            If primarykey_r.Value = agent_name Then

                row_found = True

                getTableRowNumber = i

                Exit For

            End If

            i = i + 1

        Next primarykey_r

    End With

End Function

 

Private Sub addQualityRankingProcedural(ByVal agent_name As String, ByVal score As Double, Optional ByVal this_month As String = "", Optional ByVal this_year As String = "")

    Dim i As Integer

    Dim latest_month As Integer

    Dim raw_date As String

    Dim prev_month_clmn_n As String

    Dim qr_cell As Range

    Dim found_later As Boolean

    month_proc_score_clmn_n = getQrCurrentMonthProc(this_month, this_year)

    i = 1

    found_later = False

    With output_book.Worksheets("Quality Ranking")

        If Not isTableColumnName(month_proc_score_clmn_n, qr_proc_named_range_n, "Quality Ranking") Then

            latest_month = getTableClmnNumber(leftmost_rightblockclmn_qr, qr_proc_named_range_n)

            .ListObjects(qr_proc_named_range_n).ListColumns.Add Position:=(latest_month)

            .ListObjects(qr_proc_named_range_n).ListColumns(latest_month).Range(1, 1).Value = month_proc_score_clmn_n

             prev_month_clmn_n = getQrPreviousMonthProc(this_month, this_year)

             If Not isTableColumnName(prev_month_clmn_n, qr_proc_named_range_n, "Quality Ranking", output_book) Then

                .ListObjects(qr_proc_named_range_n).ListColumns.Add Position:=(latest_month)

                .ListObjects(qr_proc_named_range_n).ListColumns(latest_month).Range(1, 1).Value = prev_month_clmn_n

                latest_month = latest_month + 1

             End If

             For Each qr_cell In .ListObjects(qr_proc_named_range_n).ListColumns(getPrimaryScoreName() & " Difference").DataBodyRange

                qr_cell.Formula = "=[@[" & month_proc_score_clmn_n & "]]-[@[" & prev_month_clmn_n & "]]"

                qr_cell.offset(0, 1).Formula = "=([@[" & month_proc_score_clmn_n & "]]-[@[" & prev_month_clmn_n & "]])/[@[" & prev_month_clmn_n & "]]"

            Next qr_cell

            created_new_month = True

        End If

        If isTableRowName(agent_name, qr_proc_named_range_n) Then

            i = getTableRowNumber(agent_name, qr_proc_named_range_n)

            .ListObjects(qr_proc_named_range_n).ListColumns(month_proc_score_clmn_n).DataBodyRange(i - 1, 1).Value = score

        Else

            For Each qr_cell In .ListObjects(qr_proc_named_range_n).ListColumns(qr_primarykey_clmn_n).DataBodyRange

                If .ListObjects(qr_proc_named_range_n).ListRows(1).Range(1, 1).Value = "||Table Start DO NOT REMOVE||" Then

                    .ListObjects(qr_proc_named_range_n).ListRows(1).Range(1, 1).Value = agent_name

                    .ListObjects(qr_proc_named_range_n).ListColumns(month_proc_score_clmn_n).DataBodyRange(1, 1).Value = score

                    found_later = True

                    Exit For

                ElseIf StrComp(agent_name, qr_cell.Value) = -1 Then

                    .ListObjects(qr_proc_named_range_n).ListRows.Add Position:=i

                    .ListObjects(qr_proc_named_range_n).ListRows(i).Range(1, 1).Value = agent_name

                    With .ListObjects(qr_proc_named_range_n).ListRows(i).Range(1, 1)

                        .Font.Color = rgbBlack

                    End With

                    .ListObjects(qr_proc_named_range_n).ListColumns(month_proc_score_clmn_n).DataBodyRange(i, 1).Value = score

                    found_later = True

                    Exit For

                End If

                i = i + 1

            Next qr_cell

            If Not found_later Then

                .ListObjects(qr_proc_named_range_n).ListRows.Add AlwaysInsert:=True

                .ListObjects(qr_proc_named_range_n).ListRows(i).Range(1, 1).Value = agent_name

                .ListObjects(qr_proc_named_range_n).ListColumns(month_proc_score_clmn_n).DataBodyRange(i, 1).Value = score

            End If

        End If

    End With

End Sub

 

Private Sub addQualityRankingEsat(ByVal agent_name As String, ByVal score As Double, Optional ByVal this_month As String = "", Optional ByVal this_year As String = "")

    Dim i As Integer

    Dim latest_month As Integer

    Dim raw_date As String

    Dim prev_month_clmn_n As String

    Dim qr_cell As Range

    Dim found_later As Boolean

    month_esat_score_clmn_n = getQrCurrentMonthEsat(this_month, this_year)

    i = 1

    found_later = False

    With output_book.Worksheets("Quality Ranking")

        If Not isTableColumnName(month_esat_score_clmn_n, qr_esat_named_range_n, "Quality Ranking") Then

            latest_month = getTableClmnNumber(leftmost_rightblockclmn_qr, qr_esat_named_range_n)

            .ListObjects(qr_esat_named_range_n).ListColumns.Add Position:=(latest_month)

            .ListObjects(qr_esat_named_range_n).ListColumns(latest_month).Range(1, 1).Value = month_esat_score_clmn_n

             prev_month_clmn_n = getQrPreviousMonthEsat(this_month, this_year)

             If Not isTableColumnName(prev_month_clmn_n, qr_esat_named_range_n, "Quality Ranking") Then

                .ListObjects(qr_esat_named_range_n).ListColumns.Add Position:=(latest_month)

                .ListObjects(qr_esat_named_range_n).ListColumns(latest_month).Range(1, 1).Value = prev_month_clmn_n

                latest_month = latest_month + 1

             End If

             For Each qr_cell In .ListObjects(qr_esat_named_range_n).ListColumns(getSecondaryScoreName() & " Difference").DataBodyRange

                qr_cell.Formula = "=[@[" & month_esat_score_clmn_n & "]]-[@[" & prev_month_clmn_n & "]]"

                qr_cell.offset(0, 1).Formula = "=([@[" & month_esat_score_clmn_n & "]]-[@[" & prev_month_clmn_n & "]])/[@[" & prev_month_clmn_n & "]]"

            Next qr_cell

            created_new_month = True

        End If

        If isTableRowName(agent_name, qr_esat_named_range_n) Then

            i = getTableRowNumber(agent_name, qr_esat_named_range_n)

            .ListObjects(qr_esat_named_range_n).ListColumns(month_esat_score_clmn_n).DataBodyRange(i - 1, 1).Value = score

        Else

            For Each qr_cell In .ListObjects(qr_esat_named_range_n).ListColumns(qr_primarykey_clmn_n).DataBodyRange

                If .ListObjects(qr_esat_named_range_n).ListRows(1).Range(1, 1).Value = "||Table Start DO NOT REMOVE||" Then

                    .ListObjects(qr_esat_named_range_n).ListRows(1).Range(1, 1).Value = agent_name

                    .ListObjects(qr_esat_named_range_n).ListColumns(month_esat_score_clmn_n).DataBodyRange(1, 1).Value = score

                    found_later = True

                    Exit For

                ElseIf StrComp(agent_name, qr_cell.Value) = -1 Then

                    .ListObjects(qr_esat_named_range_n).ListRows.Add Position:=i

                    .ListObjects(qr_esat_named_range_n).ListRows(i).Range(1, 1).Value = agent_name

                    With .ListObjects(qr_esat_named_range_n).ListRows(i).Range(1, 1)

                        .Font.Color = rgbBlack

                    End With

                    .ListObjects(qr_esat_named_range_n).ListColumns(month_esat_score_clmn_n).DataBodyRange(i, 1).Value = score

                    found_later = True

                    Exit For

                End If

                i = i + 1

            Next qr_cell

           If Not found_later Then

                .ListObjects(qr_esat_named_range_n).ListRows.Add AlwaysInsert:=True

                .ListObjects(qr_esat_named_range_n).ListRows(i).Range(1, 1).Value = agent_name

                .ListObjects(qr_esat_named_range_n).ListColumns(month_esat_score_clmn_n).DataBodyRange(i, 1).Value = score

            End If

        End If

    End With

End Sub

 

Private Function getQrCurrentMonthProc(Optional ByVal this_month As String = "", Optional ByVal this_year As String = "") As String

    Dim primary_eval_type As String

    primary_eval_type = getPrimaryScoreName()

    If Len(this_month) < 1 Then

        this_month = MonthName(Month(Now()))

    End If

    If Len(this_year) < 1 Then

        this_year = Year(Now())

    End If

    If this_month = "September" Then

        getQrCurrentMonthProc = Left(this_month, 4) & ". " & Right(this_year, 2) & " " & primary_eval_type

    ElseIf Len(this_month) > 4 Then

        getQrCurrentMonthProc = Left(this_month, 3) & ". " & Right(this_year, 2) & " " & primary_eval_type

    Else

        getQrCurrentMonthProc = this_month & ". " & Right(this_year, 2) & " " & primary_eval_type

    End If

End Function

 

Private Function getQrPreviousMonthProc(Optional ByVal this_month As String = "", Optional ByVal this_year As String = "") As String

    Dim raw_date As String

    If Len(this_month) < 1 Then

        raw_date = Format(DateAdd("M", -1, Date), "Long Date")

        this_month = LTrim(RTrim(Mid(raw_date, InStr(1, raw_date, ",") + 1, InStr(InStr(1, raw_date, ",") + 2, raw_date, " ") - (InStr(1, raw_date, ",") + 1))))

    End If

    If Len(this_year) < 1 Then

        raw_date = Format(DateAdd("M", -1, Date), "Long Date")

        this_year = LTrim(Mid(raw_date, InStr(InStr(1, raw_date, ",") + 1, raw_date, ",") + 1))

    End If

   

    Select Case this_month

        Case "January"

            this_year = CStr(CInt(this_year) - 1)

            this_month = "December"

        Case "February"

            this_month = "January"

        Case "March"

            this_month = "February"

        Case "April"

            this_month = "March"

        Case "May"

            this_month = "April"

        Case "June"

            this_month = "May"

        Case "July"

            this_month = "June"

        Case "August"

            this_month = "July"

        Case "September"

            this_month = "August"

        Case "October"

            this_month = "September"

        Case "November"

            this_month = "October"

        Case "December"

            this_month = "November"

        Case Else

    End Select

   

    If this_month = "September" Then

        getQrPreviousMonthProc = Left(this_month, 4) & ". " & Right(this_year, 2) & " Procedural"

    ElseIf Len(this_month) > 4 Then

        getQrPreviousMonthProc = Left(this_month, 3) & ". " & Right(this_year, 2) & " Procedural"

    Else

        getQrPreviousMonthProc = this_month & ". " & Right(this_year, 2) & " Procedural"

    End If

End Function

 

Private Function getQrCurrentMonthEsat(Optional ByVal this_month As String = "", Optional ByVal this_year As String = "") As String

    Dim sub_title As String

    sub_title = getSecondaryScoreName()

    If Len(this_month) < 1 Then

        this_month = MonthName(Month(Now()))

    End If

    If Len(this_year) < 1 Then

        this_year = Year(Now())

    End If

    If this_month = "September" Then

        getQrCurrentMonthEsat = Left(this_month, 4) & ". " & Right(this_year, 2) & " " & sub_title

    ElseIf Len(this_month) > 4 Then

        getQrCurrentMonthEsat = Left(this_month, 3) & ". " & Right(this_year, 2) & " " & sub_title

    Else

        getQrCurrentMonthEsat = this_month & ". " & Right(this_year, 2) & " " & sub_title

    End If

End Function

 

Private Function getQrPreviousMonthEsat(Optional ByVal this_month As String = "", Optional ByVal this_year As String = "") As String

    Dim sub_title As String

    Dim raw_date As String

    sub_title = getSecondaryScoreName()

    If Len(this_month) < 1 Then

        raw_date = Format(DateAdd("M", -1, Date), "Long Date")

        this_month = LTrim(RTrim(Mid(raw_date, InStr(1, raw_date, ",") + 1, InStr(InStr(1, raw_date, ",") + 2, raw_date, " ") - (InStr(1, raw_date, ",") + 1))))

    End If

    If Len(this_year) < 1 Then

        raw_date = Format(DateAdd("M", -1, Date), "Long Date")

        this_year = LTrim(Mid(raw_date, InStr(InStr(1, raw_date, ",") + 1, raw_date, ",") + 1))

    End If

    Select Case this_month

        Case "January"

            this_year = CStr(CInt(this_year) - 1)

            this_month = "December"

        Case "February"

            this_month = "January"

        Case "March"

            this_month = "February"

        Case "April"

            this_month = "March"

        Case "May"

            this_month = "April"

        Case "June"

            this_month = "May"

        Case "July"

            this_month = "June"

        Case "August"

            this_month = "July"

        Case "September"

            this_month = "August"

        Case "October"

            this_month = "September"

        Case "November"

            this_month = "October"

        Case "December"

            this_month = "November"

        Case Else

    End Select

    If this_month = "September" Then

        getQrPreviousMonthEsat = Left(this_month, 4) & ". " & Right(this_year, 2) & " " & sub_title

    ElseIf Len(this_month) > 4 Then

        getQrPreviousMonthEsat = Left(this_month, 3) & ". " & Right(this_year, 2) & " " & sub_title

    Else

        getQrPreviousMonthEsat = this_month & ". " & Right(this_year, 2) & " " & sub_title

    End If

End Function

 

Sub filterCommentTab(ByVal sheet_n As String)

    Dim first_filter_r As Range

    Dim second_filter_r As Range

    Dim third_filter_r As Range

    Dim custom_sort(0 To 16) As String

    Dim temp_sheet As Worksheet

   

    Application.AddCustomList ListArray:=metric_types

   

    With output_book.Worksheets(sheet_n)

        Set first_filter_r = .Range(getColumnLetter("Agent") & first_header_row, .Range(getColumnLetter("Agent") & first_header_row).End(xlDown)) ' normally "Agent"

        Set second_filter_r = .Range(getColumnLetter("Time Stamp") & first_header_row, .Range(getColumnLetter("Agent") & first_header_row).End(xlDown))

        Set third_filter_r = .Range(getColumnLetter("Metric Type") & first_header_row, .Range(getColumnLetter("Agent") & first_header_row).End(xlDown))

        .Select

        .Range(.Range(getColumnLetter(first_clmn_label) & "1:" & getColumnLetter(last_clmn_label) & "1"), .Range(getColumnLetter(first_clmn_label) & "1:" & getColumnLetter(last_clmn_label) & "1").End(xlDown)).Select

        'With .Range(.Range(getColumnLetter(first_clmn_label) & "1:" & getColumnLetter(last_clmn_label) & "1"), .Range(getColumnLetter(first_clmn_label) & "1:" & getColumnLetter(last_clmn_label) & "1").End(xlDown))

        With .Sort

            .SortFields.Clear

            '.SortFields.Add key:=third_filter_r, SortOn:=xlSortOnValues, _

             '   Order:=xlAscending, _

            '    OrderCustom:=Application.CustomListCount + 1, Orientation:=xlTopToBottom

            .SortFields.Add key:=first_filter_r, SortOn:=xlSortOnValues, _

                ORDER:=xlAscending, DataOption:=xlSortNormal

            .SortFields.Add key:=second_filter_r, SortOn:=xlSortOnValues, _

                ORDER:=xlAscending, DataOption:=xlSortNormal

            .Header = xlYes

            .MatchCase = True

            .Apply

        End With

    End With

    Application.DeleteCustomList Application.CustomListCount

End Sub

 

Sub filterPivotField(field As PivotField, Value)

    With field

        If .Orientation = xlPageField Then

            .CurrentPage = Value

        ElseIf .Orientation = xlRowField Or .Orientation = xlColumnField Then

            Dim i As Long

            On Error Resume Next ' Needed to avoid getting errors when manipulating PivotItems that were deleted from the data source.

            ' Set first item to Visible to avoid getting no visible items while working

            For i = 1 To .PivotItems.Count

                If .PivotItems(i).name = Value Then

                    .PivotItems(i).Visible = True

                    Exit For

                End If

                If .PivotItems(i).name = "FILLER-IGNORE ME" Then

                    .PivotItems(i).Visible = False

                End If

            Next i

            Err.Clear

            On Error GoTo 0

        End If

    End With

End Sub

 

Sub clearPivotFilter(field As PivotField, Optional field_name As String)

    With field

        If .Orientation = xlPageField Then

            ' Nothing

        ElseIf .Orientation = xlRowField Or .Orientation = xlColumnField Then

            Dim i As Long

            On Error Resume Next ' Needed to avoid getting errors when manipulating PivotItems that were deleted from the data source,

                    ' and merely the requirement that there be at least one visible.

            For i = 1 To field.PivotItems.Count

                If Not IsMissing(field_name) And TypeName(field_name) = "String" And Len(field_name) > 0 Then

                    If .PivotItems(i).name = field_name Then

                        .PivotItems(i).Visible = False

                        Exit For

                    End If

                Else

                    .PivotItems(i).Visible = False

                End If

            Next i

            Err.Clear

            On Error GoTo 0

        End If

    End With

End Sub

 

Sub fillEvalAvgTable(ByRef ws As Worksheet, ByVal row As Long, ByVal clmn As Long, ByVal table_n As String, ByVal this_etype As String, oEvaluations As EvaluationCollection)

    Dim rngTmp As Range

    Dim qty_total As Double

    Dim qty_othertotal As Double

    Dim sheets_i As Integer

    Dim type_ct As Integer

    Dim temp_int As Double

    Dim tmp_num As Integer

    Dim temp_text As String

    Dim eval_type As String

    Dim temp_evalarray() As EvalProcEsatTypeDate

    Dim temp_esat_evalarray() As EvalProcEsatTypeDate

    this_etype = LCase(LTrim(RTrim(this_etype)))

    With ws

        With .cells(row, clmn)

            .Value = table_n

            .offset(1, 0).Value = "Agent"

            .offset(1, 1).Value = getPrimaryScoreName()

            .offset(1, 2).Value = getSecondaryScoreName()

            .offset(1, 3).Value = "Count"

 

            tmp_num = 1

            For sheets_i = 1 To output_book.Worksheets.Count

                With output_book.Worksheets(sheets_i)

                    temp_text = .name

                End With

                If isCollectionKey(temp_text, oAgentsEvalScores) Then

                    type_ct = 0

                    qty_total = 0

                    qty_othertotal = 0

                    temp_evalarray = oAgentsEvalScores(temp_text)

                    For temp_int = LBound(temp_evalarray) To UBound(temp_evalarray)

                        With temp_evalarray(temp_int)

                            eval_type = .etype

                            If LCase(LTrim(RTrim(eval_type))) = this_etype Then

                                If this_etype = "certification" Then

                                  certification_eval = True

                                End If

                                qty_total = qty_total + .procedural

                                type_ct = type_ct + 1

                                If oEvaluations.isSecondaryAvgAvailableForAgent(temp_text) Then

                                  qty_othertotal = qty_othertotal + .esat

                                End If

                            End If

                        End With

                    Next temp_int

                    If qty_total > 0 Then

                        .offset(1 + tmp_num, 0).Value = temp_text

                        .offset(1 + tmp_num, 1).Value = qty_total / type_ct

                        If oEvaluations.isSecondaryAvgAvailableForAgent(temp_text) Then

                            .offset(1 + tmp_num, 2).Value = qty_othertotal / type_ct

                        End If

                        .offset(1 + tmp_num, 3).Value = type_ct

                        tmp_num = tmp_num + 1

                    End If

               End If

            Next sheets_i

        End With

        Set rngTmp = .cells(row, clmn)

        If truncate_numbers Then

            rngTmp.Range(rngTmp, rngTmp.End(xlDown)).NumberFormat = "0.00"

        End If

        .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range(rngTmp.offset(1, 0), .cells(rngTmp.End(xlDown).row, rngTmp.Column + 3)), _

                            XlListObjectHasheaders:=xlYes).name = Replace(table_n, " ", "")

        On Error GoTo NoBod

        If .ListObjects.Item(Replace(table_n, " ", "")).DataBodyRange.Count > 0 Then

            .ListObjects.Item(Replace(table_n, " ", "")).DataBodyRange.HorizontalAlignment = xlLeft

            ' conditional formatting

            ' Primary score range

'            If .ListObjects.Item(Replace(table_n, " ", "")).ListRows.Count = 1 Then

'                Set rngTmp = rngTmp.offset(2, 1)

'            Else

'                Set rngTmp = .Range(rngTmp.offset(2, 1), rngTmp.offset(2, 1).End(xlDown))

'            End If

            Call processPrimaryScoreQuintileFormatting(.Range(rngTmp.offset(2, 1), rngTmp.offset(tmp_num, 1)), rngTmp.offset(2, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlA1))

            ' Primary score formatting

 

            ' Secondary range

'            If .ListObjects.Item(Replace(table_n, " ", "")).ListRows.Count = 1 Then

'                Set rngTmp = .cells(row, clmn).offset(2, 2)

'            Else

'                Set rngTmp = .Range(.cells(row, clmn).offset(2, 2), .cells(row, clmn).offset(2, 2).End(xlDown))

'            End If

            Call processSecondaryScoreQuintileFormatting(.Range(rngTmp.offset(2, 2), rngTmp.offset(tmp_num, 2)), rngTmp.offset(2, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlA1))

            ' Secondary score conditional formatting

        End If

NoBod:

        Err.Clear

        On Error GoTo 0

    End With

End Sub

 

Function getEvalTypeCount(this_etype As String)

    Dim sheets_i As Integer

    Dim type_ct As Integer

    Dim temp_int As Integer

    Dim temp_text As String

    Dim temp_evalarray() As EvalProcEsatTypeDate

    Dim eval_type As String

    this_etype = LCase(LTrim(RTrim(this_etype)))

    type_ct = 0

    For sheets_i = 1 To output_book.Worksheets.Count

        With output_book.Worksheets(sheets_i)

            temp_text = .name

        End With

       If isCollectionKey(temp_text, oAgentsEvalScores) Then

           temp_evalarray = oAgentsEvalScores(temp_text)

            For temp_int = LBound(temp_evalarray) To UBound(temp_evalarray)

                With temp_evalarray(temp_int)

                    eval_type = .etype

                    If LCase(LTrim(RTrim(eval_type))) = this_etype Then

                        type_ct = type_ct + 1

                    End If

                End With

            Next temp_int

       End If

    Next sheets_i

    getEvalTypeCount = type_ct

End Function

 

 

Sub setHeaderFormatting(r As Range, ws As Worksheet)

    With ws

        With r.Interior

            .PatternColorIndex = xlAutomatic

            .ThemeColor = xlThemeColorAccent1

            .TintAndShade = 0

            .PatternTintAndShade = 0

        End With

        With r.Font

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

        End With

        With r.Font

            .name = "Tahoma"

            .Size = 10

            .Strikethrough = False

            .Superscript = False

            .Subscript = False

            .OutlineFont = False

            .Shadow = False

            .Underline = xlUnderlineStyleNone

            .ThemeColor = xlThemeColorDark1

            .TintAndShade = 0

            .ThemeFont = xlThemeFontNone

        End With

        r.Font.Bold = True

    End With

End Sub

 

Public Function isKeyOfCollection(coll As Collection, key As String) As Boolean

    On Error GoTo NoKey

    coll.Item key

    isKeyOfCollection = True

NoKey:

    Err.Clear

    On Error GoTo 0

End Function

 

Public Function getFirstEvalRow(sheet_name As String, current_row As Long, Optional eval_date As Date) As Long

    Dim this_offset As Integer

    this_offset = -1

    With output_book.Worksheets(sheet_name)

        With .Range(getColumnLetter("Time Stamp") & current_row)

            If eval_date = Empty Then

                eval_date = .Value

            End If

            Do Until Not .offset(this_offset, 0).Value = eval_date

                this_offset = this_offset - 1

            Loop

        End With

   End With

    getFirstEvalRow = current_row + this_offset

End Function

 

 

Public Function stripLeadTrailNewline(ByVal text As String)

    On Error GoTo TextLenZero

    If Len(text) > 0 Then

        While Left(text, 1) = Chr(10) Or Left(text, 1) = Chr(160)

            If Left(text, 1) = Chr(10) Then

                text = Replace(text, Chr(10), "", Count:=1)

            Else

                text = Replace(text, Chr(160), "", Count:=1)

            End If

        Wend

        While Right(text, 1) = Chr(10) Or Left(text, 1) = Chr(160)

            If Right(text, 1) = Chr(10) Then

                text = Replace(text, Chr(10), "", Count:=1, Start:=-1)

            Else

                text = Replace(text, Chr(160), "", Count:=1, Start:=-1)

            End If

        Wend

    End If

TextLenZero:

    Err.Clear

    On Error GoTo 0

    stripLeadTrailNewline = text

End Function

 

Public Sub addEsatDeficitEntry(sDeficitAgent As String, ByVal dDeficitEvalDate As Date)

  Dim r As Range

  Dim asdf As Worksheet

  If Not SheetExists("Bad Format Rows", output_book) Then

    Set asdf = getWorkSheet("Bad Format Rows")

    Call initializeCommentTab(asdf)

    garbage_row_offset = first_comment_row

  End If

  With output_book.Worksheets("Bad Format Rows")

    With .Range(getColumnLetter(last_clmn_label) & first_header_row).offset(0, 2)

      .Value = "Missing " & getSecondaryScoreName

      If IsEmpty(.offset(1, 0).Value) Then

        Set r = .offset(1, 0)

      Else

        Set r = .End(xlDown).offset(1, 0)

      End If

      r.Value = sDeficitAgent

      r.offset(0, 1).Value = dDeficitEvalDate

    End With

  End With

End Sub

 

'Public Sub ArraySort(ByRef cllctn As Variant)

'    ReDim temp_arr(LBound(cllctn) To UBound(cllctn)) As Variant

'    ReDim lesser_arr(0 To 0) As Variant

'    ReDim greater_arr(0 To 0) As Variant

'    Dim whole_lotta_nothing As Variant

'    Dim i As Integer

'    Dim incumbent As Variant

'    If Len(cllctn) > 1 Then

'        If cllctn(LBound(cllctn)) < cllctn(LBound(cllctn) + 1) Then

'            lesser_arr(LBound(lesser_arr)) = cllctn(LBound(cllctn))

'            greater_arr(LBound(greater_arr)) = cllctn(LBound(cllctn) + 1)

'        Else

'            lesser_arr(LBound(lesser_arr)) = cllctn(LBound(cllctn) + 1)

'            greater_arr(LBound(greater_arr)) = cllctn(LBound(cllctn))

'        End If

'        If Len(cllctn) > 2 Then

'            For i = LBound(cllctn) + 2 To UBound(cllctn)

'            On Error GoTo ArraySortInvalidIndexError

'                whole_lotta_nothing = cllctn(i)

'                If lesser_arr(UBound(lesser_arr)) < cllctn(i) And cllctn(i) < greater_arr(LBound(greater_arr)) Then

'                    ReDim Preserve lesser_arr(UBound(lesser_arr) + 1)

'                    lesser_arr(UBound(lesser_arr)) = cllctn(i)

'                ElseIf Not lesser_arr(UBound(lesser_arr)) < cllctn(i) And cllctn(i) < greater_arr(LBound(greater_arr)) Then

'                    ArrayAddLessThan lesser_arr, cllctn(i)

'                ElseIf lesser_arr(UBound(lesser_arr)) < cllctn(i) And Not cllctn(i) < greater_arr(LBound(greater_arr)) Then

'                    ArrayAddMoreThan greater_arr, cllctn(i)

'                Else

'                    ReDim Preserve lesser_arr(UBound(lesser_arr) + 1)

'                    lesser_arr(UBound(lesser_arr)) = cllctn(i)

'                End If

'ArraySortInvalidIndexError:

'            Next i

'        End If

'        For i = LBound(greater_arr) To UBound(greater_arr)

'            ReDim Preserve lesser_arr(UBound(lesser_arr) + 1)

'            lesser_arr(UBound(lesser_arr)) = greater_arr(i)

'        Next i

'        cllctn = lesser_arr

'    End If

'End Sub

 

Public Sub ArrayAddLessThan(ByRef arr As Variant, el As Variant)

 

    Dim i As Integer

    Dim k As Integer

    i = (UBound(arr) + 1)

    ReDim temp_array(LBound(arr) To i)

    i = UBound(arr)

    k = UBound(temp_array)

    While i > LBound(arr) - 1

        If el < arr(i) Then

            temp_array(k) = arr(i)

        Else

            temp_array(k) = el

            k = k - 1

            While i > LBound(arr) - 1

                temp_array(k) = arr(i)

                k = k - 1

                i = i - 1

            Wend

        End If

        k = k - 1

        i = i - 1

    Wend

    ReDim arr(LBound(arr) To UBound(arr) + 1)

    arr = temp_array

End Sub

 

Public Sub ArrayAddMoreThan(ByRef arr As Variant, el As Variant)

    ReDim temp_array(LBound(arr) To UBound(arr) + 1)

    Dim i As Integer

    Dim k As Integer

    i = LBound(arr)

    k = LBound(temp_array)

    While i < UBound(arr) + 1

        If el > arr(i) Then

            temp_array(k) = arr(i)

        Else

            temp_array(k) = el

            k = k + 1

            While i < UBound(arr) + 1

                temp_array(k) = arr(i)

                k = k + 1

                i = i + 1

            Wend

        End If

        k = k + 1

        i = i + 1

    Wend

    ReDim arr(LBound(arr) To UBound(arr) + 1)

    arr = temp_array

End Sub

 

Public Function copyArray(ByRef incumbent As Variant) As Variant

    Dim answer() As Variant

    'ReDim answer(LBound(incumbent) To UBound(incumbent))

    If Len(incumbent) > 0 Then

        Dim i As Integer

        For i = LBound(incumbent) To UBound(incumbent)

            If IsNumeric(incumbent(i)) Or TypeName(incumbent(i)) = "Boolean" Or IsDate(incumbent(i)) Or TypeName(incumbent(i)) = "String" Then

                answer(i) = incumbent(i)

            ElseIf IsArray(incumbent(i)) Then

                answer(i) = copyArray(incumbent(i))

            Else

                answer(i) = incumbent(i).clone()

            End If

        Next i

    End If

    copyArray = answer

End Function

'    Dim answer() As Variant

'    ReDim answer(LBound(incumbent) To UBound(incumbent))

'    If Len(incumbent) > 0 Then

'        Dim i As Integer

'        i = LBound(incumbent)

'        Do

'            If GetType(incumbent(i)).IsPrimitive Or GetType(incumbent(i)).name = "String" Then

'                answer(i) = incumbent(i)

'            ElseIf GetType(incumbent(i)).IsArray Then

'                answer(i) = StaticUtilities.copyArray(incumbent(i))

'            Else

'                answer(i) = StaticUtilities.Utilities.copyObject(incumbent(i))

'            End If

'

'    End If

'    copyArray = answer

'End Sub

 

Function SheetExists(ByVal shtName As String, Optional ByRef wb As Workbook) As Boolean

    Dim sht As Worksheet

    If wb Is Nothing Then

        Set wb = ActiveWorkbook

    End If

    On Error Resume Next

    Set sht = wb.Worksheets(shtName)

    Err.Clear

    On Error GoTo 0

    SheetExists = Not sht Is Nothing

End Function

 Function isCollectionKey(ByVal key As String, ByRef coll As Collection) As Boolean

    On Error GoTo NoKey

    coll.Item key

    isCollectionKey = True

NoKey:

    Err.Clear

    On Error GoTo 0

End Function

 

' Looking for Evaluation Type from first meta tag

Public Function SuppliedEvalType(supplied_eval_type As String)

    Select Case LCase(LTrim(RTrim(supplied_eval_type)))

        Case "verbal"

            SuppliedEvalType = "Verbal"

        Case "business"

            SuppliedEvalType = "Business"

        Case "negative"

            SuppliedEvalType = "Negative"

        Case "survey"

            SuppliedEvalType = "Survey"

        Case "written"

            SuppliedEvalType = "Written"

        Case "certification"

            SuppliedEvalType = "Certification"

        Case Else

            SuppliedEvalType = ""

    End Select

End Function

 

'Public Function getMetricTypeRevised(ByVal metric_name As Variant, ByVal meta_data As String) As String

'    Select Case LCase(LTrim(RTrim(metric_name)))

'        Case "comment"

'            If LCase(LTrim(RTrim(meta_data))) = "evaluator satisfaction" Or LCase(LTrim(RTrim(meta_data))) = "esat" Then

'                getMetricTypeRevised = "ESAT"

'            ElseIf LCase(LTrim(RTrim(meta_data))) = "hold comment" Then

'                getMetricTypeRevised = "Hold Comment"

'            ElseIf LCase(LTrim(RTrim(meta_data))) = "verification" Then

'                getMetricTypeRevised = "Verification"

'            ElseIf LCase(LTrim(RTrim(meta_data))) = "business comment" Then

'                getMetricTypeRevised = "Business Comment"

'            ElseIf Len(SuppliedEvalType(meta_data)) > 0 Then

'                getMetricTypeRevised = "Comment"

'            Else

'                getMetricTypeRevised = ""

'            End If

'        Case "did the agent provide the complete and correct answer?"

'            getMetricTypeRevised = "Accuracy / Completeness"

'        Case "was the complete and correct answer provided?"

'            getMetricTypeRevised = "Accuracy / Completeness"

'        Case "were the next steps communicated completely?"

'            getMetricTypeRevised = "Complete Expectations"

'        Case "were next steps communicated completely?"

'            getMetricTypeRevised = "Complete Expectations"

'        Case "was the clients concern resolved in a timely manner?"

'            getMetricTypeRevised = "Timely Resolution"

'        Case "was the case resolved in a timely manner?"

'            getMetricTypeRevised = "Timely Resolution"

'        Case "was world class service demonstrated on this interaction?"

'            getMetricTypeRevised = "World-Class Service"

'        Case "was the case handled professionally?"

'            getMetricTypeRevised = "World-Class Service"

'        Case "was response free from grammatical errors?"

'            getMetricTypeRevised = "Grammar Error Free"

'        Case "did agent show forward thinking for any additional questions that may arise?"

'            getMetricTypeRevised = "Forward Thinking"

'        Case "did the agent create a satisfactory hold experience?"

'            getMetricTypeRevised = "Hold Experience"

'        Case "did the agent create a satisfactory transfer experience?"

'            getMetricTypeRevised = "Transfer Experience"

'        Case "was the appropriate greeting used?"

'            getMetricTypeRevised = "Appropriate Greeting"

'        Case "was this interaction free of authentication errors?"

'            getMetricTypeRevised = "Verification"

'        Case "were correct resources used?"

'            getMetricTypeRevised = "Correct Resources"

'        Case "was the appropriate closing used?"

'            getMetricTypeRevised = "Appropriate Closing"

'        Case "were all business guidelines and areas of focus addressed?"

'            getMetricTypeRevised = "Business Processes"

'        Case "were all business requirements accurately completed?"

'            getMetricTypeRevised = "Business Processes"

'        Case "actively listened to client and correctly identified the root cause of the call"

'            getMetricTypeRevised = "Actively Listened"

'        Case "appropriately controlled the call"

'            getMetricTypeRevised = "Controlled Call"

'        Case "communicated in a clear and confident manner"

'            getMetricTypeRevised = "Clear / Confident"

'        Case "followed correct processes and procedures"

'            getMetricTypeRevised = "Process / Procedures"

'        Case "logged call correctly and added necessary notes"

'            getMetricTypeRevised = "Call Log"

'        Case "provided a warm opening and a fond farewell"

'            getMetricTypeRevised = "Opening / Farewell"

'        Case "provided accurate information"

'            getMetricTypeRevised = "Accurate Information"

'        Case "set appropriate expectation with client"

'            getMetricTypeRevised = "Expectations"

'        Case "what is the likelihood that the caller will need to call again due to the agent's handling of the interaction?"

'            getMetricTypeRevised = "Callback"

'        Case "added or updated all required information"

'            getMetricTypeRevised = "Added / Updated"

'        Case "offered survey at the end of the call"

'            getMetricTypeRevised = "Survey"

'        Case "followed appropriate hold / dial procedure"

'            getMetricTypeRevised = "Hold / Transfer"

'        Case "was call log entered correctly?"

'            getMetricTypeRevised = "Call Log Entered"

'        Case "were all details provided in call log?"

'            getMetricTypeRevised = "Call Log Details"

'        Case "was a survey offered?"

'            getMetricTypeRevised = "Survey"

'        Case "did the agent promote branch self-service?"

'            getMetricTypeRevised = "Promote Branch Self-Service"

'        Case "if challenged, had the processing been done correctly?"

'            If LCase(LTrim(RTrim(meta_data))) = "yes" Or LCase(LTrim(RTrim(meta_data))) = "no" Or LCase(LTrim(RTrim(meta_data))) = "n/a" Then

'                business_expectation_na_resolved = True

'            End If

'            getMetricTypeRevised = "Challenge Processed Correctly"

'        Case "did the agents response meet sl?"

'            getMetricTypeRevised = "Service Level Met"

'        Case Else

'            getMetricTypeRevised = ""

'    End Select

'

'End Function

 

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

            If cbxCRD.Value Then

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

            If cbxMSSN.Value Then

                getMetricSection = "Business Expectations"

            Else

                getMetricSection = "Verification"

            End If

    End Select

End Function

 

' Will set metric_max and metric_pct in Evaluation Collection with this format: k:metric_type_revised, v:array [0:comment, 1:metric_score, 2:metric_max, 3:metric_pct]

Public Sub setScoreMax(ByRef eval_notes_and_scores As Collection)

    Dim i As Integer

    Dim call_handling_i As Integer

    Dim call_handling As Collection

    ' k:metric_type_revised, v:array [0:comment, 1:metric_score, 2:metric_max, 3:metric_pct]

    Dim metric_score_i As Integer

    Dim metric_max_i As Integer

    Dim metric_pct_i As Integer

    Dim eval_comment_i As Integer

    metric_score_i = 1

    metric_max_i = 2

    metric_pct_i = 3

    eval_comment_i = 0

    Dim metric_types(0 To 16) As String

    metric_types(0) = "Comment"

    metric_types(1) = "Evaluator Satisfaction"

    metric_types(2) = "Verification"

    metric_types(3) = "Accurate Information"

    metric_types(4) = "Process / Procedures"

    metric_types(5) = "Expectations"

    metric_types(6) = "Hold / Transfer"

    metric_types(7) = "Call Log"

    metric_types(8) = "Added / Updated"

    metric_types(9) = "Survey"

    metric_types(10) = "Callback"

    metric_types(11) = "Opening / Farewell"

    metric_types(12) = "Actively Listened"

    metric_types(13) = "Controlled Call"

    metric_types(14) = "Clear / Confident"

    metric_types(15) = "Hold Comment"

    metric_types(16) = "Business Comment"

   

    Dim enum_to_test As EvaluationComment

    For i = LBound(metric_types) To UBound(metric_types)

        enum_to_test = metric_types(i)

        On Error GoTo IndexOutOfBounds

        Select Case enum_to_test

            Case "Comment"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = "--"

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = "--"

            Case "Evaluator Satisfaction"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 5

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Verification"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = "--"

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = "--"

            Case "Accurate Information"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.5

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Expectations"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.5

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Process / Procedures"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 2

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Opening / Farewell"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.25

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Actively Listened"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.25

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Controlled Call"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.25

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Clear / Confident"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = 0.25

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = eval_notes_and_scores.Item(enum_to_test)(metric_score_i) / eval_notes_and_scores.Item(enum_to_test)(metric_max_i)

            Case "Business Comment"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = "--"

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = "--"

            Case "Hold Comment"

                eval_notes_and_scores.Item(enum_to_test)(metric_max_i) = "--"

                eval_notes_and_scores.Item(enum_to_test)(metric_pct_i) = "--"

            Case "Hold / Transfer"

                call_handling.Add Item:="Hold / Transfer"

            Case "Call Log"

                call_handling.Add Item:="Call Log"

            Case "Added / Updated"

                call_handling.Add Item:="Added / Updated"

            Case "Survey"

                call_handling.Add Item:="Survey"

            Case "Callback"

                call_handling.Add Item:="Callback"

IndexOutOfBounds:

        End Select

    Next i

    Err.Clear

    On Error GoTo 0

    For call_handling_i = 1 To call_handling.Count

        eval_notes_and_scores(call_handling(call_handling_i))(metric_max_i) = Math.Round(1 / call_handling.Count, 2)

        eval_notes_and_scores(call_handling(call_handling_i))(metric_pct_i) = eval_notes_and_scores(call_handling(call_handling_i))(metric_score_i) / eval_notes_and_scores(call_handling(call_handling_i))(metric_max_i)

    Next call_handling_i

End Sub

' If found, will return the index in an array that element resides at. Else returns negative one (-1).

Public Function getIndex(el As Variant, arr As Variant) As Integer

    Dim found As Boolean

    Dim i As Integer

    On Error GoTo getindexindexoutofbounds

    found = False

    For i = UBound(arr) To LBound(arr) Step -1

        If el = arr(i) Then

            getIndex = i

            found = True

        End If

    Next i

    If Not found Then

getindexindexoutofbounds:

        getIndex = -1

    End If

    Err.Clear

    On Error GoTo 0

End Function

 

Public Function getLastContentRow(sheet_n As String)

    With output_book.Worksheets(sheet_n)

        If IsEmpty(.Range(getColumnLetter(first_clmn_label) & first_comment_row).offset(1, 0).Value) Then

            getLastContentRow = first_comment_row

        Else

            getLastContentRow = .Range(getColumnLetter(first_clmn_label) & first_comment_row).End(xlDown).row

        End If

    End With

End Function

 

'Public Function getThisMaxScore(ByVal metric_type As String, Optional ByVal meta_data As Variant) As Double

'    If cbxCRD.Value And getMetricSection(metric_type) = "Business Expectations" Then

'        Select Case metric_type

'            Case "Appropriate Greeting", "Appropriate Closing", "Call Log Entered", "Survey", "Correct Resources", "Call Log Details"

'                getThisMaxScore = 0.5

'            Case "Business Processes"

'                getThisMaxScore = 2

'        End Select

'    ElseIf (cbxNNA.Value Or cbxMSSN.Value) And metric_type = "Business Processes" Then

'        getThisMaxScore = 2

'    ElseIf (cbxNNA.Value Or cbxMSSN.Value) And metric_type = "Appropriate Closing" Then

'        getThisMaxScore = 0.5

'    'ElseIf 'getCurrentEvalType = "Written" And

'    ElseIf (metric_type = "Grammar Error Free" Or metric_type = "Forward Thinking" Or metric_type = "Appropriate Closing" Or metric_type = "Challenge Processed Correctly" _

'    Or metric_type = "Business Processes") Then

'        Select Case metric_type

'            Case "Forward Thinking", "Grammar Error Free"

'                getThisMaxScore = 0.25

'            Case "Appropriate Closing"

'                getThisMaxScore = 0.5

'            Case "Challenge Processed Correctly"

'                getThisMaxScore = 1

'            Case "Business Processes"

'                If cbxCRD.Value Then

'                    getThisMaxScore = 1

'                Else

'                    getThisMaxScore = 2

'                End If

'            Case "Service Level Met"

'                getThisMaxScore = 1

'        End Select

'    Else

'        Select Case metric_type

'            Case "Comment"

'                getThisMaxScore = 0

'            Case "Evaluator Satisfaction"

'                getThisMaxScore = 5

'            Case "ESAT"

'                getThisMaxScore = 5

'            Case "Verification"

'                getThisMaxScore = 1

'            Case "Accurate Information", "Expectations", "Appropriate Greeting"

'                getThisMaxScore = 0.5

'            Case "Process / Procedures"

'                getThisMaxScore = 2

'            Case "Opening / Farewell"

'                getThisMaxScore = 0.25

'            Case "Actively Listened"

'                getThisMaxScore = 0.25

'            Case "Controlled Call"

'                getThisMaxScore = 0.25

'            Case "Clear / Confident"

'                getThisMaxScore = 0.25

'            Case "Business Comment"

'                getThisMaxScore = 0

'            Case "Hold Comment"

'                getThisMaxScore = 0

'            Case "Hold / Transfer"

'                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

'                    hold_transfer_offset = output_row_offset

'                    call_handling_count = call_handling_count + 1

'                End If

'                getThisMaxScore = 0

'            Case "Call Log"

'                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

'                    call_log_offset = output_row_offset

'                    call_handling_count = call_handling_count + 1

'                End If

'                getThisMaxScore = 0

'            Case "Added / Updated"

'                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

'                    added_updated_offset = output_row_offset

'                    call_handling_count = call_handling_count + 1

'                End If

'                getThisMaxScore = 0

'            Case "Survey"

'                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

'                    survey_offset = output_row_offset

'                    call_handling_count = call_handling_count + 1

'                End If

'                getThisMaxScore = 0

'            Case "Callback"

'                If Not LCase(LTrim(RTrim(CStr(meta_data)))) = "n/a" Then

'                    callback_offset = output_row_offset

'                    call_handling_count = call_handling_count + 1

'                End If

'                getThisMaxScore = 0

'            Case "Accuracy / Completeness"

'                getThisMaxScore = 2

'            Case "Complete Expectations", "Timely Resolution"

'                getThisMaxScore = 0.75

'            Case "World-Class Service", "Correct Resources", "Appropriate Closing"

'                getThisMaxScore = 1

'            Case "Hold Experience", "Transfer Experience"

'                getThisMaxScore = 0.25

'            Case "Business Processes"

'                getThisMaxScore = 2.5

'            Case "Promote Branch Self-Service"

'                getThisMaxScore = 2

'        End Select

'    End If

'End Function

 

Public Sub addCallHandlingMax()

    Dim temp_agent_offset As Integer

    Dim temp_sm_offset As Integer

    Dim m_score As Double

    Dim m_max As Double

    If Not call_handling_count = 0 Then

        m_max = (1 / call_handling_count)

        If hold_transfer_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - hold_transfer_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - hold_transfer_offset)

            Call addMetricMax(hold_transfer_offset, m_max, primary_output_tab_n)

            m_score = getMetricScore(hold_transfer_offset, primary_output_tab_n)

            If getMetricScore(hold_transfer_offset, primary_output_tab_n) = 1 Then

                Call addMetricScore(hold_transfer_offset, m_score * m_max, primary_output_tab_n)

                Call addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call addMetricPercent(hold_transfer_offset, m_score, primary_output_tab_n)

            Call addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call addMetricMax(temp_sm_offset, m_max, current_sm)

                Call addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If call_log_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - call_log_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - call_log_offset)

            Call addMetricMax(call_log_offset, m_max, primary_output_tab_n)

            m_score = getMetricScore(call_log_offset, primary_output_tab_n)

            If getMetricScore(call_log_offset, primary_output_tab_n) = 1 Then

                Call addMetricScore(call_log_offset, m_score * m_max, primary_output_tab_n)

                Call addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call addMetricPercent(call_log_offset, m_score, primary_output_tab_n)

            Call addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call addMetricMax(temp_sm_offset, m_max, current_sm)

                Call addMetricPercent(temp_sm_offset, m_score / m_max, current_sm)

            End If

        End If

        If added_updated_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - added_updated_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - added_updated_offset)

            Call addMetricMax(added_updated_offset, m_max, primary_output_tab_n)

            m_score = getMetricScore(added_updated_offset, primary_output_tab_n)

            If getMetricScore(added_updated_offset, primary_output_tab_n) = 1 Then

                Call addMetricScore(added_updated_offset, m_score * m_max, primary_output_tab_n)

                Call addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call addMetricPercent(added_updated_offset, m_score, primary_output_tab_n)

            Call addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call addMetricMax(temp_sm_offset, m_max, current_sm)

                Call addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If survey_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - survey_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - survey_offset)

            Call addMetricMax(survey_offset, m_max, primary_output_tab_n)

            m_score = getMetricScore(survey_offset, primary_output_tab_n)

            If getMetricScore(survey_offset, primary_output_tab_n) = 1 Then

                Call addMetricScore(survey_offset, m_score * m_max, primary_output_tab_n)

                Call addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

            End If

            Call addMetricMax(temp_agent_offset, m_max, current_agent)

            Call addMetricPercent(survey_offset, m_score, primary_output_tab_n)

            Call addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call addMetricMax(temp_sm_offset, m_max, current_sm)

                Call addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

        If callback_offset > 0 Then

            temp_agent_offset = current_agent_offset - (output_row_offset - callback_offset)

            temp_sm_offset = current_sm_offset - (output_row_offset - callback_offset)

            Call addMetricMax(callback_offset, m_max, primary_output_tab_n)

            m_score = getMetricScore(callback_offset, primary_output_tab_n)

            If getMetricScore(callback_offset, primary_output_tab_n) = 1 Then

                Call addMetricScore(callback_offset, m_score * m_max, primary_output_tab_n)

                Call addMetricScore(temp_agent_offset, m_score * m_max, current_agent)

                If sm_known Then

                    Call addMetricScore(temp_sm_offset, m_score * m_max, current_sm)

                End If

            End If

            Call addMetricMax(temp_agent_offset, m_max, current_agent)

           

            Call addMetricPercent(callback_offset, getMetricScore(callback_offset, primary_output_tab_n) / getMetricMax(callback_offset, primary_output_tab_n), primary_output_tab_n)

            Call addMetricPercent(temp_agent_offset, m_score, current_agent)

            If sm_known Then

                Call addMetricMax(temp_sm_offset, m_max, current_sm)

                Call addMetricPercent(temp_sm_offset, m_score, current_sm)

            End If

        End If

    End If

    call_handling_count = 0

    callback_offset = 0

    survey_offset = 0

    added_updated_offset = 0

    hold_transfer_offset = 0

    call_log_offset = 0

End Sub

 

Public Sub formatFillerRowOmnibus()

    With output_book.Worksheets(current_agent)

        With .Range(getColumnLetter(first_clmn_label) & current_agent_offset & ":" & getColumnLetter(last_clmn_label) & current_agent_offset)

            With .Interior

                .Pattern = xlSolid

                .PatternThemeColor = xlThemeColorAccent1

                .ThemeColor = xlThemeColorLight1

                .TintAndShade = 0

            End With

        End With

    End With

    With output_book.Worksheets(primary_output_tab_n)

        With .Range(getColumnLetter(first_clmn_label) & output_row_offset & ":" & getColumnLetter(last_clmn_label) & output_row_offset)

            .Interior.Color = rgbBlack

        End With

    End With

    If sm_known Then

        With output_book.Worksheets(current_sm)

            With .Range(getColumnLetter(first_clmn_label) & current_sm_offset & ":" & getColumnLetter(last_clmn_label) & current_sm_offset)

                .Interior.Color = rgbBlack

            End With

        End With

    End If

End Sub

 

 

Public Sub handleDateProcScore(ByVal d As Date, ByVal score As Variant)

'    Dim cell_name As String

'    Dim a_sheet As Worksheet

'    Dim r As Range

'    Dim temp_evalarray() As EvalProcEsatTypeDate

'    Dim temp_a() As Double

'    Dim temp_cell_value As Variant

'    Dim temp_adate() As Date

'    Dim temp_bdate() As Date

'    Dim i As Integer

'    Dim temp_eval As New EvalProcEsatTypeDate

'    Dim temp_coll As Collection

'    Dim temp_coll2 As Collection

'   'If isCollectionKey(current_agent, sp_evaldate_collection) Then

'        temp_adate = sp_evaldate_collection(current_agent)

'   '    sp_evaldate_collection.Remove (current_agent)

'        If duplicate_evaluation Then

'   '        sp_evaldate_collection.Add key:=current_agent, Item:=temp_adate

'        Else

'            ReDim Preserve temp_adate(LBound(temp_adate) To (UBound(temp_adate) + 1))

'        End If

'   'Else

'        ReDim temp_adate(0 To 0)

'    End If

'    If Not duplicate_evaluation Then

'        this_agent_eval_qty = this_agent_eval_qty + 1

'        If this_agent_eval_qty > eval_qty_max Then

'            eval_qty_max = this_agent_eval_qty

'        End If

'

'        temp_adate(UBound(temp_adate)) = d

'        sp_evaldate_collection.Add key:=current_agent, Item:=temp_adate

'        If (Not Not proc_scores) <> 0 Then

'            ' Has content, increment

'            ReDim Preserve proc_scores(LBound(proc_scores) To (UBound(proc_scores) + 1))

'        Else

'            ReDim proc_scores(0 To 0)

'        End If

'        proc_scores(UBound(proc_scores)) = score

'        Set temp_eval = New EvalProcEsatTypeDate

'        temp_eval.procedural = score

'        temp_eval.edate = d

'        ReDim temp_evalarray(0 To 0)

'        If isCollectionKey(current_agent, oAgentsEvalScores) Then

'            temp_evalarray = oAgentsEvalScores.Item(current_agent)

'            oAgentsEvalScores.Remove (current_agent)

'            ReDim Preserve temp_evalarray(UBound(temp_evalarray) + 1)

'        End If

'        Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'        oAgentsEvalScores.Add Item:=temp_evalarray, key:=current_agent

'        If sm_known Then

'            If Not isCollectionKey(current_sm, sm_proc_scores) Then

'                ReDim temp_a(0 To 0)

'            Else

'                temp_a = sm_proc_scores(current_sm)

'                sm_proc_scores.Remove (current_sm)

'                ReDim Preserve temp_a(LBound(temp_a) To UBound(temp_a) + 1)

'            End If

'            temp_a(UBound(temp_a)) = score

'            sm_proc_scores.Add Item:=temp_a, key:=current_sm

'        End If

'

'        If Not first_metric Then

'            If first_eval Then

'                first_eval = False

'            End If

'            If call_handling_count > 0 Then

'                Call addCallHandlingMax

'            End If

'            Call incrementOffsetOmnibus

'        End If

'        first_metric = True

'        Call addAgentNameOmnibus(current_agent)

'        Call addTimeStampOmnibus(d)

'        Call addProceduralScoreOmnibus(score)

'        verification_comment_supplied = False

'    Else

'        With output_book.Worksheets(current_agent)

'            If .Range(getColumnLetter("Agent") & current_agent_offset).Value = current_agent And Len(.Range(getColumnLetter("Metric Type") & current_agent_offset).Value) = 0 Then

'                .Range(getColumnLetter("Agent") & current_agent_offset).Value = ""

'            End If

'        End With

'        With output_book.Worksheets(primary_output_tab_n)

'            If .Range(getColumnLetter("Agent") & output_row_offset).Value = current_agent And Len(.Range(getColumnLetter("Metric Type") & output_row_offset).Value) = 0 Then

'                .Range(getColumnLetter("Agent") & output_row_offset).Value = ""

'            End If

'        End With

'        If sm_known Then

'            With output_book.Worksheets(current_sm)

'                If .Range(getColumnLetter("Agent") & current_sm_offset).Value = current_agent And Len(.Range(getColumnLetter("Metric Type") & current_sm_offset).Value) = 0 Then

'                    .Range(getColumnLetter("Agent") & current_sm_offset).Value = ""

'                End If

'            End With

'        End If

'    End If

End Sub

 

Public Function getOutputCellLocation(offset As Integer, label As String) ' Letter-Number string cell identifier

    getOutputCellLocation = getColumnLetter(label) & offset

End Function

 

Public Function getColumnLetter(ByVal label As String)

    Select Case LCase(LTrim(RTrim(label)))

        Case "agent"

            getColumnLetter = "A"

        Case "metric"

            getColumnLetter = "B"

        Case "metric type"

            getColumnLetter = "B"

        Case "comment"

            getColumnLetter = "C"

        Case "evaluation type"

            getColumnLetter = "D"

        Case "time stamp"

            getColumnLetter = "E"

        Case "metric score label"

            getColumnLetter = "F"

        Case "metric score"

            getColumnLetter = "G"

        Case "maximum metric score"

            getColumnLetter = "H"

        Case "metric percentage"

            getColumnLetter = "I"

        Case "client satisfaction"

            getColumnLetter = "J"

        Case "procedural score"

            getColumnLetter = "J"

        Case "business expectations"

            getColumnLetter = "K"

        Case "evaluator satisfaction"

            getColumnLetter = "K"

        Case Else

            getColumnLetter = ""

    End Select

End Function

 

Public Function getColumnLetterFromNum(ByVal num As Long) As String

    Dim vArr

    vArr = Split(cells(1, num).Address(True, False), "$")

    getColumnLetterFromNum = vArr(0)

End Function

 

Public Function getSmName(Optional ByVal agent_name As String) As String

    Dim working_answer As String

  Dim sAgentName As String

  If Not agent_name = "" Then

    sAgentName = agent_name

  Else

    sAgentName = current_agent

  End If

    If Not isCollectionKey(sAgentName, sp_sm_collection) Then

        working_answer = getSm.getManagerName(sAgentName)

        If working_answer = "Braun, Jonathan" Or working_answer = "Kershaw, Ashlee" Or working_answer = "Hansen, Cory" Or working_answer = "McLaughlin, Scott" Or working_answer = "Farley, Matthew" Then

            working_answer = ""

        End If

        If Len(working_answer) > 0 Then

            sp_sm_collection.Add key:=sAgentName, Item:=working_answer

        ElseIf SheetExists(sm_sp_sheet_name) Then

            Dim temp_r As Range

            With output_book.Worksheets(sm_sp_sheet_name)

                For Each temp_r In .Range("SM_SP_SP_clmn", .Range("SM_SP_SP_clmn").End(xlDown))

                    If TypeName(temp_r.Value) = "String" And temp_r.Value = sAgentName Then

                        sp_sm_collection.Add key:=temp_r.Value, Item:=temp_r.offset(0, -1).Value

                        Exit For

                    End If

                Next temp_r

            End With

        End If

    End If

    If isCollectionKey(sAgentName, sp_sm_collection) Then

        getSmName = sp_sm_collection(sAgentName)

    Else

        getSmName = "--"

    End If

End Function

 

Public Sub addAgentNameOmnibus(ByVal name As String)

    Call addAgentName(output_row_offset, name, primary_output_tab_n)

    Call addAgentName(current_agent_offset, name, current_agent)

    If sm_known Then

        Call addAgentName(current_sm_offset, name, current_sm)

    End If

End Sub

'temp_cell_value = getAgentName(output_row_offset, primary_output_tab_n)

'Call addAgentName(output_row_offset + 1, temp_cell_value, primary_output_tab_n)

'Call addAgentName(current_agent_offset + 1, temp_cell_value, current_agent)

'If sm_known Then

'    Call addAgentName(current_sm_offset + 1, temp_cell_value, current_sm)

'End If

'temp_cell_value = "Verification"

'Call addMetricType(output_row_offset + 1, temp_cell_value, primary_output_tab_n)

'Call addMetricType(current_agent_offset + 1, temp_cell_value, current_agent)

'If sm_known Then

'    Call addMetricType(current_sm_offset + 1, temp_cell_value, current_sm)

'End If

'temp_cell_value = getEvalType(output_row_offset, primary_output_tab_n)

'Call addMetricType(output_row_offset + 1, temp_cell_value, primary_output_tab_n)

'Call addMetricType(current_agent_offset + 1, temp_cell_value, current_agent)

'If sm_known Then

'    Call addMetricType(current_sm_offset + 1, temp_cell_value, current_sm)

'End If

'temp_cell_value = getTimeStamp(output_row_offset, primary_output_tab_n)

'Call addTimeStamp(output_row_offset + 1, temp_cell_value, primary_output_tab_n)

'Call addTimeStamp(current_agent_offset + 1, temp_cell_value, current_agent)

'If sm_known Then

'    Call addTimeStamp(current_sm_offset + 1, temp_cell_value, current_sm)

'End If

'temp_cell_value = "Yes"

'Call addMetricScoreLabel(output_row_offset + 1, temp_cell_value, primary_output_tab_n)

'Call addMetricScoreLabel(current_agent_offset + 1, temp_cell_value, current_agent)

'If sm_known Then

'    Call addMetricScoreLabel(current_sm_offset + 1, temp_cell_value, current_sm)

'End If

 

 

Public Sub addMetricTypeOmnibus(ByVal metric_type As String)

    Call addMetricType(output_row_offset, metric_type, primary_output_tab_n)

    Call addMetricType(current_agent_offset, metric_type, current_agent)

    If sm_known Then

        Call addMetricType(current_sm_offset, metric_type, current_sm)

    End If

End Sub

 

Public Sub addCommentOmnibus(ByVal comment_string As String)

    Call addComment(output_row_offset, comment_string, primary_output_tab_n)

    Call addComment(current_agent_offset, comment_string, current_agent)

    If sm_known Then

        Call addComment(current_sm_offset, comment_string, current_sm)

    End If

End Sub

 

Public Sub addTimeStampOmnibus(ByVal time_stamp As Date)

    Call addTimeStamp(output_row_offset, time_stamp, primary_output_tab_n)

    Call addTimeStamp(current_agent_offset, time_stamp, current_agent)

    If sm_known Then

        Call addTimeStamp(current_sm_offset, time_stamp, current_sm)

    End If

End Sub

 

Public Sub addEvalTypeOmnibus(ByVal eval_type As String)

    Call addEvalType(output_row_offset, eval_type, primary_output_tab_n)

    Call addEvalType(current_agent_offset, eval_type, current_agent)

    If sm_known Then

        Call addEvalType(current_sm_offset, eval_type, current_sm)

    End If

End Sub

 

Public Sub addMetricScoreLabelOmnibus(ByVal label As String)

    Call addMetricScoreLabel(output_row_offset, label, primary_output_tab_n)

    Call addMetricScoreLabel(current_agent_offset, label, current_agent)

    If sm_known Then

        Call addMetricScoreLabel(current_sm_offset, label, current_sm)

    End If

End Sub

 

Public Sub addMetricScoreOmnibus(ByVal score As Double)

    Call addMetricScore(output_row_offset, score, primary_output_tab_n)

    Call addMetricScore(current_agent_offset, score, current_agent)

    If sm_known Then

        Call addMetricScore(current_sm_offset, score, current_sm)

    End If

End Sub

 

Public Sub addMetricMaxOmnibus(ByVal max As Double)

    Call addMetricMax(output_row_offset, max, primary_output_tab_n)

    Call addMetricMax(current_agent_offset, max, current_agent)

    If sm_known Then

        Call addMetricMax(current_sm_offset, max, current_sm)

    End If

End Sub

 

Public Sub addMetricPercentageOmnibus(ByVal percent As Double)

    Call addMetricPercent(output_row_offset, percent, primary_output_tab_n)

    Call addMetricPercent(current_agent_offset, percent, current_agent)

    If sm_known Then

        Call addMetricPercent(current_sm_offset, percent, current_sm)

    End If

End Sub

 

Public Sub addProceduralScoreOmnibus(ByVal proc As Double)

    Call addProcScore(output_row_offset, proc, primary_output_tab_n)

    Call addProcScore(current_agent_offset, proc, current_agent)

    If sm_known Then

        Call addProcScore(current_sm_offset, proc, current_sm)

    End If

End Sub

 

Public Sub addEsatScoreOmnibus(ByVal esat As Double)

    Call addEsatScore(output_row_offset, esat, primary_output_tab_n)

    Call addEsatScore(current_agent_offset, esat, current_agent)

    If sm_known Then

        Call addEsatScore(current_sm_offset, esat, current_sm)

    End If

End Sub

 

Public Sub addMetricMax(ByVal offset As Integer, ByVal metric_max As Double, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    If Not SheetExists(sheet_n, output_book) Then

        Set ws = output_book.Sheets.Add

        ws.name = sheet_n

        Call initializeCommentTab(output_book.Worksheets(sheet_n))

    Else

        Set ws = output_book.Sheets(sheet_n)

    End If

    cell = getOutputCellLocation(offset, "maximum metric score")

    ws.Range(cell).Value = metric_max

End Sub

 

Public Sub addEsatScore(ByVal offset As Integer, ByVal score, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    If Not SheetExists(sheet_n, output_book) Then

        Set ws = output_book.Sheets.Add

        ws.name = sheet_n

        Call initializeCommentTab(output_book.Worksheets(sheet_n))

    Else

        Set ws = output_book.Sheets(sheet_n)

    End If

    cell = getOutputCellLocation(offset, "Evaluator Satisfaction")

    ws.Range(cell).Value = score

End Sub

 

Public Sub addComment(ByVal offset As Integer, ByVal comment As String, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    If Not SheetExists(sheet_n, output_book) Then

        Set ws = output_book.Sheets.Add

        ws.name = sheet_n

        Call initializeCommentTab(output_book.Worksheets(sheet_n))

    Else

        Set ws = output_book.Sheets(sheet_n)

    End If

    cell = getOutputCellLocation(offset, "Comment")

    ws.Range(cell).Value = comment

End Sub

 

Public Sub addAgentName(ByVal offset As Integer, ByVal name As String, Optional ByVal sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    cell = getOutputCellLocation(offset, "Agent")

    ws.Range(cell).Value = name

End Sub

 

Public Sub addMetricType(ByVal offset As Integer, ByVal m_type As String, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    cell = getOutputCellLocation(offset, "Metric Type")

    ws.Range(cell).Value = m_type

End Sub

 

Public Sub addTimeStamp(ByVal offset As Integer, ByVal time_stamp As Date, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Time Stamp")

        .Range(cell).Value = time_stamp

    End With

End Sub

 

Public Sub addMetricScore(ByVal offset As Integer, ByVal m_score, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Score")

        .Range(cell).Value = m_score

    End With

End Sub

 

Public Sub addMetricPercent(ByVal offset As Integer, ByVal metric_percent, Optional sheet_n As String = "All Agents")

    Dim cell As String

    Dim ws As Worksheet

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Percentage")

        .Range(cell).Value = metric_percent

    End With

End Sub

 

 

Public Sub addProcScore(ByVal offset As Integer, ByVal proc_score As Double, Optional sheet_n As String = "All Agents")

    Dim cell As String

    Dim ws As Worksheet

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Procedural Score")

        .Range(cell).Value = proc_score

    End With

End Sub

 

Public Sub addEvalType(ByVal offset As Integer, ByVal eval_type As String, Optional sheet_n As String = "All Agents")

    Dim cell As String

    Dim ws As Worksheet

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Evaluation Type")

        .Range(cell).Value = eval_type

    End With

End Sub

 

Public Sub addMetricScoreLabel(ByVal offset As Integer, ByVal score_label As String, Optional sheet_n As String = "All Agents")

    Dim cell As String

    Dim ws As Worksheet

    Dim bSheetNeedsInitialization As Boolean

    If offset < first_comment_row Then

        offset = first_comment_row

    End If

    If Not SheetExists(sheet_n, output_book) And Len(sheet_n) > 0 Then

      bSheetNeedsInitialization = True

    End If

    Set ws = getWorkSheet(sheet_n)

    If bSheetNeedsInitialization Then

      Call initializeCommentTab(output_book.Worksheets(sheet_n))

    End If

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Score Label")

        .Range(cell).Value = score_label

    End With

End Sub

 

 

Public Sub fillEvalTypeUp(ByVal offset As Integer, ByVal eval_type As String, Optional sheet_n As String = "All Agents")

    Dim cell As String

    Dim i As Long

    With output_book.Sheets(sheet_n)

        cell = getOutputCellLocation(offset, "Evaluation Type")

        With .Range(cell, .Range(cell).End(xlUp))

            For i = .cells.Count To 2 Step -1

                .Item(i).Value = eval_type

            Next i

        End With

    End With

End Sub

 

Public Sub fillTimeStampUp(ByVal offset As Integer, ByVal t_stamp As Date, Optional ByVal sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim ctrl_shift_up As Range

    Dim i As Long

    Set ws = output_book.Sheets(sheet_n)

    cell = getOutputCellLocation(offset, "Time Stamp")

    Set ctrl_shift_up = ws.Range(cell, ws.Range(cell).End(xlUp))

    For i = ctrl_shift_up.cells.Count To 2 Step -1

        ctrl_shift_up(i).Value = t_stamp

    Next i

End Sub

 

Public Sub fillProcScoreUp(ByVal offset As Integer, ByVal proc_score As Double, Optional ByVal sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim ctrl_shift_up As Range

    Dim i As Long

    Set ws = output_book.Sheets(sheet_n)

    cell = getOutputCellLocation(offset, "Procedural Score")

    Set ctrl_shift_up = ws.Range(cell, ws.Range(cell).End(xlUp))

    For i = ctrl_shift_up.cells.Count To 2 Step -1

        ctrl_shift_up(i).Value = proc_score

    Next i

End Sub

 

Public Sub fillEsatScoreUp(ByVal offset As Integer, ByVal eval_type As String, Optional sheet_n As String = "All Agents")

    Dim ws As Worksheet

    Dim cell As String

    Dim ctrl_shift_up As Range

    Dim i As Long

    Set ws = output_book.Sheets(sheet_n)

    cell = getOutputCellLocation(offset, "Evaluator Satisfaction")

    Set ctrl_shift_up = ws.Range(cell, ws.Range(cell).End(xlUp))

    For i = ctrl_shift_up.cells.Count To 2 Step -1

        ctrl_shift_up(i).Value = eval_type

    Next i

End Sub

 

Public Function getMetricScoreLabelMax(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "metric score label")

        getMetricScoreLabelMax = .Range(cell).Value

    End With

End Function

 

Public Function getMetricMax(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "maximum metric score")

        getMetricMax = .Range(cell).Value

    End With

End Function

 

Public Function getEsatScore(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Evaluator Satisfaction")

        getEsatScore = .Range(cell).Value

    End With

End Function

 

Public Function getComment(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As String

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Comment")

        getComment = .Range(cell).Value

    End With

End Function

 

Public Function getAgentName(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As String

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Agent")

        getAgentName = .Range(cell).Value

    End With

End Function

 

Public Function getMetricType(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As String

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Type")

        getMetricType = .Range(cell).Value

    End With

End Function

 

Public Function getTimeStamp(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Date

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Time Stamp")

        getTimeStamp = .Range(cell).Value

    End With

End Function

 

Public Function getMetricScore(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Score")

        getMetricScore = .Range(cell).Value

    End With

End Function

 

Public Function getMetricPercent(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Metric Percentage")

        getMetricPercent = .Range(cell).Value

    End With

End Function

 

 

Public Function getProcScore(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As Double

    Dim cell As String

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Procedural Score")

        getProcScore = .Range(cell).Value

    End With

End Function

 

Public Function getEvalType(ByVal offset As Integer, Optional sheet_n As String = "All Agents") As String

    Dim cell As String

    If IsMissing(sheet_n) Or Len(sheet_n) = 0 Then

        sheet_n = primary_output_tab_n

    End If

    If Not offset > first_comment_row And Not offset = first_comment_row Then

        offset = first_comment_row

    End If

    With getWorkSheet(sheet_n)

        cell = getOutputCellLocation(offset, "Evaluation Type")

        getEvalType = .Range(cell).Value

    End With

End Function

 

Public Function getRevisedMetricScore(ByVal incumbent As Variant, ByVal max As Double, oCurrentEval As Evaluation, Optional ByVal mtype_revised As String, Optional bIsClientSatisfaction As Boolean, Optional sEvalType As String) As Double

  If IsNumeric(incumbent) Then

    getRevisedMetricScore = incumbent

  ElseIf TypeName(incumbent) = "String" Then

    If sEvalType = "" Then

        sEvalType = oCurrentEval.getCurrentEvalType

    End If

    incumbent = LCase(LTrim(RTrim(incumbent)))

    If max = 0 Then

      If incumbent = "yes" Or incumbent = "not likely" Then

        getRevisedMetricScore = 1

      Else

         getRevisedMetricScore = 0

      End If

    ElseIf mtype_revised = "Complete Expectations" Or mtype_revised = "Timely Resolution" Then

      If incumbent = "yes" Then

        getRevisedMetricScore = 0.75

      ElseIf incumbent = "partial" Then

        getRevisedMetricScore = 0.25

      Else

        getRevisedMetricScore = 0

      End If

    ElseIf oCurrentEval.isClientSatisfaction And incumbent = "n/a" Then

      getRevisedMetricScore = max

    ElseIf mtype_revised = "Business Processes" And (oCurrentEval.isNna Or sEvalType = "Written") Then

      If incumbent = "yes" Then

        getRevisedMetricScore = 2

      ElseIf incumbent = "partial" Then

        getRevisedMetricScore = 1

      Else

        getRevisedMetricScore = 0

      End If

    ElseIf mtype_revised = "Business Processes" And oCurrentEval.isCrd Then

      If incumbent = "yes" Then

        getRevisedMetricScore = 2

      ElseIf incumbent = "partial" Then

        getRevisedMetricScore = 0.5

      Else

        getRevisedMetricScore = 0

      End If

    ElseIf mtype_revised = "Business Processes" Then

      If incumbent = "yes" Then

        getRevisedMetricScore = 2.5

      ElseIf incumbent = "partial" Then

        getRevisedMetricScore = 1

      Else

        getRevisedMetricScore = 0

      End If

    Else

      Select Case incumbent

        Case "yes"

          getRevisedMetricScore = max

        Case "partial"

          getRevisedMetricScore = max / 2

        Case "no"

          getRevisedMetricScore = 0

        Case Else

          getRevisedMetricScore = 0

      End Select

    End If

  End If

End Function

 

Public Function getRevisedMetricLabel(ByVal score As Variant, ByVal max As Double, ByVal metric_type As String) As String

    Dim score_label As String

    If containsNonNumericCharacters(score) Then

        score_label = StrConv(LTrim(RTrim(score)), vbProperCase)

        score_label = Replace(score_label, "  ", " ")

        If score_label = "Likley" Then

            score_label = "Likely"

        End If

        If score_label = "Not Likley" Then

            score_label = "not likely"

        End If

        If score_label = "Parital" Then

            score_label = "Partial"

        End If

        If score_label = "Partail" Then

            score_label = "Partial"

        End If

        If score_label = "No Likely" Then

            score_label = "Not Likely"

        End If

        If score_label = "N/a" Then

            score_label = "N/A"

        End If

        getRevisedMetricLabel = score_label

    ElseIf IsNumeric(score) Then

        If (score = 2 And metric_type = "Accurate Information") Or (score = 0.5 And (metric_type = "Process / Procedures" Or metric_type = "Expectations")) Then

            getRevisedMetricLabel = "Yes"

        ElseIf Not max = 0 Then

            Select Case score / max

                Case 1

                    If max = 5 Then

                        getRevisedMetricLabel = "Strongly Agree"

                    Else

                        getRevisedMetricLabel = "Yes"

                    End If

                Case 0.75

                    getRevisedMetricLabel = "Agree"

                Case 0.5

                    If max = 5 Then

                        getRevisedMetricLabel = "Neutral"

                    Else

                        getRevisedMetricLabel = "Partial"

                    End If

                Case 0.25

                    getRevisedMetricLabel = "Disagree"

                Case 0

                    If max = 5 Then

                        getRevisedMetricLabel = "Strongly Disagree"

                    Else

                        getRevisedMetricLabel = "No"

                    End If

                Case Else

                    getRevisedMetricLabel = "Partial"

            End Select

        ElseIf LCase(LTrim(RTrim(metric_type))) = "callback" Then

            If score = 1 Then

                getRevisedMetricLabel = "Not Likely"

            ElseIf score = 0 Then

                getRevisedMetricLabel = "Likely"

            End If

        ElseIf metric_type = "Complete Expectations" Or metric_type = "Timely Resolution" Then

            If score = 0.75 Then

                getRevisedMetricLabel = "Yes"

            ElseIf score = 0.25 Then

                getRevisedMetricLabel = "Partial"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Accuracy / Completeness" Then

            If score = 2 Then

                getRevisedMetricLabel = "Yes"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "World-Class Service" Then

            If score = 1 Then

                getRevisedMetricLabel = "Yes"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Hold Experience" Or metric_type = "Transfer Experience" Then

            If score = 0.25 Then

                getRevisedMetricLabel = "Yes"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Appropriate Greeting" Then

            If score = 0.5 Then

                getRevisedMetricLabel = "Yes"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Correct Resources" Or metric_type = "Appropriate Closing" Then

            If score = 1 Then

                getRevisedMetricLabel = "Yes"

            ElseIf score = 0.5 Then

                getRevisedMetricLabel = "Partial"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Business Processes" Then

            If score = 2.5 Then

                getRevisedMetricLabel = "Yes"

            ElseIf score = 1 Then

                getRevisedMetricLabel = "Partial"

            Else

                getRevisedMetricLabel = "No"

            End If

        ElseIf metric_type = "Survey" Then

            If is_client_experience Then

                If cbxCRD.Value Then

                    If score = 0.5 Then

                        getRevisedMetricLabel = "Yes"

                    Else

                        getRevisedMetricLabel = "No"

                    End If

                End If

            Else

                If score > 0 Then

                    getRevisedMetricLabel = "Yes"

                Else

                    getRevisedMetricLabel = "No"

                End If

            End If

        Else

            If score = 1 Then

                getRevisedMetricLabel = "Yes"

            Else

                getRevisedMetricLabel = "No"

            End If

        End If

    End If

End Function

 

Public Function getRevisedEsatScore(ByVal raw_esat As Variant)

    If IsNumeric(raw_esat) Then

        Select Case CDbl(raw_esat)

            Case 5

                getRevisedEsatScore = CDbl(raw_esat)

            Case 3.75

                getRevisedEsatScore = CDbl(raw_esat)

            Case 2.5

                getRevisedEsatScore = CDbl(raw_esat)

            Case 1.25

                getRevisedEsatScore = CDbl(raw_esat)

            Case 0

                getRevisedEsatScore = CDbl(raw_esat)

            Case Else

                getRevisedEsatScore = 999999999

        End Select

    ElseIf TypeName(raw_esat) = "String" Then

        Select Case LCase(LTrim(RTrim(raw_esat)))

            Case "strongly agree"

                getRevisedEsatScore = 5

            Case "agree"

                getRevisedEsatScore = 3.75

            Case "neutral"

                getRevisedEsatScore = 2.5

            Case "disagree"

                getRevisedEsatScore = 1.25

            Case "strongly disagree"

                getRevisedEsatScore = 0

            Case Else

                getRevisedEsatScore = 999999999

        End Select

    End If

End Function

 

Public Sub addSmTab(ByVal name As String)

    Dim first_sm As Boolean

    Dim temp_sheet As Worksheet

    Dim sheets_i As Integer

    Dim found_later_name

    On Error GoTo ContinueAddSmTab

    first_sm = Len(current_sm) = 0

    If InStr(1, name, "(SM) ") = 0 Then

      current_sm = "(SM) " & name

    Else

      current_sm = name

    End If

ContinueAddSmTab:

    Err.Clear

    On Error GoTo 0

    found_later_name = False

    If first_sm Then

        Set temp_sheet = output_book.Sheets.Add(after:=output_book.Worksheets(output_book.Worksheets.Count))

        temp_sheet.name = current_sm

    Else

        For sheets_i = 2 To output_book.Worksheets.Count

            With output_book.Worksheets(sheets_i)

                If Not .Visible = xlSheetHidden And Not .Visible = xlSheetVeryHidden Then

                    If Not InStr(1, .name, "(SM) ") = 0 Then

                        Exit For

                    End If

                End If

            End With

        Next sheets_i

        For sheets_i = sheets_i To output_book.Worksheets.Count

            With output_book.Worksheets(sheets_i)

                If Not .Visible = xlSheetHidden And Not .Visible = xlSheetVeryHidden Then

                    If StrComp(current_sm, .name) = -1 Then

                        Set temp_sheet = output_book.Sheets.Add(before:=output_book.Worksheets(sheets_i))

                        found_later_name = True

                        temp_sheet.name = current_sm

                        Exit For

                    End If

                End If

            End With

        Next sheets_i

        If Not found_later_name Then

            Set temp_sheet = output_book.Sheets.Add(after:=output_book.Worksheets(sheets_i - 1))

            temp_sheet.name = current_sm

        End If

    End If

    Call initializeCommentTab(temp_sheet)

EndAddSmTab:

End Sub

 

'Public Function getCurrentEvalType() As String

'    Dim temp_evalarray() As EvalProcEsatTypeDate

'    Dim temp_eval As EvalProcEsatTypeDate

'    If isKeyOfCollection(oAgentsEvalScores, current_agent) Then

'        temp_evalarray = oAgentsEvalScores(current_agent)

'        Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'        getCurrentEvalType = temp_eval.etype

'    Else

'        getCurrentEvalType = ""

'    End If

'End Function

 

Public Sub handleDefaultText(ByVal m_type As String, ByVal comment As String)

'    Dim meta_data As String

'    Dim garbage_text As String

'    Dim revised_mtype As String

'    Dim sheet_name As String

'    Dim second_arg As String

'    Dim agent_n As String

'    Dim eval_type As String

'    Dim score_label As String

'    Dim metric_label As String

'    Dim time_stamp As Date

'    Dim metric_score As Double

'    Dim max_score As Double

'    Dim score_pct As Double

'    Dim proc_score As Double

'    Dim esat_score As Double

'    Dim clear_cells As Worksheet

'    Dim temp_eval As EvalProcEsatTypeDate

'    Dim temp_evalarray() As EvalProcEsatTypeDate

'    Dim temp_arr() As Double

'    Dim temp_arr2() As Double

'    Dim temp_adate() As Date

'    Dim temp_coll As Collection

'    Dim temp_coll2 As Collection

'    Dim i As Integer

'    Dim array_i_offset As Integer

'    second_arg = ""

'    garbage_text = comment

'    ' separate variable because may have to change to Bad Format Rows

'    sheet_name = primary_output_tab_n

'    ' Adding values from prior row

'    If Not first_metric Then

'        If Len(SuppliedEvalType(getEvalType(output_row_offset, primary_output_tab_n))) > 0 Then

'            eval_type = getEvalType(output_row_offset, primary_output_tab_n)

'            Call addEvalType(output_row_offset + 1, eval_type, primary_output_tab_n)

'            Call addEvalType(current_agent_offset + 1, eval_type, current_agent)

'            If sm_known Then

'                Call addEvalType(current_sm_offset + 1, eval_type, current_sm)

'            End If

'        End If

'        With output_book.Worksheets(primary_output_tab_n)

'            If Len(.Range(getColumnLetter("Evaluator Satisfaction") & output_row_offset).Value) > 0 And IsNumeric(getEsatScore(output_row_offset, primary_output_tab_n)) Then

'                esat_score = getEsatScore(output_row_offset, primary_output_tab_n)

'                Call addEsatScore(output_row_offset + 1, esat_score, primary_output_tab_n)

'                Call addEsatScore(current_agent_offset + 1, esat_score, current_agent)

'                If sm_known Then

'                    Call addEsatScore(current_sm_offset + 1, esat_score, current_sm)

'                End If

'            End If

'        End With

'        time_stamp = getTimeStamp(output_row_offset, primary_output_tab_n)

'        proc_score = getProcScore(output_row_offset, primary_output_tab_n)

'        Call incrementOffsetOmnibus

'        Call addAgentNameOmnibus(current_agent)

'        Call addTimeStampOmnibus(time_stamp)

'        Call addProceduralScoreOmnibus(proc_score)

'    Else

'        If current_agent_offset < first_comment_row Then

'            current_agent_offset = first_comment_row

'        End If

'        Call addAgentNameOmnibus(current_agent)

'        'first_metric = False

'    End If

'

'    If InStr(1, comment, "||", vbBinaryCompare) = 0 And InStr(1, comment, ":", vbBinaryCompare) < InStr(1, comment, Chr(10), vbBinaryCompare) Then

'        comment = Replace(comment, "Yes:", "||Yes||", 1, 1)

'        comment = Replace(comment, "Partial:", "||Partial||", 1, 1)

'        comment = Replace(comment, "No:", "||No||", 1, 1)

'        comment = Replace(comment, "Not Likely:", "||Not Likely||", 1, 1)

'        comment = Replace(comment, "Not likely:", "||Not Likely||", 1, 1)

'        comment = Replace(comment, "Likely:", "||Likely||", 1, 1)

'        comment = Replace(comment, "Definitely:", "||Definitely||", 1, 1)

'        comment = Replace(comment, "Strongly Disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "Strongly Agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "Strongly disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "Strongly agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "strongly disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "strongly agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "Agree:", "||ESAT||3.75||", 1, 1)

'        comment = Replace(comment, "Neutral:", "||ESAT||2.5||", 1, 1)

'        comment = Replace(comment, "Disagree:", "||ESAT||1.25||", 1, 1)

'    ElseIf InStr(1, comment, "||", vbBinaryCompare) = 0 And InStr(1, comment, ":", vbBinaryCompare) < InStr(InStr(1, comment, " ", vbBinaryCompare) + 1, comment, " ", vbBinaryCompare) Then

'        comment = Replace(comment, "Strongly Disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "Strongly Agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "Strongly disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "Strongly agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "strongly disagree:", "||ESAT||0||", 1, 1)

'        comment = Replace(comment, "strongly agree:", "||ESAT||5||", 1, 1)

'        comment = Replace(comment, "Agree:", "||ESAT||3.75||", 1, 1)

'        comment = Replace(comment, "Neutral:", "||ESAT||2.5||", 1, 1)

'        comment = Replace(comment, "Disagree:", "||ESAT||1.25||", 1, 1)

'    End If

'

'    ' Resolve common error of more than two vertical pipes

'    comment = Replace(comment, "|||", "||")

'

'    ' Get first argument, Bad Formatted Rows otherwise

'    If InStr(1, comment, "||", vbBinaryCompare) > 0 And InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) > 0 Then

'        meta_data = LCase(LTrim(RTrim(Mid(comment, InStr(1, comment, "||", vbBinaryCompare) + 2, InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) - (InStr(1, comment, "||", vbBinaryCompare) + 2)))))

'        If Len(getMetricTypeRevised(m_type, meta_data)) = 0 And Len(SuppliedEvalType(meta_data)) = 0 Then

'            sheet_name = "Bad Format Rows"

'            GoTo GarbageMetaDataHandling

'        ElseIf Len(SuppliedEvalType(meta_data)) > 0 Then

'            'Call addEvalType(output_row_offset, SuppliedEvalType(meta_data), sheet_name)

'            'Call addEvalType(current_agent_offset, SuppliedEvalType(meta_data), current_agent)

'            Call fillEvalTypeUp(output_row_offset, SuppliedEvalType(meta_data), sheet_name)

'            Call fillEvalTypeUp(current_agent_offset, SuppliedEvalType(meta_data), current_agent)

'            If sm_known Then

'                Call fillEvalTypeUp(current_sm_offset, SuppliedEvalType(meta_data), current_sm)

'            End If

'            temp_evalarray = oAgentsEvalScores(current_agent)

'            oAgentsEvalScores.Remove (current_agent)

'            Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'            temp_eval.etype = SuppliedEvalType(meta_data)

'            Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'            oAgentsEvalScores.Add key:=current_agent, Item:=temp_evalarray

'        End If

'

'        revised_mtype = getMetricTypeRevised(m_type, meta_data)

'        ' Verification comment processing

'        If revised_mtype = "Verification" Then

'            If InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||") + 2, comment, "||") > 0 Then

'                second_arg = StrConv(LTrim(RTrim(Mid(comment, InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||") + 2, InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||") + 2, comment, "||") - (InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||") + 2)))), vbProperCase)

'                If Not LCase(second_arg) = "no" And Not LCase(second_arg) = "yes" Then

'                    GoTo GarbageMetaDataHandling

'                End If

'            ElseIf Not StrConv(RTrim(LTrim(meta_data)), vbProperCase) = "Verification" Then

'                second_arg = StrConv(RTrim(LTrim(meta_data)), vbProperCase)

'            Else

'                second_arg = "Yes"

'            End If

'

'            temp_evalarray = oAgentsEvalScores(current_agent)

'            oAgentsEvalScores.Remove (current_agent)

'            Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'            If second_arg = "Yes" Then

'                temp_eval.everification = True

'            Else

'                temp_eval.everification = False

'            End If

'            Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'            oAgentsEvalScores.Add key:=current_agent, Item:=temp_evalarray

'

'            If isCollectionKey("Verification", section_metric_label_qty) Then

'                Set temp_coll = section_metric_label_qty("Verification")

'                section_metric_label_qty.Remove ("Verification")

'            Else

'                Set temp_coll = New Collection

'            End If

'            If isCollectionKey("Verification", temp_coll) Then

'                Set temp_coll2 = temp_coll("Verification")

'                temp_coll.Remove ("Verification")

'            Else

'                Set temp_coll2 = New Collection

'            End If

'            If isCollectionKey(second_arg, temp_coll2) Then

'                i = temp_coll2(second_arg)

'                temp_coll2.Remove (second_arg)

'            Else

'                i = 0

'            End If

'            temp_coll2.Add key:=second_arg, Item:=i + 1

'            temp_coll.Add key:="Verification", Item:=temp_coll2

'            section_metric_label_qty.Add key:="Verification", Item:=temp_coll

'

'            verification_comment_supplied = True

'        End If

'    End If

'    ' Evaluator Satisfaction comment processing

'    If revised_mtype = "ESAT" Then

'        If InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) > 0 And InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) Then

'            If InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||") + 2, comment, "||") = 0 Then

'                GoTo GarbageMetaDataHandling

'            End If

'            second_arg = Mid(comment, InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, InStr(InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) - (InStr(InStr(1, comment, "||", vbBinaryCompare) + 2, comment, "||", vbBinaryCompare) + 2))

'            ' Different processing if score is number or label

'            If Not getRevisedEsatScore(second_arg) = 999999999 Then 'LCase(LTrim(RTrim(second_arg))) = "strongly agree" Or LCase(LTrim(RTrim(second_arg))) = "agree" Or LCase(LTrim(RTrim(second_arg))) = "neutral" Or LCase(LTrim(RTrim(second_arg))) = "disagree" Or LCase(LTrim(RTrim(second_arg))) = "strongly disagree" Then

'                has_esat = True

'                second_arg = getRevisedEsatScore(second_arg)

'                ' Stuff it in collections for final values

'                Set temp_eval = New EvalProcEsatTypeDate

'                temp_eval.esat = second_arg

'                temp_eval.edate = getTimeStamp(current_agent_offset, current_agent)

'                If Not isCollectionKey(current_agent, oAgentsEvalScores) Then

'                    ReDim temp_evalarray(0 To 0)

'                    Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'                    oAgentsEvalScores.Add Item:=temp_evalarray, key:=current_agent

'                    first_agent = False

'                Else

'                    temp_evalarray = oAgentsEvalScores(current_agent)

'                    oAgentsEvalScores.Remove (current_agent)

'                    ReDim Preserve temp_evalarray(UBound(temp_evalarray) + 1)

'                    Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'                    oAgentsEvalScores.Add key:=current_agent, Item:=temp_evalarray

'                End If

'                If (Not Not esat_scores) <> 0 Then

'                    ReDim Preserve esat_scores(LBound(esat_scores) To (UBound(esat_scores) + 1))

'                Else

'                    ReDim esat_scores(0 To 0)

'                End If

'                esat_scores(UBound(esat_scores)) = second_arg

'                max_score = getThisMaxScore(revised_mtype)

'                If sm_known Then

'                    If isKeyOfCollection(sm_esat_scores, current_sm) Then

'                        this_sm_proc = sm_esat_scores(current_sm)

'                        sm_esat_scores.Remove current_sm

'                        ReDim Preserve this_sm_proc(LBound(this_sm_proc) To UBound(this_sm_proc) + 1)

'                    Else

'                        ReDim this_sm_proc(0 To 0)

'                    End If

'                    this_sm_proc(UBound(this_sm_proc)) = second_arg

'                    sm_esat_scores.Add key:=current_sm, Item:=this_sm_proc

'                End If

'                ' Put the values on the row

'                Call addEsatScoreOmnibus(second_arg)

'                'garbage_text = getEsatScore(output_row_offset - 1, primary_output_tab_n)

'                With output_book.Worksheets(primary_output_tab_n)

'                    If IsEmpty(.Range(getColumnLetter("Evaluator Satisfaction") & (output_row_offset - 1))) Then

'                        Call fillEsatScoreUp(output_row_offset, second_arg, primary_output_tab_n)

'                        Call fillEsatScoreUp(current_agent_offset, second_arg, current_agent)

'                        If sm_known Then

'                            Call fillEsatScoreUp(current_sm_offset, second_arg, current_sm)

'                        End If

'                    End If

'                End With

'                Call addMetricScoreOmnibus(second_arg)

'                Call addMetricMaxOmnibus(max_score)

'                Call addMetricPercentageOmnibus(second_arg / max_score)

'                Call addMetricScoreLabelOmnibus(getRevisedMetricLabel(second_arg, max_score, revised_mtype))

'                If isCollectionKey("Evaluator Satisfaction", section_metric_label_qty) Then

'                    Set temp_coll = section_metric_label_qty("Evaluator Satisfaction")

'                    section_metric_label_qty.Remove ("Evaluator Satisfaction")

'                Else

'                    Set temp_coll = New Collection

'                End If

'                If isCollectionKey("ESAT", temp_coll) Then

'                    Set temp_coll2 = temp_coll("ESAT")

'                    temp_coll.Remove ("ESAT")

'                Else

'                    Set temp_coll2 = New Collection

'                End If

'                If isCollectionKey(getRevisedMetricLabel(second_arg, max_score, revised_mtype), temp_coll2) Then

'                    i = temp_coll2(getRevisedMetricLabel(second_arg, max_score, revised_mtype))

'                    temp_coll2.Remove (getRevisedMetricLabel(second_arg, max_score, revised_mtype))

'                Else

'                    Set temp_coll2 = New Collection

'                    i = 0

'                End If

'                temp_coll2.Add key:=getRevisedMetricLabel(second_arg, max_score, revised_mtype), Item:=i + 1

'                temp_coll.Add key:="ESAT", Item:=temp_coll2

'                section_metric_label_qty.Add key:="Evaluator Satisfaction", Item:=temp_coll

'            Else

'                GoTo GarbageMetaDataHandling

'            End If

'        End If

'

'

'    ElseIf Not LCase(LTrim(RTrim(m_type))) = "comment" Or revised_mtype = "Verification" Then

'        ' resolve common metadata errors

'        If Not IsNumeric(meta_data) And Not IsDate(meta_data) Then

'            If LCase(LTrim(RTrim(meta_data))) = "likley" Then

'                meta_data = "likely"

'            End If

'            If LCase(LTrim(RTrim(meta_data))) = "not likley" Then

'                meta_data = "not likely"

'            End If

'            If LCase(LTrim(RTrim(meta_data))) = "parital" Then

'                meta_data = "partial"

'            End If

'            If LCase(LTrim(RTrim(meta_data))) = "partail" Then

'                meta_data = "partial"

'            End If

'            If LCase(LTrim(RTrim(meta_data))) = "no likely" Then

'                meta_data = "not likely"

'            End If

'

'            ' resolve N/A

'

'            If Not LCase(second_arg) = "yes" And Not LCase(second_arg) = "no" And Not LCase(LTrim(RTrim(meta_data))) = "n/a" And Not LCase(LTrim(RTrim(meta_data))) = "yes" And Not LCase(LTrim(RTrim(meta_data))) = "no" And Not LCase(LTrim(RTrim(meta_data))) = "partial" And Not LCase(LTrim(RTrim(meta_data))) = "not likely" And Not LCase(LTrim(RTrim(meta_data))) = "likely" And Not LCase(LTrim(RTrim(meta_data))) = "definitely" Then

'                GoTo GarbageMetaDataHandling

'            Else

'                Call addMetricScoreLabelOmnibus(metric_label)

'            End If

'        End If

'        max_score = getThisMaxScore(revised_mtype, meta_data)

'        If Not getMetricSection(revised_mtype) = "Verification" And Not getMetricSection(revised_mtype) = "Evaluator Satisfaction" And Not revised_mtype = "Comment" And Not revised_mtype = "Hold Comment" And Not revised_mtype = "Business Comment" Then

'            If IsNumeric(meta_data) Then

'                metric_label = StrConv(getRevisedMetricLabel(meta_data, max_score, revised_mtype), vbProperCase)

'                metric_score = meta_data

'            Else

'                metric_label = StrConv(meta_data, vbProperCase)

'                If metric_label = "N/a" Then

'                    metric_label = "N/A"

'                End If

'                metric_score = getRevisedMetricScore(meta_data, max_score, revised_mtype)

'            End If

'            If isCollectionKey(getMetricSection(revised_mtype), section_metric_label_qty) Then

'                Set temp_coll = section_metric_label_qty(getMetricSection(revised_mtype))

'                section_metric_label_qty.Remove (getMetricSection(revised_mtype))

'            Else

'                Set temp_coll = New Collection

'            End If

'            If isCollectionKey(revised_mtype, temp_coll) Then

'                Set temp_coll2 = temp_coll(revised_mtype)

'                temp_coll.Remove (revised_mtype)

'            Else

'                Set temp_coll2 = New Collection

'            End If

'            If isCollectionKey(metric_label, temp_coll2) Then

'                i = temp_coll2(metric_label)

'                temp_coll2.Remove (metric_label)

'            Else

'                i = 0

'            End If

'            temp_coll2.Add key:=metric_label, Item:=i + 1

'            temp_coll.Add key:=revised_mtype, Item:=temp_coll2

'            section_metric_label_qty.Add key:=getMetricSection(revised_mtype), Item:=temp_coll

'

'            Set temp_coll = New Collection

'            If isCollectionKey(current_agent, agents_section_scores) Then

'                Set temp_coll = agents_section_scores(current_agent)

'                agents_section_scores.Remove (current_agent)

'            Else

'                Set temp_coll = New Collection

'            End If

'            If isCollectionKey(getMetricSection(revised_mtype), temp_coll) Then

'                temp_arr = temp_coll(getMetricSection(revised_mtype))

'                temp_coll.Remove (getMetricSection(revised_mtype))

'                ReDim Preserve temp_arr(LBound(temp_arr) To UBound(temp_arr) + 1)

'            Else

'                ReDim temp_arr(0 To 0)

'            End If

'            temp_arr(UBound(temp_arr)) = metric_score

'            temp_coll.Add key:=getMetricSection(revised_mtype), Item:=temp_arr

'            agents_section_scores.Add key:=current_agent, Item:=temp_coll

'            If getMetricSection(revised_mtype) = "Business Expectations" Then

'                second_arg = metric_score

'                has_esat = True

'                ' Stuff it in collections for final values

'                If Not isCollectionKey(current_agent, oAgentsEvalScores) Then

'                    Set temp_eval = New EvalProcEsatTypeDate

'                    temp_eval.esat = second_arg

'                    temp_eval.edate = getTimeStamp(current_agent_offset, current_agent)

'                    ReDim temp_evalarray(0 To 0)

'                    Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'                    oAgentsEvalScores.Add Item:=temp_evalarray, key:=current_agent

'                    first_agent = False

'                Else

'                    temp_evalarray = oAgentsEvalScores(current_agent)

'                    oAgentsEvalScores.Remove (current_agent)

'                    Set temp_eval = temp_evalarray(UBound(temp_evalarray))

'                    If temp_eval.edate = getTimeStamp(current_agent_offset, current_agent) Then

'                        temp_eval.esat = temp_eval.esat + second_arg

'                    Else

'                        ReDim Preserve temp_evalarray(UBound(temp_evalarray) + 1)

'                        Set temp_eval = New EvalProcEsatTypeDate

'                        temp_eval.esat = second_arg

'                        temp_eval.edate = getTimeStamp(current_agent_offset, current_agent)

'                    End If

'                    Set temp_evalarray(UBound(temp_evalarray)) = temp_eval

'                    oAgentsEvalScores.Add key:=current_agent, Item:=temp_evalarray

'                End If

'            End If

'        Else

'            metric_score = getRevisedMetricScore(meta_data, max_score, revised_mtype)

'            metric_label = StrConv(getRevisedMetricLabel(meta_data, max_score, revised_mtype), vbProperCase)

'        End If

'        Call addMetricScoreOmnibus(metric_score)

'        Call addMetricScoreLabelOmnibus(metric_label)

'        If max_score > 0 Then

'            Call addMetricMaxOmnibus(max_score)

'            Call addMetricPercentageOmnibus(metric_score / max_score)

'        End If

'    End If

'    comment = stripLeadTrailNewline(LTrim(RTrim(Mid(comment, InStrRev(comment, "||", Compare:=vbBinaryCompare) + 2, Len(comment) - (InStrRev(comment, "||", Compare:=vbBinaryCompare) + 1)))))

'

'

'    If Len(revised_mtype) = 0 Then

'GarbageMetaDataHandling:

'        ' Not an error, just a jump point

'        garbage_row_offset = garbage_row_offset + 1

'        agent_n = current_agent

'        eval_type = getEvalType(output_row_offset - 1, current_agent)

'        If isCollectionKey(current_agent, sp_evaldate_collection) Then

'            temp_adate = sp_evaldate_collection(current_agent)

'            time_stamp = temp_adate(UBound(temp_adate))

'        Else

'            time_stamp = getTimeStamp(output_row_offset, current_agent)

'        End If

'        If Len(revised_mtype) = 0 Then

'            revised_mtype = getMetricTypeRevised(m_type, meta_data)

'            If Len(revised_mtype) = 0 Then

'                On Error GoTo OtherUnknown

'                If Not InStr(1, comment, "ESAT") = 0 Then

'                    revised_mtype = "ESAT"

'                ElseIf Not InStr(1, comment, "Ver") = 0 And Not InStrRev(comment, "Ver", InStr(1, comment, "|")) = 0 Then

'                    revised_mtype = "Verification"

'                Else

'OtherUnknown:

'                    revised_mtype = "OTHER-UNKNOWN"

'                End If

'                Err.Clear

'                On Error GoTo 0

'            End If

'        End If

'        metric_score = getMetricScore(output_row_offset, current_agent)

'        max_score = getMetricMax(output_row_offset, current_agent)

'        score_pct = getMetricPercent(output_row_offset, current_agent)

'        proc_score = getProcScore(output_row_offset, current_agent)

'        esat_score = getEsatScore(output_row_offset, current_agent)

'        sheet_name = "Bad Format Rows"

'        Call addAgentName(garbage_row_offset, agent_n, sheet_name)

'        Call addEvalType(garbage_row_offset, eval_type, sheet_name)

'        Call addTimeStamp(garbage_row_offset, time_stamp, sheet_name)

'        Call addMetricScore(garbage_row_offset, metric_score, sheet_name)

'        Call addMetricMax(garbage_row_offset, max_score, sheet_name)

'        Call addMetricPercent(garbage_row_offset, score_pct, sheet_name)

'        Call addProcScore(garbage_row_offset, proc_score, sheet_name)

'        Call addEsatScore(garbage_row_offset, esat_score, sheet_name)

'        Call addComment(garbage_row_offset, garbage_text, sheet_name)

'        Call addMetricType(garbage_row_offset, revised_mtype, sheet_name)

'        If m_type = "Comment" Then

'            Call addMetricTypeOmnibus(m_type)

'        Else

'            Call addMetricTypeOmnibus(getMetricTypeRevised(m_type, "Yes"))

'        End If

'        Call addCommentOmnibus(garbage_text)

'        first_metric = False

'

'    Else

'        Call addMetricType(output_row_offset, revised_mtype, primary_output_tab_n)

'        Call addMetricType(current_agent_offset, revised_mtype, current_agent)

'        Call addComment(output_row_offset, comment, sheet_name)

'        Call addComment(current_agent_offset, comment, current_agent)

'        If sm_known Then

'            Call addMetricType(current_sm_offset, revised_mtype, current_sm)

'            Call addComment(current_sm_offset, comment, current_sm)

'        End If

'

'        first_metric = False

'    End If

   

End Sub

 

Private Sub initializeCommentTab(ByRef comment_tab As Worksheet)

    Dim ict_index As Integer

    Dim labels

    labels = Array("Agent", "Metric Type", "Comment", "Evaluation Type", "Time Stamp", "Metric Score Label", "Metric Score", "Maximum Metric Score", "Metric Percentage", "Client Satisfaction", "Business Expectations")

   

    For ict_index = UBound(labels) To LBound(labels) Step -1

        Call addHeaderLabel(comment_tab, labels(ict_index))

    Next ict_index

 

   

    With comment_tab.Range(getColumnLetter(first_clmn_label) & first_header_row & ":" & getColumnLetter(last_clmn_label) & first_header_row).Interior

        .PatternColorIndex = xlAutomatic

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

        .PatternTintAndShade = 0

    End With

    With comment_tab.Range(getColumnLetter(first_clmn_label) & first_header_row & ":" & getColumnLetter(last_clmn_label) & first_header_row).Font

        .ThemeColor = xlThemeColorDark1

        .TintAndShade = 0

    End With

    With comment_tab.Range(getColumnLetter(first_clmn_label) & first_header_row & ":" & getColumnLetter(last_clmn_label) & first_header_row).Font

        .name = "Tahoma"

        .Size = 10

        .Strikethrough = False

        .Superscript = False

        .Subscript = False

        .OutlineFont = False

        .Shadow = False

        .Underline = xlUnderlineStyleNone

        .ThemeColor = xlThemeColorDark1

        .TintAndShade = 0

        .ThemeFont = xlThemeFontNone

    End With

    comment_tab.Range(getColumnLetter(first_clmn_label) & first_header_row & ":" & getColumnLetter(last_clmn_label) & first_header_row).Font.Bold = True

End Sub

 

Private Sub addHeaderLabel(ByRef ws As Worksheet, ByVal label As String)

    With ws

        .Range(getColumnLetter(label) & first_header_row).Value = label

    End With

End Sub

 

' Averages the Procedural and Evaluator Satisfaction scores for the supplied agent name and adds them to that agent's sheet.

Private Sub addAgentAverages(ByVal agent_name As String, ByVal avg_type As String, dPrimary As Double, dSecondary As Double, isSecondaryValid As Boolean)

  Dim r As Range

  Dim sWorkingName As String

  If LCase(avg_type) = "sm" And InStr(1, agent_name, "(SM) ") = 0 Then

        sWorkingName = "(SM) " & agent_name

  Else

    sWorkingName = agent_name

  End If

  With output_book.Worksheets(sWorkingName)

    Set r = .Range(getColumnLetter("Procedural Score") & .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).row).offset(avg_row_offset, 0)

      r.Value = dPrimary

      On Error GoTo NoEsatScoresOmni

      If isSecondaryValid Then

        r.offset(0, 1).Value = dSecondary

      Else

        r.offset(0, 1).Value = "Unable to Calculate. Please see Bad Format Rows tab."

      End If

NoEsatScoresOmni:

      Err.Clear

      On Error GoTo 0

    If truncate_numbers Then

      r.NumberFormat = "0.00"

      r.offset(0, 1).NumberFormat = "0.00"

    End If

    If LCase(avg_type) = "agent" Then

      r.offset(1, 0).Value = getPrimaryScoreName & " Average"

      r.offset(1, 1).Value = getSecondaryScoreName & " Average"

    ElseIf LCase(avg_type) = "sm" Then

      r.offset(1, 0).Value = getPrimaryScoreName & " SM Team Average"

      r.offset(1, 1).Value = getSecondaryScoreName & " SM Team Average"

    Else

      r.offset(1, 0).Value = getPrimaryScoreName & " Total Average"

      r.offset(1, 1).Value = getSecondaryScoreName & " Total Average"

    End If

     

  End With

End Sub

 

Public Function isAgentsFirstEvalPopulated(Optional ByVal name As String) As Boolean

'    If IsMissing(name) Or Len(name) < 1 Then

'        name = current_agent

'    End If

'    isAgentsFirstEvalPopulated = isCollectionKey(name, sp_evaldate_collection)

End Function

 

'Private Sub addCurrentAgentEsatAvg(ByVal agent_name As String)

'    Dim sub_total As Double

'    Dim r As Range

'    Dim i As Integer

'    sub_total = 0

'    For i = UBound(current_agent_esat_scores) To LBound(current_agent_esat_scores) Step -1

'        sub_total = sub_total + current_agent_esat_scores(i)

   ' Next i

  '  With output_book.Worksheets(agent_name)

'       Set r = .Range(getColumnLetter("Evaluator Satisfaction") & .Range(getColumnLetter(first_clmn_label) & "1").End(xlDown).offset(3, 0).Row)

'        '.Value = sub_total / (UBound(current_agent_esat_scores) + 1)

        'agent_esat_avgs.Add key:=agent_name, Item:=sub_total / (UBound(current_agent_esat_scores) + 1)

  '      r.offset(1, 0).Value = "Evaluator Satisfaction Average"

'   End With

'

'End Sub

 

Public Function containsNonNumericCharacters(ByVal str As String) As Boolean

    If Not InStr(1, str, "A") = 0 Or Not InStr(1, str, "B") = 0 Or Not InStr(1, str, "C") = 0 Or Not InStr(1, str, "D") = 0 Or Not InStr(1, str, "E") = 0 Or Not InStr(1, str, "F") = 0 Or Not InStr(1, str, "G") = 0 Or Not InStr(1, str, "H") = 0 Or Not InStr(1, str, "I") = 0 Or Not InStr(1, str, "J") = 0 Or Not InStr(1, str, "K") = 0 Or Not InStr(1, str, "L") = 0 Or Not InStr(1, str, "M") = 0 Or Not InStr(1, str, "N") = 0 Or Not InStr(1, str, "O") = 0 Or Not InStr(1, str, "P") = 0 Or Not InStr(1, str, "Q") = 0 Or Not InStr(1, str, "R") = 0 Or Not InStr(1, str, "S") = 0 Or Not InStr(1, str, "T") = 0 Or Not InStr(1, str, "U") = 0 Or Not InStr(1, str, "V") = 0 Or Not InStr(1, str, "W") = 0 Or Not InStr(1, str, "X") = 0 Or Not InStr(1, str, "Y") = 0 Or Not InStr(1, str, "Z") = 0 Or Not InStr(1, str, "a") = 0 Or Not InStr(1, str, "b") = 0 Or Not InStr(1, str, "c") = 0 Or Not InStr(1, str, "d") = 0 Or Not InStr(1, str, "e") = 0 Or Not InStr(1, str, "f") = 0 Or Not InStr(1, str, "g") = 0 Then

        containsNonNumericCharacters = True

    ElseIf Not InStr(1, str, "h") = 0 Or Not InStr(1, str, "i") = 0 Or Not InStr(1, str, "j") = 0 Or Not InStr(1, str, "k") = 0 Or Not InStr(1, str, "l") = 0 Or Not InStr(1, str, "m") = 0 Or Not InStr(1, str, "n") = 0 Or Not InStr(1, str, "o") = 0 Or Not InStr(1, str, "p") = 0 Or Not InStr(1, str, "q") = 0 Or Not InStr(1, str, "r") = 0 Or Not InStr(1, str, "s") = 0 Or Not InStr(1, str, "t") = 0 Or Not InStr(1, str, "u") = 0 Or Not InStr(1, str, "v") = 0 Or Not InStr(1, str, "w") = 0 Or Not InStr(1, str, "x") = 0 Or Not InStr(1, str, "y") = 0 Or Not InStr(1, str, "z") = 0 Then

        containsNonNumericCharacters = True

    Else

        containsNonNumericCharacters = False

    End If

End Function

 

Private Sub cbxFullMonth_Click()

    If cbxFullMonth.Value Then

        cbxQualityRanking.enabled = True

        sltEvalMonth.enabled = True

        sltEvalYear.enabled = True

        dtpBeginDate.enabled = False

        dtpEndDate.enabled = False

    Else

        sltEvalMonth.enabled = False

        sltEvalYear.enabled = False

        cbxQualityRanking.enabled = False

        dtpBeginDate.enabled = True

        dtpEndDate.enabled = True

    End If

End Sub

 

Private Sub btnReset_Click()

    Unload Me

    End

End Sub

 

Private Sub cbxIsClientExperience_Click()

    If Not cbxIsClientExperience.Value Then

        cbxCRD.Value = False

        cbxCRD.enabled = False

        cbxNNA.Value = False

        cbxNNA.enabled = False

    End If

End Sub
