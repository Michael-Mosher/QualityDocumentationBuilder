Private section_metric_label_qty As Collection ' Collection with keys for the five sections, _

           Evaluator Satisfaction, Procedural Accuracy, Call Handling, Client Experience, Verification pointing to collection with names of the metrics as keys, _

           pointing to collections with the possible score labels as keys, pointing to integers that represent the count

Private aMetricSections() As String

Private aMetricLabels() As String

Private aScoreLabels() As String

Private iIndex As Integer

Private iMetricSectionIndex As Integer

Private iMetricLabelIndex As Integer

' section_metric_label_qty As Collection

 

Private Sub Class_Initialize()

  Set section_metric_label_qty = New Collection

  iIndex = 0

  iMetricSectionIndex = 0

  iMetricLabelIndex = 0

End Sub

 

Public Sub incrementQuantity(sMetricSection As String, sMetricType As String, sScoreLabel As String)

  If Not sMetricSection = "Comment" And Len(sMetricType) > 0 And Len(sScoreLabel) > 0 Then

    Dim oAColl As Collection

    If frmReportBuilderSubmit.isKeyOfCollection(section_metric_label_qty, sMetricSection) Then

      Set oAColl = section_metric_label_qty(sMetricSection)

      section_metric_label_qty.Remove (sMetricSection)

    Else

      Set oAColl = New Collection

      Call incrementStringArrayAndAppend(aMetricSections, sMetricSection)

    End If

    Call addToMetricType(oAColl, sMetricType, sScoreLabel)

    section_metric_label_qty.Add key:=sMetricSection, Item:=oAColl

  End If

End Sub

 

Private Sub addToMetricType(ByRef oTheCollection As Collection, sMetricType As String, sScoreLabel As String)

  Dim oTempColl As Collection

  If frmReportBuilderSubmit.isKeyOfCollection(oTheCollection, sMetricType) Then

    Set oTempColl = oTheCollection(sMetricType)

    oTheCollection.Remove (sMetricType)

  Else

    Set oTempColl = New Collection

    Call incrementStringArrayAndAppend(aMetricLabels, sMetricType)

  End If

  Call addToScoreLabel(oTempColl, sScoreLabel)

  oTheCollection.Add key:=sMetricType, Item:=oTempColl

End Sub

 

Private Sub addToScoreLabel(ByRef oCollection As Collection, sScoreLabel As String)

  Dim qty As Integer

  If frmReportBuilderSubmit.isKeyOfCollection(oCollection, sScoreLabel) Then

    qty = oCollection(sScoreLabel)

    oCollection.Remove (sScoreLabel)

  Else

    qty = 0

    Call incrementStringArrayAndAppend(aScoreLabels, sScoreLabel)

  End If

  qty = qty + 1

  oCollection.Add key:=sScoreLabel, Item:=qty

End Sub

 

Private Sub incrementStringArrayAndAppend(ByRef aArrayToAppend() As String, sNewElement As String)

  If (Not Not aArrayToAppend) <> 0 Then

    ReDim Preserve aArrayToAppend(LBound(aArrayToAppend) To UBound(aArrayToAppend) + 1)

  Else

    ReDim aArrayToAppend(0 To 0)

  End If

  aArrayToAppend(UBound(aArrayToAppend)) = sNewElement

End Sub

 

 

Public Function getCurrentMetricSectionEntryLength() As Integer

  Dim i As Integer

  Dim iMidIndex As Integer

  Dim iInnerIndex As Integer

  Dim answer As Integer

  Dim oMidColl As Collection

  Dim oInnerColl As Collection

  For i = iMetricSectionIndex To LBound(aMetricSections) - 1

    Set oMidColl = section_metric_label_qty(i)

    For iMidIndex = 1 To oMidColl.Count

      Set oInnerColl = oMidColl(iMidIndex)

      For iInnerIndex = 1 To oInnerColl.Count

        answer = answer + oInnerColl(iInnerIndex)

      Next iInnerIndex

    Next iMidIndex

  Next i

  getCurrentMetricSectionEntryLength = answer

End Function

 

Public Function getSpecificMetricSectionEntryLength(iSomeIndex As Integer) As Integer

  Dim i As Integer

  Dim iMidIndex As Integer

  Dim iInnerIndex As Integer

  Dim answer As Integer

  Dim oMidColl As Collection

  Dim oInnerColl As Collection

  For i = iSomeIndex To LBound(aMetricSections) - 1

    Set oMidColl = section_metric_label_qty(i)

    For iMidIndex = 1 To oMidColl.Count

      Set oInnerColl = oMidColl(iMidIndex)

      For iInnerIndex = 1 To oInnerColl.Count

        answer = answer + oInnerColl(iInnerIndex)

      Next iInnerIndex

    Next iMidIndex

  Next i

  getSpecificMetricSectionEntryLength = answer

End Function

 

Public Function getSpecificMetricLabelEntryLength(ByRef oMetricLblColl As Collection) As Integer

  getSpecificMetricLabelEntryLength = oMetricLblColl.Count

End Function

 

Public Function getSpecificScoreLabelEntryLength(ByRef oMetricLblColl As Collection, sMetricLabel As String) As Integer

  Dim oAColl As Collection

  Set oAColl = oMetricLblColl(sMetricLabel)

  getSpecificScoreLabelEntryLength = oAColl.Count

End Function

 

Public Sub resetIndex()

  iIndex = 0

  iMetricSectionIndex = 0

  iMetricLabelIndex = 0

End Sub

 

Public Function hasNext() As Boolean

  hasNext = iIndex < (UBound(aScoreLabels) - LBound(aScoreLabels) + 1)

End Function

 

Public Sub nextEntry()

  iIndex = iIndex + 1

  If Not getCurrentMetricSectionEntryLength() > iIndex Then

    iAgentIndex = iAgentIndex + 1

  End If

  'if not getCurre

End Sub

 

Public Function getCurrentMetricSection() As String

  getCurrentMetricSection = aMetricSections(iMetricSectionIndex)

End Function

 

Public Function getCurrentMetricType() As String

  getCurrentMetricType = aMetricLabels(iMetricLabelIndex)

End Function

 

Public Function getCurrentScoreLabel() As String

  getCurrentScoreLabel = aScoreLabels(iIndex)

End Function

 

Public Function getCurrentScoreLblQty() As Integer

  getCurrentScoreAverage = getSpecificQty(aMetricSections(iMetricSectionIndex), aMetricLabels(iMetricLabelIndex), aScoreLabels(iIndex))

End Function

 

Public Function getSpecificQty(sMetricSection As String, sMetricType As String, sScoreLabel As String) As Integer

  Dim oMetricTypes As Collection

  Dim oScoreLabel As Collection

  Dim bFoundMatch As Boolean

  If frmReportBuilderSubmit.isKeyOfCollection(section_metric_label_qty, sMetricSection) Then

    Set oMetricTypes = section_metric_label_qty(sMetricSection)

    If frmReportBuilderSubmit.isKeyOfCollection(oMetricTypes, sMetricType) Then

      Set oScoreLabel = oMetricTypes(sMetricType)

      If frmReportBuilderSubmit.isKeyOfCollection(oScoreLabel, sScoreLabel) Then

        getSpecificQty = oScoreLabel(sScoreLabel)

        bFoundMatch = True

      Else

        bFoundMatch = False

      End If

    Else

      bFoundMatch = False

    End If

  Else

    bFoundMatch = False

  End If

  If Not bFoundMatch Then

    getSpecificQty = 0

  End If

End Function
