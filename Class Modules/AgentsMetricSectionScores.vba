Private agents_section_scores As Collection ' with keys for agents pointing to Collection with keys of section names (Procedural Accuracy, Call Handling, Client Experience) pointing to arrays of metric scores

Private aMetricSections() As String

Private aAgentNames() As String

Private iIndex As Integer

Private iAgentIndex As Integer

 

Private Sub Class_Initialize()

  iIndex = 0

  Set agents_section_scores = New Collection

  iAgentIndex = 0

End Sub

 

 

Public Sub addScoreToAgentMetricSection(sAgent As String, sMetricSection As String, dScore As Double)

  Dim oInnerColl As Collection

  If frmReportBuilderSubmit.isKeyOfCollection(agents_section_scores, sAgent) Then

    Set oInnerColl = agents_section_scores(sAgent)

    agents_section_scores.Remove (sAgent)

  Else

    If (Not Not aAgentNames) <> 0 Then

      If Not sAgent = aAgentNames(UBound(aAgentNames)) Then

        ReDim Preserve aAgentNames(LBound(aAgentNames) To UBound(aAgentNames) + 1)

      End If

    Else

      ReDim aAgentNames(0 To 0)

    End If

    aAgentNames(UBound(aAgentNames)) = sAgent

    Set oInnerColl = New Collection

  End If

  Call addCollectionStringKeyNewDouble(oInnerColl, sMetricSection, dScore)

  agents_section_scores.Add key:=sAgent, Item:=oInnerColl

End Sub

 

Private Sub addCollectionStringKeyNewDouble(ByRef oTheCollection As Collection, sTheKey As String, dValue As Double)

  Dim aTheDoubles() As Double

  If frmReportBuilderSubmit.isKeyOfCollection(oTheCollection, sTheKey) Then

    aTheDoubles = oTheCollection(sTheKey)

    oTheCollection.Remove (sTheKey)

    ReDim Preserve aTheDoubles(LBound(aTheDoubles) To UBound(aTheDoubles) + 1)

  Else

    ReDim aTheDoubles(0 To 0)

    If (Not Not aMetricSections) <> 0 Then

      ReDim Preserve aMetricSections(LBound(aMetricSections) To UBound(aMetricSections) + 1)

    Else

      ReDim aMetricSections(0 To 0)

    End If

    aMetricSections(UBound(aMetricSections)) = sTheKey

  End If

  aTheDoubles(UBound(aTheDoubles)) = dValue

  oTheCollection.Add key:=sTheKey, Item:=aTheDoubles

End Sub

 

Private Function getCurrentAgentEntryLength()

  Dim i As Integer

  Dim answer As Integer

  For i = iAgentIndex To LBound(aAgentNames) - 1

    answer = answer + frmReportBuilderSubmit.getArrayLength(agents_section_scores(aAgentNames(i)))

  Next i

End Function

 

Private Function getSpecificAgentEntryLength(iSpecificAgentIndex)

  Dim i As Integer

  Dim answer As Integer

  Dim oTempColl As Collection

  For i = iSpecificAgentIndex To LBound(aAgentNames) - 1

    Set oTempColl = agents_section_scores(aAgentNames(i))

    answer = answer + oTempColl.Count

  Next i

End Function

 

Public Sub resetIndex()

  iIndex = 0

  iAgentIndex = 0

End Sub

 

Public Function hasNext() As Boolean

  hasNext = iIndex < getSpecificAgentEntryLength(UBound(aAgentNames))

End Function

 

Public Sub nextEntry()

  iIndex = iIndex + 1

  If Not getCurrentAgentEntryLength() > iIndex Then

    iAgentIndex = iAgentIndex + 1

  End If

End Sub

 

Public Function getCurrentAgent() As String

  getCurrentAgent = aAgentNames(iAgentIndex)

End Function

 

Public Function getCurrentMetricSection() As String

  getCurrentMetricSection = aMetricSections(iIndex)

End Function

 

Public Function getCurrentScoreAverage() As Double

  Dim oSectionColl As Collection

  Set oSectionColl = agents_section_scores(aAgentNames(iAgentIndex))

  getCurrentScoreAverage = (UBound(oSectionColl(aMetricSections(iIndex))) - LBound(oSectionColl(aMetricSections(iIndex))) + 1)

End Function

 

Public Function getSpecificScoreAverage(sAgentName As String, sMetricSection As String) As Double

  Dim oMetricSection As Collection

  Dim aScores() As Double

  Dim el As Variant

  Dim score As Double

  Dim dSubtotal As Double

  Dim bResolved As Boolean

  If frmReportBuilderSubmit.isKeyOfCollection(agents_section_scores, sAgentName) Then

    Set oMetricSection = agents_section_scores(sAgentName)

    If frmReportBuilderSubmit.isKeyOfCollection(oMetricSection, sMetricSection) Then

      aScores = oMetricSection(sMetricSection)

      For Each el In aScores

        score = el

        dSubtotal = dSubtotal + score

      Next el

      getSpecificScoreAverage = dSubtotal / CDbl(UBound(aScores) - LBound(aScores) + 1)

      Exit Function

    Else

      bResolved = False

    End If

  Else

    bResolved = False

  End If

  If Not bResolved Then

    getSpecificScoreAverage = 0

  End If

End Function
