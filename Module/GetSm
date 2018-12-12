Const FirstURL = "http://fwdirectory.ms.com/itsmg/slashn/webapp/cxf/fwdServices/fwd/search.json?enablesizelimit=no&function=FWDAll&keyword="

Const heirarchyURL = "http://fwdirectory.ms.com/itsmg/slashn/webapp/cxf/fwdServices/fwd/search.json?dn=msfwid%3D||FWID||,ou%3Dpeople,o%3Dmorgan+stanley"

   

Public Function getManagerName(agent_name As String) As String

    Dim GetData As New WinHttp.WinHttpRequest

    Dim HTMLDoc As New MSHTML.HTMLDocument

    Dim AllText As String

    Dim temp_obj As Object

    Dim json_array_size As String

    JsonParser.InitScriptEngine

     GetData.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = True

     GetData.Open "GET", FirstURL & agent_name, True

     GetData.send

     GetData.waitForResponse

    

     Set temp_obj = JsonParser.DecodeJsonString(GetData.responseText)

     Set temp_obj = JsonParser.GetObjectProperty(temp_obj, "people")

     json_array_size = JsonParser.GetProperty(temp_obj, "size")

    If json_array_size = "1" Then

        AllText = queryAndParse(temp_obj)

    ElseIf CInt(json_array_size) > 1 Then

       Dim j As Integer

       Dim cost_center_count As Integer

       Dim a_person As Object

       Dim cc As String

       Dim cc_match As Object

       cost_center_count = 0

       Set temp_obj = JsonParser.GetObjectProperty(temp_obj, "item")

       For j = 0 To (CInt(json_array_size) - 1)

           Set a_person = JsonParser.GetObjectProperty(temp_obj, j & "")

           cc = JsonParser.GetProperty(a_person, "costCenter")

           If cc = "4S69" Or cc = "4T02" Or cc = "4S86" Or cc = "4S71" Or cc = "4T00" Or cc = "4S96" Or cc = "4S93" Or cc = "4S80" Or cc = "4S62" Or cc = "4S61" Or cc = "4S65" Then

               cost_center_count = cost_center_count + 1

               Set cc_match = a_person

           End If

           If cost_center_count > 1 Then

               Exit For

           End If

       Next j

       If cost_center_count = 1 Then

           AllText = queryAndParse(cc_match)

       Else

            AllText = ""

       End If

    Else

        If StrComp(Left(agent_name, 1), "O") = 0 And InStr(1, agent_name, "'") = 0 Then

            AllText = getManagerName("O'" & Right(agent_name, Len(agent_name) - 1))

        Else

            AllText = ""

        End If

    End If

     getManagerName = AllText

End Function

 

Private Function queryAndParse(current As Object) As String

    Dim GetData As New WinHttp.WinHttpRequest

    Dim HTMLDoc As New MSHTML.HTMLDocument

    Dim fwid As String

    Dim heirarchy_query_string As String

    JsonParser.InitScriptEngine

    If JsonParser.isProperty(current, "item") Then

        Set current = JsonParser.GetObjectProperty(current, "item")

    End If

    If JsonParser.isProperty(current, "0") Then

        Set current = JsonParser.GetObjectProperty(current, "0")

    End If

    If JsonParser.isProperty(current, "FWID") Then

        fwid = JsonParser.GetProperty(current, "FWID")

        heirarchy_query_string = Replace(heirarchyURL, "||FWID||", fwid, , , vbBinaryCompare)

       

        GetData.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects) = True

        GetData.Open "GET", heirarchy_query_string, True

        GetData.send

        GetData.waitForResponse

       

        Set current = JsonParser.DecodeJsonString(GetData.responseText)

        Set current = JsonParser.GetObjectProperty(current, "people")

        Set current = JsonParser.GetObjectProperty(current, "item")

        Set current = JsonParser.GetObjectProperty(current, "0")

        If JsonParser.isProperty(current, "manager") Then

            Set current = JsonParser.GetObjectProperty(current, "manager")

        Else

            Set current = JsonParser.GetObjectProperty(current, "assignmentContact")

        End If

        Set current = JsonParser.GetObjectProperty(current, "0")

        AllText = JsonParser.GetProperty(current, "surname") & ", "

        queryAndParse = AllText & JsonParser.GetProperty(current, "givenName")

    Else

        queryAndParse = ""

    End If

End Function
