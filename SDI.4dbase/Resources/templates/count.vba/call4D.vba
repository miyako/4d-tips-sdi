Private Function getURL() As String

Set wsREST = Worksheets("REST")

vURL = ""

'Host
vHOST = wsREST.Range("HOST").Value
vHOST_len = Len(vHOST)
If vHOST_len <> 0 And Right(vHOST, 1) = "/" Then
'remove trailing "/"
    vHOST = Left(vHOST, vHOST_len - 1)
End If

'Port
VPort = wsREST.Range("PORT")
If Left(vHOST, 5) = "http:" Then
    Select Case VPort
        Case 80
            vURL = vHOST
        Case Else
            vURL = vHOST & ":" & VPort
    End Select
Else
If Left(vHOST, 6) = "https:" Then
        Select Case VPort
            Case 443
                vURL = vHOST
            Case Else
                vURL = vHOST & ":" & VPort
        End Select
    End If
End If

getURL = vURL & "/" & _
    wsREST.Range("PREFIX") & "/" & _
    wsREST.Range("VERSION") & "/"

End Function

Function updateProductCountForCode(vCount As String, vValue As String, _
rCount As Range, _
rLevel As Range, _
rUnit As Range, _
rOrder As Range)

  Dim wc4D As New WebClient
  wc4D.BaseUrl = getURL
  Dim vRequest As String

  vRequest = "updateProductCountForCode" & _
  "?" & "code=" & UrlEncode(vValue) & _
  "&" & "count=" & UrlEncode(vCount)

  Dim wrResponse As WebResponse
  Set wrResponse = wc4D.GetJson(vRequest)

  If wrResponse.StatusCode = 200 And wrResponse.Data("status") = "OK" Then
    rCount.Value = wrResponse.Data("product")("stock")
    rLevel.Value = wrResponse.Data("product")("level")
    rUnit.Value = wrResponse.Data("product")("units")
    rOrder.Value = wrResponse.Data("product")("order")
  End If


End Function

Function getProductNameForCode(vValue As String, _
rName As Range, _
rCount As Range, _
rStock As Range, _
rLevel As Range, _
rUnit As Range, _
rOrder As Range)

  Dim wc4D As New WebClient
  wc4D.BaseUrl = getURL
  Dim vRequest As String

  vRequest = "getProductNameForCode" & _
  "?" & "code=" & UrlEncode(vValue)

  Dim wrResponse As WebResponse
  Set wrResponse = wc4D.GetJson(vRequest)

  If wrResponse.StatusCode = 200 And wrResponse.Data("status") = "OK" Then
    rName.Value = wrResponse.Data("product")("name")
    rCount.ClearContents
    rStock.Value = wrResponse.Data("product")("stock")
    rLevel.Value = wrResponse.Data("product")("level")
    rUnit.Value = wrResponse.Data("product")("units")
    rOrder.Value = wrResponse.Data("product")("order")
   Else
    rName.ClearContents
    rCount.ClearContents
    rStock.ClearContents
    rLevel.ClearContents
    rUnit.ClearContents
    rOrder.ClearContents
  End If

End Function

