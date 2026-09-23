Option Explicit

' CONFIGURAZIONE: compilare questi 4 valori dopo aver creato il repository GitHub.
Private Const GITHUB_OWNER As String = "INSERISCI_USERNAME"
Private Const GITHUB_REPO As String = "travelling-tour"
Private Const GITHUB_BRANCH As String = "main"
Private Const GITHUB_FILE As String = "data.js"

Public Sub AggiornaSitoTravellingTour()
    Dim token As String, payload As String, content As String, sha As String
    token = Environ$("TRAVELLING_GITHUB_TOKEN")
    If Len(token) = 0 Then
        MsgBox "Manca la variabile ambiente TRAVELLING_GITHUB_TOKEN sul PC.", vbExclamation
        Exit Sub
    End If
    content = CreaDataJS()
    sha = GitHubFileSha(token)
    If sha = "" Then
        MsgBox "Non riesco a leggere data.js dal repository GitHub. Controlla username/repository/token.", vbCritical
        Exit Sub
    End If
    payload = "{""message"":""Aggiornamento classifica Travelling Tour"",""content"":""" & Base64UTF8(content) & """,""branch"":""" & GITHUB_BRANCH & """,""sha"":""" & sha & """}"
    If PutGitHub(token, payload) Then
        MsgBox "Sito Travelling Tour aggiornato online!", vbInformation
    End If
End Sub

Private Function CreaDataJS() As String
    Dim ws As Worksheet, lastRow As Long, i As Long, s As String
    Set ws = ThisWorkbook.Worksheets("Classifica")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    s = "window.TT_DATA={" & Chr(34) & "meta" & Chr(34) & ":{" & Chr(34) & "lastStage" & Chr(34) & ":" & UltimaTappa() & "," & Chr(34) & "players" & Chr(34) & ":" & Application.Max(0, lastRow - 1) & "}," & Chr(34) & "rows" & Chr(34) & ":["
    For i = 2 To lastRow
        If Len(Trim$(ws.Cells(i, "B").Value)) > 0 Then
            If i > 2 Then s = s & ","
            s = s & "{""pos"":" & NzNum(ws.Cells(i, "A").Value) & ",""player"":" & JsonStr(ws.Cells(i, "B").Value) & ",""matches"":" & NzNum(ws.Cells(i, "C").Value) & ",""points_pct"":" & NzNum(ws.Cells(i, "D").Value) & ",""points"":" & NzNum(ws.Cells(i, "E").Value) & ",""gw"":" & NzNum(ws.Cells(i, "F").Value) & ",""gl"":" & NzNum(ws.Cells(i, "G").Value) & ",""diff"":" & NzNum(ws.Cells(i, "H").Value) & ",""match_pct"":" & NzNum(ws.Cells(i, "I").Value) & ",""cards"":" & NzNum(ws.Cells(i, "L").Value) & "}"
        End If
    Next i
    CreaDataJS = s & "]};"
End Function

Private Function UltimaTappa() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long, m As Long
    Set ws = ThisWorkbook.Worksheets("DATABASE")
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, "F").Value) Then If CLng(ws.Cells(i, "F").Value) > m Then m = CLng(ws.Cells(i, "F").Value)
    Next i
    UltimaTappa = m
End Function

Private Function NzNum(v As Variant) As String
    If IsNumeric(v) Then NzNum = Replace(CStr(CDbl(v)), ",", ".") Else NzNum = "0"
End Function
Private Function JsonStr(v As Variant) As String
    JsonStr = """" & Replace(CStr(v), """" , "\""") & """"
End Function

Private Function GitHubFileSha(token As String) As String
    Dim http As Object, url As String, body As String, p As Long, q As Long
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & GITHUB_REPO & "/contents/" & GITHUB_FILE & "?ref=" & GITHUB_BRANCH
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.SetRequestHeader "User-Agent", "Travelling-Tour-Excel"
    http.Send
    If http.Status <> 200 Then Exit Function
    body = http.ResponseText: p = InStr(body, """sha"":""")
    If p > 0 Then p = p + 7: q = InStr(p, body, """"): If q > p Then GitHubFileSha = Mid$(body, p, q - p)
End Function

Private Function PutGitHub(token As String, payload As String) As Boolean
    Dim http As Object, url As String
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & GITHUB_REPO & "/contents/" & GITHUB_FILE
    http.Open "PUT", url, False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/vnd.github+json"
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "User-Agent", "Travelling-Tour-Excel"
    http.Send payload
    PutGitHub = (http.Status = 200 Or http.Status = 201)
    If Not PutGitHub Then MsgBox "GitHub ha restituito: " & http.Status & vbCrLf & http.ResponseText, vbCritical
End Function

Private Function Base64UTF8(text As String) As String
    Dim stm As Object, xml As Object, node As Object, b() As Byte
    Set stm = CreateObject("ADODB.Stream"): stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.WriteText text: stm.Position = 0: stm.Type = 1: b = stm.Read: stm.Close
    Set xml = CreateObject("MSXML2.DOMDocument.6.0"): Set node = xml.createElement("b64"): node.DataType = "bin.base64": node.nodeTypedValue = b: Base64UTF8 = Replace(node.Text, vbLf, "")
End Function
