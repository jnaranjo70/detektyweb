<%
' Improved ASP Classic - Save Partner Requests as CSV
Option Explicit

Function CleanInput(str)
  If IsNull(str) Then str = ""
  str = Replace(str, """", "'")
  str = Replace(str, vbCrLf, " ")
  str = Replace(str, vbCr, " ")
  str = Replace(str, vbLf, " ")
  CleanInput = Trim(str)
End Function

Dim fso, file, savePath, name, company, email, comments, logline, sep, nowStr

sep = ","
savePath = Server.MapPath("partners_requests.csv")

name = CleanInput(Request.Form("name"))
company = CleanInput(Request.Form("company"))
email = CleanInput(Request.Form("email"))
comments = CleanInput(Request.Form("comments"))
nowStr = Year(Now) & "-" & Right("0" & Month(Now),2) & "-" & Right("0" & Day(Now),2) & " " & Right("0" & Hour(Now),2) & ":" & Right("0" & Minute(Now),2)

If name = "" Or company = "" Or email = "" Then
  Response.Status = "400 Bad Request"
  Response.Write "Missing required fields."
  Response.End
End If

Dim isNewFile: isNewFile = False
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(savePath) Then isNewFile = True

Set file = fso.OpenTextFile(savePath, 8, True)
If isNewFile Then
  file.WriteLine """DateTime"",""Name"",""Company"",""Email"",""Comments"""
End If

logline = """" & nowStr & """" & sep & _
          """" & name & """" & sep & _
          """" & company & """" & sep & _
          """" & email & """" & sep & _
          """" & comments & """"
file.WriteLine logline
file.Close
Set file = Nothing
Set fso = Nothing

Response.End
%>
