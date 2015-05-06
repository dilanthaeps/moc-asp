<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=PaleGreen>
 
<h1>List all Server Variables </h1>
<table border=1  bgcolor=Yellow>
<tr><td><b>Server Variable Name</b></td> <td><b> Server Variable Value </b></td></tr>
<%
for each i in Request.ServerVariables 

Response.Write "<tr><td>" & i & "</td><td>" & Request.ServerVariables(i) & "&nbsp;</td></tr>"

next

%>
</table>

<h1>List all Form Request Variables </h1>
<table border=1 bgcolor=Lime>
<tr><td><b>Form Variable Name</b></td> <td><b>Form  Variable Value </b></td></tr>
<%
for each i in Request.Form

Response.Write "<tr><td>" & i & "</td><td>" & Request.Form(i) & "&nbsp;</td></tr>"

next

%>
</table>

<h1>List all Query String Request Variables </h1>
<table border=1  bgcolor=Fuchsia>
<tr><td><b>Query String Variable Name</b></td> <td><b> Query String  Variable Value </b></td></tr>
<%
for each i in Request.QueryString 

Response.Write "<tr><td>" & i & "</td><td>" & Request.QueryString(i) & "&nbsp;</td></tr>"

next

%>
</table>

<h1>List all Session Variables </h1>
<table border=1  bgcolor=Silver>
<tr><td><b>Session Variable Name</b></td> <td><b> Session Variable Value </b></td></tr>
<%

  Dim sessitem
  Dim anArray(2)
  response.write "SessionID: " & Session.SessionID & "<P>"

For Each sessitem in Session.Contents
    If IsObject(Session.Contents(sessitem)) Then
      Response.write "<tr><td>"&sessitem & "</td><td> Session object cannot be displayed </td></tr>" 
    Else
      If IsArray(Session.Contents(sessitem)) Then
         Response.write "<tr><td>"&sessitem & "</td><td> Session Array cannot be displayed </td></tr>" 
      Else
             Response.write "<tr><td>"&sessitem & "</td><td>" & Session.Contents(sessitem) & "&nbsp;</td></tr>" 
       End If
    End If
Next 


%>
</table>
<%
response.write Request.QueryString 
%>


</BODY>
</HTML>
