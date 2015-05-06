<%@ Language=VBScript %>
<%
	Set fso = Createobject("Scripting.FileSystemObject")

	'Set docs_dir = fso.GetFolder("d:\wls\docs")
	
	Response.Write server.MapPath("docs")

	'Set docs_dir = Nothing
	Set fso = Nothing
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
</BODY>
</HTML>
