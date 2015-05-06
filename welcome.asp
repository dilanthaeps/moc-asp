<%	'===========================================================================
	'	Template Name	:	Welcome
	'	Template Path	:	welcome.asp
	'	Functionality	:	To show the list of Inspector available
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	
'by default the system should show the inspections screen.
	Response.Redirect "ins_request_maint.asp"

	Response.Buffer = false
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
</head>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4>Welcome to  MOC System</h4>
Please Click any of the menu listed above as required.
</center>
<p></p>
 
<%
connObj.close
set connObj = nothing
%>
</body>
</html>