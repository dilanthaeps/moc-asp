<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%
dim DOC_ID,cnt
DOC_ID = Request.QueryString("DOC_ID")

connObj.execute "update MOC_DOCUMENTS set deleted='Y', deleted_by='" & left(USER,15) & "', deleted_date=sysdate where doc_id=" & DOC_ID, cnt

%>
<html>
<head>
<title>Delete MOC Document</title>
<link REL="stylesheet" HREF="moc.css"></link>
<script language="javascript" type="text/javascript">
function windowOnload(){
	setTimeout(closeMe,2000);
}
function closeMe(){
	//try{
		window.opener.history.go(0);
	//}catch(){}
	window.close();
}
</script>
</head>
<body style="text-align:center" onload="windowOnload()">
<div style="color:red;font-size:16px;font-weight:bold;padding-top:30px;">
<%=cnt%> document deleted successfully
</div>
</body>
</html>
