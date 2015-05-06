<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%
if Request.QueryString("ID")="GRADE" then
	connObj.execute("Update MOC_INSPECTION_REQUESTS set inspection_grade=" & Request.QueryString("VALUE1") & " where request_id =" & Request.QueryString("KEYFIELD1"))
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

</BODY>
</HTML>
<script>
self.close()
</script>