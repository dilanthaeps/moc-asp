<%
v_page="login_process.asp"
%>
<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Login Process
	'	Template Path	:	login_process.asp
	'	Functionality	:	To process Login
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer =false
%>

<%
strSql = "select user_id, access_level from wls_user_master where user_id='" 
strSql = strSql & Request.Form("user_id") & "' and password='"& Request.Form("password")&"'"
'Response.Write strSql
Set rsObj=connObj.Execute(strSql)
v_cnt=""
v_user_id=""
if not (rsObj.bof or rsObj.eof) then
	while not rsObj.eof
	v_cnt="1"
	v_user_id = rsObj("user_id")
	v_access_level = rsObj("access_level")
	rsObj.movenext
	wend
end if
if v_cnt="1" then
session("moc_user_id")= Request.Form("user_id")
Response.Redirect "welcome.asp"
else
	session("moc_user_id")=""
	v_message=server.URLEncode("Either User ID or Password or Both are wrong. Please re-enter correctly. <br>User ID and Password are case sensitive." )
	Response.Redirect "user_login.asp?v_message="&v_message 
end if   
%>
       <!--#include file="common_footer.asp"-->  