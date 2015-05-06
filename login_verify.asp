<%	'===========================================================================
	'	Template Name	:	Login Verificaiton
	'	Template Path	:	login_verify.asp
	'	Functionality	:	To verify the login credentials
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	if   v_page <> "login_process.asp" then
		'if not session("moc_user_id")>"" then
		'Response.Buffer = true
		'Response.Write Request.ServerVariables("path_info") & "<br>"
		 'Response.Redirect "user_login.asp?v_message=Please+login+to+access+the+system"
		'end if
	end if
%>