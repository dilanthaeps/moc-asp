	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	Dept Delete Screen
	'	Template Path	:	.../dept_delete.asp
	'	Functionality	:	To Delete Department information
	'	Called By		:	../dept_delete.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	v_message=""
	dim i
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_system_parameters where sys_para_id="
		strSql=strSql & "'"&i&"'" 
		connObj.Execute(strSql)
		v_message = v_message&" System Parameter <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	Response.Redirect "system_parameter_maint.asp?v_message="&message
%>   


