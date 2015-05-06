	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Agent Delete Screen
	'	Template Path	:	.../agent_delete.asp
	'	Functionality	:	To MOC Agent information
	'	Called By		:	../agent_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	v_message=""
	dim i
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_agents_master where agent_id="
		strSql=strSql & "'"&i&"'" 
		connObj.Execute(strSql)
		v_message = v_message&" MOC Agent Detail <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	Response.Redirect "agent_maint.asp?v_message="&message
%>   


