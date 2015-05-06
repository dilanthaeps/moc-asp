	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC TC Delete Screen
	'	Template Path	:	.../tc_delete.asp
	'	Functionality	:	To MOC TC information
	'	Called By		:	../tc_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	v_message=""
	dim i
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_time_charterers where time_charterer_id="
		strSql=strSql & "'"&i&"'" 
		connObj.Execute(strSql)
		v_message = v_message&" MOC Time Charterer Detail <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	Response.Redirect "tc_maint.asp?v_message="&message
%>   


