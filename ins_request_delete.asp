<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Inspector Delete Screen
	'	Template Path	:	.../inspector_delete.asp
	'	Functionality	:	To MOC Inspector information
	'	Called By		:	../inspector_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	dim i,v_message
	v_message=""
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_inspectors where inspector_id="
		strSql=strSql & "'"&i&"'" 
		connObj.Execute(strSql)
		v_message = v_message&" MOC Inspector Detail <i>: " &i&"</i> deleted successfully !<br>"
	next      
	v_message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	Response.Redirect "inspector_maint.asp?v_message=" & v_message
%>   


