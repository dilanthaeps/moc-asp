	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Menu Group Delete Screen
	'	Template Path	:	.../menu_group_delete.asp
	'	Functionality	:	To Menu Group Deletion 
	'	Called By		:	../moc_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	v_mess=""
	dim i
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_menu_groups where menu_grp_id="
		strSql=strSql & "'"&i&"'" 
		connObj.Execute(strSql)
		v_mess  = v_mess &" Menu Group Detail <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	mess  = Server.URLEncode(v_mess )
	connObj.close
	set connObj=nothing
	Response.Redirect "menu_grp_maint.asp?v_mess="&mess 
%>   


