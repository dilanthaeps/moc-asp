<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Delete Screen
	'	Template Path	:	.../moc_delete.asp
	'	Functionality	:	To MOC information
	'	Called By		:	../moc_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================

	dim i,v_message_success,v_message_fail,message
	on error resume next
	For each i in Request.Form("v_deleteditems")
		strSql="DELETE FROM moc_master where moc_id=" & i
		
		err.Clear
		connObj.Execute(strSql)
		if err.number=0 then
			v_message_success = v_message_success & " MOC Detail <i>: " & i & "</i> deleted successfully !<br>"
		elseif InStr(1,err.Description,"ORA-02292")>0 then
			v_message_fail = v_message_fail & " MOC Detail <i>: " & i & "</i> could not be deleted because inspections are assigned to it !<br>"
		else
			v_message_fail = v_message_fail & " MOC Detail <i>: " & i & "</i> could not be deleted !<br>"
		end if
	next      
	message = Server.URLEncode(v_message_fail & "<br>" & v_message_success)
	connObj.close
	set connObj=nothing
	Response.Redirect "moc_maint.asp?v_message=" & message
%>   


