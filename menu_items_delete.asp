<%	'===========================================================================
	'	Template Name	:	MOC Menu Items Delete  
	'	Template Path	:	menu_items_delete.asp
	'	Functionality	:	To delete the Menu items
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->

<% 
   dim i
   v_message = ""
   'Response.write Request.Form("v_deleteditems") & "<br>"
   For each i in Request.Form("v_deleteditems")
      strSql="DELETE FROM moc_menu_items where menu_id ="
      strSql=strSql & "'"&i&"'" 
      v_message = v_message&"Menu id :<i>"&i&"</i> deleted Successfully !!<br>"
      'Response.Write strSql
      ConnObj.Execute(strSql)
   next      
   
   ConnObj.Close
   Set ConnObj=nothing
%>   
<% Response.Redirect "menu_items_maint.asp?v_message="&v_message %>
