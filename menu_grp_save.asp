<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Menu Group Save
	'	Template Path	:	menu_grp_save.asp
	'	Functionality	:	To Save the changes
	'	Called By		:	menu_grp_entry.asp 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	    
	if request("menu_grp_name")<>"" then
		menu_grp_name=request("menu_grp_name")
	else
		menu_grp_name=null
	end if
	if request("sort_order")<>"" then
		sort_order=request("sort_order")
	else
		sort_order=null
	end if

	Dim Idval
	Idval = Request("v_menu_grp_id")
	if Idval = "" then
		strSql = "INSERT INTO moc_menu_groups(menu_grp_id, menu_grp_name, create_date, sort_order) Values(0, '" & menu_grp_name & "', sysdate, "&sort_order&")"
		v_message = "Menu Group Created Successfully"
	else
		strSql = "Update moc_menu_groups set menu_grp_name='"&menu_grp_name&"',sort_order="&sort_order&" where menu_grp_id="
		strSql = strSql & "'" & Idval & "'"
		v_message = "Menu Group Updated Successfully"
	end if	
	   
	connObj.Execute(strSql)
	connObj.Close
	set connObj=nothing

	'Response.Redirect "menu_grp_maint.asp?v_message="&v_message
%>   
<SCRIPT LANGUAGE="JavaScript">
	self.parent.opener.document.form1.action = "menu_grp_maint.asp?v_message=<% =v_message %>";
	//alert(self.parent.opener.document.v_form.action);
	self.close();
	self.parent.opener.document.form1.target = "";
	self.parent.opener.document.form1.submit();
</SCRIPT>