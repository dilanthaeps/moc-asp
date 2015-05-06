<%	'===========================================================================
	'	Template Name	:	MOC Menu Items save
	'	Template Path	:	menu_items_save.asp
	'	Functionality	:	To save the menu items entered or edited
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<% 	if request("v_menu_sub_grp_id")<>"" then
		v_menu_sub_grp_id=request("v_menu_sub_grp_id")
	else
		v_menu_sub_grp_id="null"
	end if
	'if request("menu_item")<>"N" then
	'	v_menu_sub_grp_id="null"
	'end if
 
   Idval = Request("v_menu_id")
   if Idval = "" then
     'strSql = "INSERT INTO user_table(id,name,email,tel_no,appointment,department,spare_check,store_check,forwarding_check,drydock_check,capeq_check,repair_check) Values('" & request("port") & "', '" & request("country") & "', '" & request("airport_code") & "', '" & request("port_code") & "')"
     strSql = "insert into moc_menu_items (menu_id, menu_path, menu_desc, menu_grp_id, sort_order,menu_item,grp_menu_id) values (seq_vpd_menu_items.nextval,'"& request("v_menu_path") &"','"&request("v_menu_desc")&"','"&request("v_menu_grp_id")&"',"&request("v_sort_order")&",'"&request("menu_item")&"',"&v_menu_sub_grp_id&")"
     'Response.Write "<br>"&strSql
     v_message = "Menu Item Created Successfully<br>" 
   else
	strSql = "update moc_menu_items set menu_path = '"&request("v_menu_path")&"',menu_desc='"&request("v_menu_desc")&"',menu_grp_id='"&request("v_menu_grp_id")&"',sort_order="&request("v_sort_order")&",menu_item='"&request("menu_item")&"' ,grp_menu_id="&v_menu_sub_grp_id&" where menu_id="&request("v_menu_id")
	'strSql = strSql & "'" & Idval & "'"
	v_message = "Menu Item Updated Successfully"
   end if	
   'Response.Write strSql
   connObj.Execute(strSql)
   connObj.Close
   set connObj=nothing
   
   'Response.Redirect "menu_items_maint.asp?v_message="&v_message
%>
<SCRIPT LANGUAGE="JavaScript">
	self.parent.opener.document.form1.action = "menu_items_maint.asp?v_message=<% =v_message %>";
	//alert(self.parent.opener.document.v_form.action);
	self.close();
	self.parent.opener.document.form1.target = "";
	self.parent.opener.document.form1.submit();
</SCRIPT>
</BODY>
</HTML>
