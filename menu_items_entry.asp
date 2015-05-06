<%	'===========================================================================
	'	Template Name	:	MOC Menu Group Maintenance
	'	Template Path	:	menu_items_entry.asp
	'	Functionality	:	To add or edit menu items
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<% Dim Idval
	Idval = Request.QueryString("v_menu_id")
	if Idval <> "" then
	   v_mode="edit"
	   v_header="Update Menu Items Detils"
%>

<%
	   strSql = "Select * from moc_menu_items where menu_id="
	   strSql = strSql & "'" & Idval & "'"
	   Set rsObj = connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Add New Menu Item"
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="VBScript" runat=server>
	   function SFIELD(fname)
	      if v_mode="edit" then
				rsObj.MoveFirst
	      		Do Until rsObj.EOF
					v_tem = rsObj(cstr(fname))
					rsObj.MoveNext
				Loop
	         SFIELD=v_tem

	      else
	         SFIELD = ""
	      end if
	   End function
</script>
<SCRIPT LANGUAGE="JAVASCRIPT">
	  function validate_fields ()
		  {
			if (document.form1.v_menu_desc.value == "")
			{
				alert ("Enter Menu Item value");
				document.form1.v_menu_desc.focus();
				return false;
			}
			if(document.form1.v_menu_desc.value.charAt(0)==" ")
			{
				alert("Name cannot start with blank");
				document.v_menu_desc.focus();
				return false;
			}
			if (document.form1.v_menu_path.value == "")
			{
				alert ("Enter Menu path value");
				document.form1.v_menu_path.focus();
				return false;
			}
			if(document.form1.v_menu_path.value.charAt(0)==" ")
			{
				alert("Menu Path cannot start with blank");
				document.v_menu_path.focus();
				return false;
			}
			if(document.form1.v_menu_grp_id.selectedIndex == 0) {
				alert("Choose Menu Group Value");
				document.form1.v_menu_grp_id.focus();
				return false;
			}

  		  }
</SCRIPT>

<TITLE>Tanker Pacific - MOC - Menu Item Entry/Edit</TITLE>
</HEAD>
<BODY>
<h3><%= v_header %></h3>
<form name=form1  action=menu_items_save.asp method=post  OnSubmit="return validate_fields();">
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1>
  <% if v_mode = "edit" then %>
  <TR>
    <TD class=tableheader>
		Menu Items Id
	</TD>
    <TD class=columncolor><INPUT type=text class=textyellowcolor name=menu_id readonly value="<%=(SFIELD("menu_id"))%>"></TD></TR>
  <% end if   %>
    <input type=hidden name=v_menu_id value="<%=(SFIELD("menu_id"))%>">
  <TR>
    <TD class=tableheader>Menu Description</TD>
    <TD class=columncolor><INPUT type=text name=v_menu_desc value="<%=(SFIELD("menu_desc"))%>" size=50 maxlength=200><font color=red>*</font></TD></TR>
  <TR>
    <TD class=tableheader>Menu Path</TD>
    <TD class=columncolor><INPUT type=text name=v_menu_path value="<%=(SFIELD("menu_path"))%>" size=50  maxlength=200><font color=red>*</font></TD></TR>
  <TR>
    <TD class=tableheader>Menu Group</TD>
    <TD class=columncolor>
    <select name=v_menu_grp_id>
		<option value=<%=null%>></option>
		<%
			strSql4="select menu_grp_id,menu_grp_name from moc_menu_groups order by menu_grp_name"
			set rsObj4=connObj.execute(strSql4)
			if not(rsObj4.eof or rsObj4.bof) then
				while not rsObj4.eof
		%>
				<option value="<%=rsObj4("menu_grp_id")%>" <%if not(isnull(SFIELD("menu_grp_id"))) then%><%if cstr(rsObj4("menu_grp_id"))=cstr(SFIELD("menu_grp_id")) then%>selected<%end if%><%end if%>><%=rsObj4("menu_grp_name")%></option>
		<%		rsObj4.movenext
				wend
			end if
		%>
    </select>
    </TD> </TR>
    <TR>
    <TD class=tableheader>Menu Sub Group</TD>
    <TD class=columncolor>
    <select name=v_menu_sub_grp_id>
		<option value=<%=null%>></option>
		<%
			strSql5="select menu_id,menu_desc from moc_menu_items where menu_item='Y' order by menu_desc"
			set rsObj5=connObj.execute(strSql5)
			if not(rsObj5.eof or rsObj5.bof) then
				while not rsObj5.eof
		%>
				<option value="<%=rsObj5("menu_id")%>" <%if not(isnull(SFIELD("grp_menu_id"))) then%><%if cstr(rsObj5("menu_id"))=cstr(SFIELD("grp_menu_id")) then%>selected<%end if%><%end if%>><%=rsObj5("menu_desc")%></option>
		<%		rsObj5.movenext
				wend
			end if
		%>
    </select>
    </TD> </TR>
    <tr>
    <td>
    </td>
    </tr>
    <!--<INPUT type=text name=v_menu_grp_id value="<%=(SFIELD("menu_grp_id"))%>"></TD></TR>-->
    <TR>
      <TD class=tableheader>Admin Menu Item</TD>
      <TD class=columncolor> Yes &nbsp;<INPUT type="radio" <% if SFIELD("menu_item")="Y" then Response.Write("checked") end if %> name=menu_item value='Y'> &nbsp; No &nbsp;<INPUT type="radio"  <% if SFIELD("menu_item")<>"Y" then Response.Write("checked") end if %> name=menu_item value="N"></TD>
    </TR>
  <TR>
    <TD class=tableheader>Sort Order</TD>
    <TD class=columncolor>
		<%strSql5="select nvl(max(sort_order),0)+10 sort_order from moc_menu_items"
		  'Response.Write strSql5
		  set rsObj5=connObj.execute(strSql5)
		  if not(rsObj5.eof or rsObj5.bof) then
			  rsObj5.movefirst
			  set sort_order=rsObj5("sort_order")
		  end if
		%>
		<%if v_mode="edit" then%>
			<input type=text name="v_sort_order" style="width:150px" value="<%=SFIELD("sort_order")%>">
		<%else%>
			<input type=text name="v_sort_order" style="width:150px" value="<%=sort_order%>">
		<%end if%>
	</td>
</tr>
  <% if v_mode = "edit" then %>
  <TR>
    <TD class=tableheader>Create Date</TD>
    <TD class=columncolor><INPUT type=text  class=textyellowcolor name=v_create_date readonly value="<%=FormatDateTime(SFIELD("create_date"),2)%>"></TD></TR>
  <% end if   %>
   </TABLE>
  <input type=submit value=save name=submit>
  <input type=reset value=reset name=reset>
  <input type=button value=Refresh name=r_but onclick="javascript:location.reload();">
</form>
</BODY>
</HTML>
