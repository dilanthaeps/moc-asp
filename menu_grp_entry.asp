<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Menu Group Entry/Edit
	'	Template Path	:	menu_grp_entry.asp
	'	Functionality	:	To enter or edit the Menu Groups for MOC
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	
	Dim Idval
	Idval = Request.QueryString("v_menu_grp_id")
	if Idval <> "" then
		v_mode="edit"
		v_header="Update Menu Group Detils"
		strSql = "Select menu_grp_id,menu_grp_name,to_char(create_date,'dd/mm/yyyy') create_date,sort_order from moc_menu_groups where menu_grp_id="
		strSql = strSql & "'" & Idval & "'"
		Set rsObj = connObj.Execute(strSql)
	else 
		v_mode="Add"
		v_header="Add New Menu Group"
	end if
%>
<HTML>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
	  function validate_fields (thisForm) 
		  {
			//Menu Group Name blank check
			if (thisForm.menu_grp_name.value == "")
			{
				alert ("Enter Menu Group Name value");
				thisForm.menu_grp_name.focus();
				return false;
			}
			//Menu Group Name first character blanck check
			if(thisForm.menu_grp_name.value.charAt(0)==" ")
			{
				alert("Menu Group Name cannot start with blank");
				thisForm.menu_grp_name.focus();
				return false;
			}
  		  }
</SCRIPT>
<TITLE> Menu Group Entry Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<form name=thisform  action=menu_grp_save.asp method=post  OnSubmit="return validate_fields(this)">
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1>
  
  <TR>
    <TD class=tableheader>Menu Group Name</TD>
    <TD class=columncolor><INPUT type=text name=menu_grp_name value="<%=(SFIELD("menu_grp_name"))%>"><font color=red>*</font></TD></TR>
  <%if v_mode="edit" then%>
    <input type=hidden name=v_menu_grp_id value="<%=(SFIELD("menu_grp_id"))%>">
  <%end if%>  
 <TR>
    <TD class=tableheader>Sort Order</TD>
    <TD class=columncolor>
		<%strSql5="select nvl(max(sort_order),0)+10 sort_order from moc_menu_groups"
		  'Response.Write strSql5
		  set rsObj5=connObj.execute(strSql5)
		  if not(rsObj5.eof or rsObj5.bof) then
			  rsObj5.movefirst
			  set sort_order=rsObj5("sort_order")
		  end if
		%> 
		<%if v_mode="edit" then%> 
			<input type=text name="sort_order" style="width:150px" value="<%=SFIELD("sort_order")%>">
		<%else%>
			<input type=text name="sort_order" style="width:150px" value="<%=sort_order%>">
		<%end if%>	
	</td>
</tr>

  </TABLE>
  <input type=submit value=save name=submit>
  <input type=reset value=reset name=reset>
</form>
		 <% if v_mode="edit" then%>
			<!--#include file="common_footer.asp"-->
         <% end if%>			
</BODY>
</HTML>
