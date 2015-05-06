<%	'===========================================================================
	'	Template Name	:	MOC Menu Group Maintenance
	'	Template Path	:	menu_grp_maint.asp
	'	Functionality	:	To show the list of Menu Groups  available
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	26th August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = false
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<SCRIPT LANGUAGE="JavaScript">
function addEditRecord(menuGroupID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 600) + ',height=' + (screen.height - 600)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("menu_grp_entry.asp?v_menu_grp_id=" + menuGroupID, "menuGroupAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<title>Menu Group Maintenance</title>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4>Menu Group Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>	
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>   
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New Menu Group
<% 
   strSql="select a.menu_grp_id,a.menu_grp_name,to_char(a.create_date,'dd/mm/yyyy') create_date,a.sort_order from moc_menu_groups a order by sort_order"
   Set rsObj=connObj.Execute(strSql)
%>
<form name=form1 action=menu_grp_delete.asp method=post>
	<h4>Active Menu Groups</h4>
	<table>
		<tr>
			<td class=tableheader>Menu Group Name</td>
			<td class=tableheader>Create Date</td>
			<td class=tableheader>Sort Order</td>
		</tr>	
		<% if not(rsObj.eof or rsObj.bof) then
		dim c,cclass,r_count
		c=0
		r_count=0
		while not rsObj.EOF 
		if c=0 then
		cclass="columncolor2"
		c=1
		elseif c=1 then
		cclass="columncolor3"
		c=0
		end if
		%>
		   
		<tr>
			<td class="<%=cclass%>">
				<INPUT type="checkbox"   name=v_deleteditems value=<%=rsObj("menu_grp_id")%>>
				<a href="JavaScript:addEditRecord('<%=rsObj("menu_grp_id")%>')">
				<%=rsObj("menu_grp_name")%></a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("create_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("sort_order")%>&nbsp;</td>
	    </tr>
       <%rsObj.MoveNext
		 r_count=r_count+1
         wend
         Response.Write "<tr><td colspan=3 align=left><b>Record Count :</b>"&r_count&"</td></tr>"
         else
         Response.Write "<tr><td colspan=3 align=center><b> No Data Found </b></td></tr>"
         end if
       %>  

  </table>
  <p></p>
  <INPUT type="submit" value="Delete" id=submit1 name=submit1><input type=button value=Refresh onclick="Javascript:location.reload();">

</form>
</table>   
       <!--#include file="common_footer.asp"-->   
</body>
</html>