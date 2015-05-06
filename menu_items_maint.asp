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
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
   <title>Tanker Pacific - MOC - Menu Items Maintenance </title>
<SCRIPT LANGUAGE="JavaScript">
function addEditRecord(menuID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 450)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("menu_items_entry.asp?v_menu_id=" + menuID, "menuAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<body>
<!--#include file="menu_include.asp"-->
<h4> Menu Items Maintenance</h4>
<p><%=now()%></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New Menu Item
<%
   'strSql="SELECT MI.MENU_ID , MI.MENU_PATH , MI.MENU_DESC ,MI.MENU_ITEM, MG.MENU_GRP_NAME , to_char(MI.CREATE_DATE,'dd/mm/yyyy') create_date , MI.SORT_ORDER FROM moc_MENU_ITEMS MI,vpd_menu_groups mg where mi.menu_grp_id=mg.menu_grp_id(+) order by MG.SORT_ORDER,MI.MENU_ITEM desc,MI.SORT_ORDER"
   strSql="SELECT MI.MENU_ID , MI.MENU_PATH ,MG.SORT_ORDER MG_SORT_ORDER,MIS.MENU_DESC SUB_MENU_DESC, MI.MENU_DESC ,MI.GRP_MENU_ID,MI.MENU_ITEM, MG.MENU_GRP_NAME , to_char(MI.CREATE_DATE,'dd/mm/yyyy') create_date ,MIS.SORT_ORDER MIS_SORT_ORDER, MI.SORT_ORDER,MIS.MIS_MG_SORT_ORDER FROM moc_MENU_ITEMS MI, moc_menu_groups mg , (SELECT A.MENU_ID,A.MENU_DESC, A.SORT_ORDER,B.SORT_ORDER MIS_MG_SORT_ORDER FROM moc_MENU_ITEMS A, moc_MENU_GROUPS B WHERE A.MENU_ITEM='Y' AND A.MENU_GRP_ID=B.MENU_GRP_ID) MIS where mi.menu_grp_id=mg.menu_grp_id(+) AND MIS.MENU_ID (+)=MI.GRP_MENU_ID     order by MG.SORT_ORDER,MIS_MG_SORT_ORDER,MIS_SORT_ORDER,MI.SORT_ORDER ,MI.MENU_ITEM desc,MI.SORT_ORDER"
   Set rsObj=connObj.Execute(strSql)
%>
<form name=form1 action=menu_items_delete.asp method=post>
	<table>
		<tr>
			<!--<td class=tableheader>MENU ID</td>-->
			<td class=tableheader width=40%>MENU PATH</td>
			<td class=tableheader >MENU DESCRIPTION</td>
			<td class=tableheader >MENU GROUP</td>
			<td class=tableheader >MENU SUB GROUP</td>
			<td class=tableheader >ADMIN MENU</td>
			<td class=tableheader>CREATE DATE</td>
			<td class=tableheader>SORT_ORDER</td>
		</tr>
		<%
		if not(rsObj.eof or rsObj.bof) then
		dim c,cclass,r_count
		c=0
		r_count=0
		while not rsObj.EOF
		if c=0 then
		cclass="columncolor"
		c=1
		elseif c=1 then
		cclass="columncolor1"
		c=0
		end if
%>

		<tr><td class=<%=cclass%>>
			<INPUT type="checkbox"  name=v_deleteditems value="<%=rsObj("menu_id")%>">
			 <a href="JavaScript:addEditRecord('<%=rsObj("MENU_ID")%>')"><%=rsObj("MENU_PATH")%></a>
		</td>
		<td class=<%=cclass%>><%=rsObj("MENU_DESC")%></td>
		<td class=<%=cclass%>><%=rsObj("MENU_GRP_NAME")%></td>
		<td class=<%=cclass%>><%=rsObj("SUB_MENU_DESC")%></td>
		<td class=<%=cclass%>><%=rsObj("MENU_ITEM")%></td>
		<td class=<%=cclass%>><%=rsObj("CREATE_DATE")%>&nbsp;</td>
		<td class=<%=cclass%>><%=rsObj("SORT_ORDER")%>&nbsp;</td>
		</tr>
<%      rsObj.MoveNext
        r_count=r_count+1
        wend
         Response.Write "<tr><td colspan=4 align=left><b>Record Count :</b>"&r_count&"</td></tr>"
         else
         Response.Write "<tr><td colspan=4 align=center><b> No Data Found </b></td></tr>"
         end if
       %>
       <!--#include file="common_footer.asp"-->
  </table>
  <p></p>
  <INPUT type="submit" value="Delete" id=submit1 name=submit1>&nbsp;&nbsp;&nbsp;&nbsp;
  <input type=button name=r_but value="refresh" onclick="javascript:location.reload();">
</form>
</table>
<p><%=now()%></p>
</body>
</html>
