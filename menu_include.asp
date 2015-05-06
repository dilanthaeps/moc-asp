<%	'===========================================================================
	'	Template Name	:	MOC Agent Maintenance
	'	Template Path	:	agent_maint.asp
	'	Functionality	:	To show the list of Agents available
	'	Called By		:	. 
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
dim strSql_include, rsObj_include, v_admin_menu_condition
dim rsObj_include1
strSql_include = "SELECT   A.MENU_GRP_ID , A.MENU_GRP_NAME , A.CREATE_DATE , A.SORT_ORDER "
strSql_include = strSql_include & " FROM MOC_MENU_GROUPS A order by A.SORT_ORDER"
set rsObj_include = connObj.execute(strSql_include)
%>
<style>
td.menu{
background:lightblue;
color:white;
}

table.menu
{
font-size:100%;
position:absolute;
visibility:hidden;
}
</style>
<script language="vbscript">
sub HideCombos
	set col = document.getElementsByTagName("SELECT")
	for each o in col
		if o.className="menuHide" then o.style.visibility="hidden"
	next
end sub
sub ShowCombos
	set col = document.getElementsByTagName("SELECT")
	for each o in col
		if o.className="menuHide" then o.style.visibility="visible"
	next
end sub
</script>
<script type="text/javascript">
function showmenu(elmnt)
{
	document.all(elmnt).style.visibility = "visible"
	HideCombos()
}

function hidemenu(elmnt)
{
	document.all(elmnt).style.visibility = "hidden"
	ShowCombos()
}
</script>

<table width="100%"  bgcolor=Blue>
 <tr HEIGHT="18pt">
 <%
 while not rsObj_include.eof

 %>
  <td class=tableheader width=25% onmouseover="showmenu('<%=rsObj_include("menu_grp_name")%>')" onmouseout="hidemenu('<%=rsObj_include("menu_grp_name")%>')">
   <%=rsObj_include("menu_grp_name")%><br>
   <table class="menu" id="<%=rsObj_include("menu_grp_name")%>" width="100%" bgcolor=Blue>
	<%

	v_admin_menu_condition = " AND A.MENU_ITEM = 'N' "
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Then
	if UserIsAdmin then
		v_admin_menu_condition = ""
	End If

	strSql_include = "SELECT A.MENU_ID , A.MENU_GRP_ID , A.MENU_PATH , A.MENU_DESC , A.CREATE_DATE "
	strSql_include = strSql_include & ", A.SORT_ORDER , A.MENU_ITEM , A.GRP_MENU_ID " 
	strSql_include = strSql_include & " FROM MOC_MENU_ITEMS A WHERE (A.MENU_GRP_ID ='"& rsObj_include("menu_grp_id")&"') "
	strSql_include = strSql_include & v_admin_menu_condition
	strSql_include = strSql_include & " order by a.sort_order"
	set rsObj_include1=connObj.Execute(strSql_include)
	if not (rsObj_include1.bof or rsObj_include1.eof) then
	while not rsObj_include1.eof
	%>
	<tr HEIGHT="18pt">
		<td WIDTH="100%" class="menu">
			<a href="<%=rsObj_include1("menu_path")%>"><font color=blue size=-2><b><%=rsObj_include1("menu_desc")%></b></font></a>
		</td>
	</tr>
   <%
   rsObj_include1.movenext
   wend
   end if
   %>
   </table>
  </td>
<%
rsObj_include.movenext
wend
%>
 </tr> 
</table>
