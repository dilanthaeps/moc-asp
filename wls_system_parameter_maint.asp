<%	'===========================================================================
	'	Template Name	:	System Parameter Maintenance
	'	Template Path	:	system_parameter_maint.asp
	'	Functionality	:	To show the list of system parameters available
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	21st August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
function REPEAT_STRING(STRING_TO_REPEAT,NUMBER_OF_TIMES_TO_REPEAT)
	K=""
	CTR=0
	WHILE CTR< NUMBER_OF_TIMES_TO_REPEAT 
	K=K&STRING_TO_REPEAT
	CTR=CTR+1
	WEND
	REPEAT_STRING=K
END FUNCTION

</SCRIPT>

<html>
<head>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
</head>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4> Wls System Parameters Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>
Click <a href="system_parameter_entry.asp">Here</a> to Create a New System Parameter
<%
   'strSql="select sys_para_id,para_desc,remarks,parent_id,to_char(create_date,'dd/Mon/yyyy') create_date,to_char(last_modified_date,'dd/Mon/yyyy') last_modified_date,related_asp_pages  from MOC_system_parameters order by parent_id,sys_para_id"
   strSql="select level,para_code,para_value,parent_id,to_char(create_date,'dd/Mon/yyyy') create_date,to_char(last_modified_date,'dd/Mon/yyyy') last_modified_date,related_asp_pages  from wls_system_parameters connect by prior trim(para_code)=trim(parent_id) start with trim(parent_id) is null"
   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action=system_parameter_delete.asp method=post>
	<table >
		<tr>
			<td class=tableheader>Parameter ID</td>
			<td class=tableheader>Parameter Description</td>
			<td class=tableheader>Remarks</td>
			<td class=tableheader>Parent ID</td>
			<td class=tableheader>Sort Order</td>
			<td class=tableheader>Create Date</td>
			<td class=tableheader>Last Modified Date</td>
			<td class=tableheader>Related Pages</td>
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
			<td class="<%=cclass%>"><%Response.Write repeat_string("&nbsp;",cint(rsObj("level"))*5)%><INPUT type="checkbox"  name=v_deleteditems  value="<%=rsObj("para_code")%>"><a href="system_parameter_entry.asp?v_sys_para=<%=rsObj("para_code")%>"><%=replace(rsObj("para_code")," ","&nbsp;")%></a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("para_value")%>-<%=rsObj("level")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("para_code")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("parent_id")%>&nbsp;</td>
			<td class="<%=cclass%>"><%if rsObj("level")="1" then Response.Write "&nbsp;" else Response.write rsObj("para_code")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("create_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("last_modified_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("related_asp_pages")%>&nbsp;</td>
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

 <center> <INPUT type="submit" value="Delete" id=submit1 name=submit1> </center>
<!--#include file="common_footer.asp"-->
</form>
</center>
</table>
</body>
</html>