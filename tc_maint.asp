<%	'===========================================================================
	'	Template Name	:	MOC Time Charterer Maintenance
	'	Template Path	:	tc_maint.asp
	'	Functionality	:	To show the list of TCs available
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd August, 2002
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
function addEditRecord(tcID)
{
	var windNew;

	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'

	if (navigator.appName.indexOf("Microsoft") >= 0)
	{
		winStats += ',left=50,top=50,width=' + (screen.width - 400) + ',height=' + (screen.height - 175)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}

	windNew = window.open("tc_entry.asp?v_time_charterer_id=" + tcID, "tcAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4> MOC Time Charterer Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New MOC Time Charterer
<%
   strSql = "SELECT    A.TIME_CHARTERER_ID , A.SHORT_NAME , A.FULL_NAME ,  A.TELEPHONE "
   strSql = strSql & " , A.FAX_NO , A.EMAIL , A.PIC ,to_char(A.CREATE_DATE,'DD-MON-YYYY') create_date "
   strSql = strSql & " , A.CREATED_BY , to_char(A.LAST_MODIFIED_DATE,'DD-MON-YYYY') last_modified_date , A.LAST_MODIFIED_BY "
   strSql = strSql & " FROM MOC_TIME_CHARTERERS A "
   strSql = strSql & " ORDER BY UPPER(A.SHORT_NAME) "
   'Response.Write strSql
   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action=tc_delete.asp method=post>
	<table >
		<tr>
			<td class=tableheader>TC ID</td>
			<td class=tableheader>Short Name</td>
			<td class=tableheader>Full Name</td>
			<td class=tableheader>Telephone</td>
			<td class=tableheader>Fax No</td>
			<td class=tableheader>Person Incharge</td>
			<td class=tableheader>Last Modified Date</td>
			<td class=tableheader>Create Date Pages</td>
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
			<INPUT type="checkbox"  name=v_deleteditems  value="<%=rsObj("TIME_CHARTERER_ID")%>">
			<a href="JavaScript:addEditRecord('<%=rsObj("TIME_CHARTERER_ID")%>')"><%=replace(rsObj("TIME_CHARTERER_ID")," ","&nbsp;")%>
			</a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("short_name")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("full_name")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("telephone")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("fax_no")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("pic")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("last_modified_date")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("create_date")%>&nbsp;</td>
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