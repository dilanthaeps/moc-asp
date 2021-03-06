<%	'===========================================================================
	'	Template Name	:	MOC Inspector Maintenance
	'	Template Path	:	inspecor_maint.asp
	'	Functionality	:	To show the list of Inspector available
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
<script language="vbscript">
dim oColl,objSel
dim timerID,sSearch
sub window_onload
	set oColl = document.getElementsByName("inspector_name")
end sub
sub document_onkeypress
	if window.event.keyCode<32 or window.event.keyCode>122 then exit sub
	if timer<>0 then
		if not IsEmpty(objSel) then
			objSel.style.backgroundColor = ""
			window.status = ""
		end if
		clearTimeout timerID
		timerID = setTimeout("ClearTimer",1000)
	end if
	
	sSearch = ucase(sSearch & chr(window.event.keyCode))
	window.status = sSearch
	for i=0 to oColl.length-1
		if sSearch < ucase(oColl(i).innerText) then exit for
	next
	set objSel = oColl(i)
	objSel.scrollIntoView
	objSel.style.backgroundColor = "yellow"
end sub
sub ClearTimer
	sSearch = ""
	objSel.style.backgroundColor = ""
	window.status = ""
	timerID = 0
end sub
</script>
<SCRIPT LANGUAGE="JavaScript">
function addEditRecord(inspectorID)
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

	windNew = window.open("inspector_entry.asp?v_inspector_id=" + inspectorID, "inspectorAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4> MOC Inspector Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New MOC Inspector
<%
   strSql = "SELECT    A.INSPECTOR_ID , A.SHORT_NAME , A.FULL_NAME ,  A.TELEPHONE "
   strSql = strSql & " , A.MOBILE , A.EMAIL , A.PIC ,to_char(A.CREATE_DATE,'DD-MON-YYYY') create_date "
   strSql = strSql & " , A.CREATED_BY , to_char(A.LAST_MODIFIED_DATE,'DD-MON-YYYY') last_modified_date , A.LAST_MODIFIED_BY "
   strSql = strSql & " FROM MOC_INSPECTORS A "
   strSql = strSql & " ORDER BY UPPER(A.SHORT_NAME) "
   'Response.Write strSql
   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action=inspector_delete.asp method=post>
	<table width=100%>
		<tr>
			<td class=tableheader>ID</td>
			<td class=tableheader>Short Name</td>
			<td class=tableheader>Company Name</td>
			<td class=tableheader>Telephone</td>
			<td class=tableheader>Mobile</td>
			<td class=tableheader>Email</td>
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
			<td class="<%=cclass%>" nowrap>
			<INPUT type="checkbox"  name=v_deleteditems  value="<%=rsObj("INSPECTOR_ID")%>">
			<a href="JavaScript:addEditRecord('<%=rsObj("INSPECTOR_ID")%>')"><%=replace(rsObj("INSPECTOR_ID")," ","&nbsp;")%>
			</a>
			</td>
			<td class="<%=cclass%>" width=300px id="inspector_name"><%=rsObj("short_name")%>&nbsp;</td>
			<td class="<%=cclass%>" width=250px><%=rsObj("pic")%>&nbsp;</td>
			<td class="<%=cclass%>" nowrap><%=rsObj("telephone")%>&nbsp;</td>
			<td class="<%=cclass%>" nowrap><%=rsObj("mobile")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("email")%>&nbsp;</td>
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