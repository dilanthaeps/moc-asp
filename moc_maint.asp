<%	'===========================================================================
	'	Template Name	:	MOC Maintenance
	'	Template Path	:	moc_maint.asp
	'	Functionality	:	To show the list of MOC available
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<TITLE>MOC Maintenance</TITLE>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="vbscript">
dim oColl,objSel
dim timerID,sSearch
sub window_onload
	set oColl = document.getElementsByName("moc_name")
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
function addEditRecord(MOCID)
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

	windNew = window.open("moc_entry.asp?v_moc_id=" + MOCID, "MOCAddEdit", winStats);

	windNew.focus();
}
</SCRIPT>
</head>
<body class=bgcolorlogin>
<!--#include file="menu_include.asp"-->
<center>
<h4> MOC Maintenance</h4>
<p></p>
<%  v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>
   <font color=red size=+1><%=v_mess%></font>
<% end if%>
<p>
Click <a href="JavaScript:addEditRecord('')">Here</a> to Create a New MOC
<%
   strSql = "SELECT    A.MOC_ID , A.SHORT_NAME , A.FULL_NAME, A.ENTRY_TYPE ,  A.TELEPHONE "
   strSql = strSql & " , A.FAX_NO , A.EMAIL , A.PIC ,to_char(A.CREATE_DATE,'DD-MON-YYYY') create_date "
   strSql = strSql & " , A.CREATED_BY , to_char(A.LAST_MODIFIED_DATE,'DD-MON-YYYY') last_modified_date , A.LAST_MODIFIED_BY "
   strSql = strSql & "FROM MOC_MASTER A "
   strSql = strSql & "ORDER BY UPPER(A.SHORT_NAME) "

   Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action=moc_delete.asp method=post>
	<table >
		<tr>
			<td class=tableheader>MOC&nbsp;ID&nbsp;
			<td class=tableheader>Type</td>
			<td class=tableheader>Short Name</td>
			<td class=tableheader>Full Name</td>
			<td class=tableheader>Person Incharge</td>
			<td class=tableheader>Telephone</td>
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
			<td class="<%=cclass%>">
			<INPUT type="checkbox"  name=v_deleteditems  value="<%=rsObj("moc_id")%>">
			<a href="JavaScript:addEditRecord('<%=rsObj("moc_id")%>')"><%=replace(rsObj("moc_id")," ","&nbsp;")%>
			</a></td>
			<td class="<%=cclass%>"><%=rsObj("entry_type")%>&nbsp;</td>
			<td class="<%=cclass%>" id=moc_name><%=rsObj("short_name")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("full_name")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("pic")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("telephone")%>&nbsp;</td>
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