<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%
dim v_vessel_code,v_vessel_name,v_def_id,v_job_code
dim rs,SQL,sColor

v_vessel_code = Request.QueryString("v_vessel_code")
v_vessel_name = Request.QueryString("v_vessel_name")
v_def_id = Request.QueryString("v_def_id")
v_job_code = Request.QueryString("v_job_code")

if Request.QueryString("Action") = "SAVE" then
	Response.Write "<HTML><body style='font-family:arial;BACKGROUND-COLOR: beige'><h3 align=center>"
	Response.Write "Updating...<br>Please wait"
	Response.Write "</h3></body></html>"
	if request.form("job_code")<>"" then
		connObj.execute("Update MOC_DEFICIENCIES set wls_job_code=" & request.form("job_code") & ", vessel_code='" & v_vessel_code & "' where deficiency_id=" & v_def_id)
	else
		connObj.execute("Update MOC_DEFICIENCIES set wls_job_code=null, vessel_code=null where deficiency_id=" & v_def_id)
	end if
%>
	<script language=vbscript>
	sub window_onload
		setTimeout "CloseWindow",1000
	end sub
	sub CloseWindow
		'window.returnValue="<%=request.form("job_list_id")%>"
		'self.close
	end sub
	</script>
<%

else

SQL = " Select job_code,vessel_code,job_description,status,action_planned,"
SQL = SQL &  " to_char(date_assigned,'dd-Mon-yy')date_assigned,to_char(end_date,'dd-Mon-yy')end_date"
SQL = SQL &  " from WLS_JOB_LIST"
SQL = SQL &  " where assignor_dept in('ASGMOC','ASGPTSTC')" 'MOC and PORT STATE observations
SQL = SQL &  " and vessel_code = '" & v_vessel_code & "'"
SQL = SQL &  " order by WLS_JOB_LIST.date_assigned desc"

set rs = connObj.execute(SQL)

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>NCR List - <%=v_vessel_name%></title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	tdMOC.innerText = window.dialogArguments
End Sub
sub Hilite(obj)
	obj.style.backgroundColor = "palegoldenrod"
end sub
sub RemoveHilite(obj)
	obj.style.backgroundColor = ""
end sub
dim jobcode
window.returnValue = "<%=v_job_code%>"
sub form1_onsubmit
	window.returnValue = jobcode
	setInterval "CloseWindow",500
	'self.close
end sub
sub CloseWindow
	if fra1.document.body.innerText<>"" then
		self.close
	end if
end sub
sub SetJobCode
	if window.event.srcElement.checked then
		jobcode = window.event.srcElement.value
	else
		jobcode = ""
	end if
end sub
sub ClearSelection
	on error resume next
	form1.job_code.checked = false
	for each obj in form1.job_code
		obj.checked=false
	next
end sub
-->
</SCRIPT>
<style>
BODY TABLE TD{font-size:11px;}
</style>
</HEAD>
<BODY class="bgcolorlogin" scroll=no leftmargin=5px rightmargin=5px>
<center>
<div style="text-align:center;width:100%;"><h4 style="padding:0;margin:0;">List of Jobs in Worklist</h4></div>
<span style="float:left">
<iframe id=fra1 name=fra1 style="display:none"></iframe>
<form name=form1 method=post action="ncr_list.asp?v_vessel_code=<%=v_vessel_code%>&v_def_id=<%=v_def_id%>&ACTION=SAVE" target=fra1>
<button onclick="ClearSelection">Clear Selection</button>&nbsp;&nbsp;
<input type=submit value="Save">
</span><br>
<hr style="visibility:hidden">
<table bgcolor=lightgrey cellspacing=1 width=100%>
  <tr>
    <td class=tableheader colspan=5>Observation in MOC database
  <tr>
    <td colspan=5 id=tdMOC>
</table>
<hr>
<div style="height:450px;overflow-y:auto;">
<table width=100%>    
  <tr>
    <td class=tableheader>Select
    <td class=tableheader>Date<br>created
    <td class=tableheader>Description
    <td class=tableheader>Status
    <td class=tableheader>Date<br>completed
  <%
  while not rs.eof%>
  <tr bgcolor="<%=toggleColor(sColor)%>"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td><input type=radio name="job_code" id="job_code" value="<%=rs("job_code")%>"
			onpropertychange=SetJobCode
			<%if cstr(rs("job_code"))=v_job_code then Response.Write " checked"%>>
    <td nowrap><%=rs("date_assigned")%>
    <td><%=rs("job_description")%>
    <td><%=rs("status")%>
    <td nowrap><%=rs("end_date")%>
  <%rs.movenext
  wend
  %>
</table>
</div>
</form>
</center>
</BODY>
</HTML>
<%end if%>