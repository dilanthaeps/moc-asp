<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<%
dim SQL,rs,rsMOC,sColor,MODE,cnt,border
dim SDATE1,SDATE2,INSP_TYPE
dim lastGroup,GROUPBY,sDisp

MODE = Request.QueryString("MODE")
SDATE1=Request.QueryString("SDATE1")
SDATE2=Request.QueryString("SDATE2")
GROUPBY=Request.QueryString("GROUPBY")
INSP_TYPE=Request.QueryString("INSP_TYPE")

if GROUPBY="" then GROUPBY="fleet_code"
if INSP_TYPE="" then INSP_TYPE="MOC"
border=0
if MODE="EXCEL" then
	Response.ContentType = "application/vnd.ms-excel"
	border=1
end if

SQL = " select mq.question_number,mq.question_text,"
if GROUPBY="moc_name" then
	SQL = SQL &  " mm.short_name moc_name,"
end if
if GROUPBY="fleet_code" then
SQL = SQL &  " mir.vessel_code,v.vessel_name,"
end if
SQL = SQL &  " count(*)no_of_deficiencies"
SQL = SQL &  " from moc_inspection_requests mir,moc_deficiencies md,moc_viq_questions mq,moc_master mm,vessels v"
SQL = SQL &  " where mir.request_id=md.request_id"
SQL = SQL &  " and mir.moc_id=mm.moc_id"
SQL = SQL &  " and md.section=mq.question_number"
SQL = SQL &  " and mir.vessel_code=v.vessel_code"
'SQL = SQL &  " and mir.vessel_code='959'"
SQL = SQL &  " group by mq.question_number,"
if GROUPBY="moc_name" then
	SQL = SQL &  " mm.short_name,"
end if
if GROUPBY="fleet_code" then
	SQL = SQL &  " mir.vessel_code,v.vessel_name,"
end if
SQL = SQL &  " mq.question_text"
SQL = SQL &  " order by question_number,"
if GROUPBY="moc_name" then
	SQL = SQL &  " moc_name"
end if
if GROUPBY="fleet_code" then
	SQL = SQL &  " vessel_name"
end if

'Response.Write sql
'Response.End
set rs=connObj.execute(SQL)

if GROUPBY="fleet_code" then
	sDisp = "FLEET"
else
	sDisp = "MOC"
end if
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>MOC Inspections - Management Report</title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css"></link>
<style>
.clsFleet
{
	font-size:14px;
	font-weight:bold;
}
.num
{
	text-align:right;
}

</style>
</head>
<body class="bgcolorlogin" style="text-align:center">
<form name=v_form>
<table style="border:1px solid blue" border="0" cellpadding="2" cellspacing="0" width="">
  <caption><h3 style="margin-bottom:0">Overall Statistics grouped by <%=sDisp%></h3></caption>
  <tr>
    <td colspan="5" style="text-align:center">
    <table cellpadding=5>
      <tr>
        <td>Group by<br>
        <select name="groupby">
          <option value="fleet_code">FLEET
          <option value="moc_name">MOC
        </select>
        <td>Type<br>
        <select name="insp_type">
          <option value="MOC">MOC
          <option value="PSC">PSC
          <option value="TMNL">Terminal
        </select>
        <td nowrap>Date from<br>
          <nobr>
          <input TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<%=SDATE1%>" SIZE="12" onblur="vbscript:valid_date v_form.v_insp_from_date,'Inspection Date From','v_form'">
				<a HREF="javascript:show_calendar('v_form.v_insp_from_date',v_form.v_insp_from_date.value);">
				<img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0"></a>
		  </nobr>
		<td nowrap>Date to<br>
		  <input TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_to_date" VALUE="<%=SDATE2%>" SIZE="12" onblur="vbscript:valid_date v_form.v_insp_to_date,'Inspection Date From','v_form'">
				<a HREF="javascript:show_calendar('v_form.v_insp_to_date',v_form.v_insp_to_date.value);">
				<img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0"></a>
	  <tr>
		<td colspan="4" align="center">
		  <button id="cmdRefresh" onclick="RefreshPage" class="hideonprint">Refresh</button>
	</table>
  </tr>
</table>
</form>
<table width=500px>
  <%
  lastGroup=""
  dim i
  i=0
  while not rs.eof
    if lastGroup<>cstr(rs(GROUPBY).value) then
    i=i+1%>
  <tr id="tr<%=i%>" bgcolor="<%=toggleColor(sColor)%>">
    <td colspan="3" class="clsFleet">
    <img src="images/collapsed.gif" onclick="ToggleDetails">
    <%=rs(GROUPBY)%>
    <td width=110px><span id="cnt<%=i%>" style="font-size:10px"></span>
  <tr id="tr<%=i%>" style="display:none" bgcolor="<%=sColor%>">
    <td class="tableheader">Vessel
    <td class="tableheader num">No of inspections
    <td class="tableheader num">No of observations
    <td class="tableheader num">Avg per inspection
  </tr>
  <%end if%>
  <tr id="tr<%=i%>" style="display:none" bgcolor="white">
    <td><%=rs("vessel_name")%>
    <td class="num"><%=rs("no_of_inspections")%>
    <td class="num"><%=rs("no_of_deficiencies")%>
    <td class="num"><%=rs("avg_per_inspection")%>
  </tr>
  <%
    lastGroup = cstr(rs(GROUPBY))
    rs.movenext
  wend
  %>
</table>
</body>
</html>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim imgExpanded,imgCollapsed
Sub window_onload
	dim s
	v_form.groupby.value = "<%=GROUPBY%>"
	v_form.insp_type.value = "<%=INSP_TYPE%>"
	for i=1 to <%=i%>
		s = "tr" & i
		set oColl = document.getElementsByName(s)
		s = "cnt" & i
		set obj = document.getElementById(s)
		obj.innertext = oColl.length-2 & " vessels"
	next
	
	set imgExpanded = document.createElement("IMG")
	set imgCollapsed = document.createElement("IMG")
	imgExpanded.src = "images/expanded.gif"
	imgCollapsed.src = "images/collapsed.gif"
End Sub

sub RefreshPage
	dim sURL
	sURL = "rpt_management.asp?GROUPBY=" & v_form.groupby.value & "&SDATE1=" & v_form.v_insp_from_date.value & "&SDATE2=" & v_form.v_insp_to_date.value & "&INSP_TYPE=" & v_form.insp_type.value
	window.location.href = sURL
end sub

sub ToggleDetails()
	dim obj,objImg,sOp
	set obj = window.event.srcElement
	do
		set obj = obj.parentElement
	loop while obj.tagname<>"TR"
	set objImg = obj.cells(0).children(0)
	if instr(1,ucase(objImg.src),"COLLAPSED")>0 then
		objImg.src = imgExpanded.src
		sOp = ""
	else
		objImg.src = imgCollapsed.src
		sOp = "none"
	end if
	HideGroup obj,sOp
end sub

sub HideGroup(objTr,op)
	set oColl = document.getElementsByName(objTr.id)
	for i=1 to oColl.length-1
		oColl(i).style.display = op
	next
end sub
Sub window_onbeforeprint
	v_form.cmdRefresh.style.display="none"
End Sub

Sub window_onafterprint
	v_form.cmdRefresh.style.display=""
End Sub

</script>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>