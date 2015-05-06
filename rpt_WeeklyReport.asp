<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Weekly Report
	'	Template Path	:	rpt_WeeklyReport.asp
	'	Functionality	:	Weekly report of MOC inspections
	'	Called By		:	ins_request_maint.asp
	'	Created By		:	Prashant Kumar
	'   Create Date		:	9 August, 2006
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
dim SQL,rsRequested,rsExpired,rsDueToExpire,lastVessel,lastMOC
dim sClassExp,sClassHilite

SQL = " Select ir.request_id, v.vessel_name, moc_fn_moc_short_name(moc_id)moc_short_name,"
SQL = SQL &  " nvl2(vessel_advised_date,'<img src=""images/tick_green.gif"">',null)vessel_advised,"
SQL = SQL &  " nvl2(agent_advised_date,'<img src=""images/tick_green.gif"">',null)agent_advised,"
SQL = SQL &  " nvl2(po_number,'<img src=""images/tick_green.gif"">',null)invoice_received,"
SQL = SQL &  " ir.inspection_port,ir.status, insp_status,"
SQL = SQL &  " to_char(inspection_date,'dd-Mon-yyyy')inspection_date,"
SQL = SQL &  " to_char(expiry_date,'dd-Mon-yyyy')expiry_date"
SQL = SQL &  " from moc_inspection_requests ir, wls_vw_vessels_new v"
SQL = SQL &  " where ir.vessel_code = v.vessel_code"
SQL = SQL &  " and upper(ir.status) <> 'COMPLETED'"
SQL = SQL &  " and upper(ir.insp_status) in ('REQUESTED INSPECTION','INSPECTION CONFIRMED')"
SQL = SQL &  " order by ir.inspection_date"

set rsRequested = connObj.execute(SQL)

SQL = " select * from("
SQL = SQL &  " Select ir.request_id, v.vessel_name, moc_fn_moc_short_name(ir.moc_id)moc_short_name,"
SQL = SQL &  " ir.inspection_port,insp_status,"
SQL = SQL &  " to_char(inspection_date,'dd-Mon-yyyy')inspection_date,"
SQL = SQL &  " to_char(expiry_date,'dd-Mon-yyyy')expiry_date,expiry_date expiry_date2,"
SQL = SQL &  " (case when trunc(expiry_date)<trunc(sysdate) then 'EXPIRED'"
SQL = SQL &  " 	  when trunc(expiry_date)<trunc(sysdate)+60 then 'DUE TO EXPIRE' end)status"
SQL = SQL &  " from moc_inspection_requests ir, moc_master mm, vpd_vessels_master v"
SQL = SQL &  " where ir.vessel_code = v.vessel_code"
SQL = SQL &  " and ir.moc_id = mm.moc_id"
SQL = SQL &  " and mm.entry_type in('MOC','TVEL')"
SQL = SQL &  " and upper(ir.status) <> 'COMPLETED'"
SQL = SQL &  " and (upper(ir.insp_status) = 'REQUEST TO BE SENT' or upper(ir.insp_status) like 'ACCEPTED%')"
SQL = SQL &  " and upper(v.status) = 'ACTIVE'"
'SQL = SQL &  " and ir.vessel_code='959'"
SQL = SQL &  " )"
SQL = SQL &  " where ((expiry_date is not null and status is not null)"
SQL = SQL &  " 	  or (expiry_date is null and status is null))"
SQL = SQL &  " order by vessel_name, moc_short_name,expiry_date2"

set rsExpired = connObj.execute(SQL)



%>
<HTML>
<HEAD>
<title>MOC Inspections - Weekly Report</title>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css">
<style>
TD
{
	 font-size:10px;
}
.clsVessel
{
	font-size:12px;
	font-weight:bold;
	
}
.clsExpired
{
	color:red;
	font-weight:bold;
}
.clsHighlighted
{
	color:blue;
	font-weight:bold;
}
</style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	'window.resizeTo 900,600
End Sub

Sub window_onbeforeprint
	divTopMenu.style.display="none"
End Sub
Sub window_onafterprint
	divTopMenu.style.display=""
End Sub
sub ShowInspection(v_ins_request_id)
	dim adWindow,winStats
	winStats="toolbar=no,location=no,directories=no,menubar=no,scrollbars=yes," & _
		"resizable=yes,status=yes,left=50,top=10,width=770,height=650"
	
	set adWindow=window.open("ins_request_entry.asp?v_ins_request_id=" & v_ins_request_id,"moc_request_entry",winStats)
	adWindow.focus
end sub
sub Hilite(obj)
	obj.style.backgroundColor = "palegoldenrod"
end sub
sub RemoveHilite(obj)
	obj.style.backgroundColor = ""
end sub
-->
</SCRIPT>
</HEAD>
<BODY class="bgcolorlogin">
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<center>
<table>
  <tr>
    <td align=center><h3 style="margin-bottom:0">MOC Weekly Inspection Forecast</h3>
    <br><%=dateTimeDisplay(now)%>
</table>
<br>
<h4 style="margin:0;">Requested Inspections</h4>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <tr class="tableheader">
    <td style="font-size:12px">Vessel
    <td style="font-size:12px">Date
    <td style="font-size:12px">MOC
    <td style="font-size:12px">Port
    <td style="font-size:12px">Status
    <td style="font-size:12px">V
    <td style="font-size:12px">A
    <td style="font-size:12px">PO
  <%dim sColor
while not rsRequested.eof
  if lastVessel<>rsRequested("vessel_name") then%>
  <tr class="bgcolorlogin"><td colspan=8 style="height:15px">
  
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rsRequested("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td class=clsVessel><%=rsRequested("vessel_name")%>
<%else%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rsRequested("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
	<td>
<%end if%>    
	<td>&nbsp;<%=rsRequested("inspection_date")%>
    <td>&nbsp;<%=rsRequested("moc_short_name")%>
    <td>&nbsp;<%=rsRequested("inspection_port")%>
    <td>&nbsp;<%=rsRequested("insp_status")%>
    <td><%=rsRequested("vessel_advised")%>
    <td><%=rsRequested("agent_advised")%>
    <td><%=rsRequested("invoice_received")%>
  <%
  lastVessel = rsRequested("vessel_name")
  rsRequested.movenext
wend%>
</table>
<br>
<br>
<h4 style="margin:0;">Inspections that have expired or are due to expire in the next two months</h4>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <tr class="tableheader">
    <td style="font-size:12px">Vessel
    <td style="font-size:12px">Date
    <td style="font-size:12px">MOC
    <td style="font-size:12px">Port
    <td style="font-size:12px">Status
    <td style="font-size:12px">Expiry Date
  <%
  lastVessel=""
  while not rsExpired.eof
	if lastMOC="" then lastMOC = rsExpired("moc_short_name")
  	sClassExp=""
  	sClassHilite=""
  	if rsExpired("status")="EXPIRED" then
		sClassExp = " clsExpired"
	elseif rsExpired("status")="DUE TO EXPIRE" then
		sClassHilite = " clsHighlighted"
	end if
	
	if lastVessel<>rsExpired("vessel_name") then%>
  <tr class="bgcolorlogin"><td colspan=6 style="height:15px">
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rsExpired("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td class=clsVessel><%=rsExpired("vessel_name")%>
	<%else%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rsExpired("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td>
	<%end if%>

    <td>&nbsp;<%=rsExpired("inspection_date")%>
    <td>&nbsp;<%=rsExpired("moc_short_name")%>
    <td>&nbsp;<%=rsExpired("inspection_port")%>
    <td class="<%=sClassExp%> <%=sClassHilite%>">&nbsp;<%=rsExpired("status")%>
    <td>&nbsp;<%=rsExpired("expiry_date")%>
  <%lastVessel=rsExpired("vessel_name")
	lastMOC = rsExpired("moc_short_name")
    rsExpired.movenext
  wend%>
</table>
</center>
</BODY>
</HTML>
