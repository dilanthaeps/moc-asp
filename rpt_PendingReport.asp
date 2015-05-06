<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Pending Report
	'	Template Path	:	rpt_PendingReport.asp
	'	Functionality	:	Current status report of pending MOC inspections
	'	Called By		:	ins_request_maint.asp
	'	Created By		:	Prashant Kumar
	'   Create Date		:	9 August, 2006
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
dim SQL,rs,lastVessel,sClassExp, sClassHilite
lastVessel = ""

set rs=Server.CreateObject("ADODB.Recordset")
with rs
	.CursorLocation = adUseClient
	.CursorType = adOpenDynamic
	.LockType = adLockReadonly
end with
SQL = " Select ir.request_id, v.vessel_name, moc_fn_moc_short_name(moc_id)moc_short_name,"
SQL = SQL &  " inspection_port, insp_status, to_char(inspection_date,'dd-Mon-yyyy')inspection_date,"
SQL = SQL &  " to_char(ir.expiry_date,'dd-Mon-yyyy')expiry_date,"
SQL = SQL &  " sire_recd_date sire_recd_date2,inspection_date inspection_date2,"
SQL = SQL &  " to_char(sire_recd_date,'dd-Mon-yyyy')sire_recd_date,"
SQL = SQL &  " to_char(date_replied_to_SIRE,'dd-Mon-yyyy')date_replied_to_SIRE,"
SQL = SQL &  " date_replied_to_SIRE date_replied_to_SIRE2,"
SQL = SQL &  " to_char((sire_recd_date+13),'dd-Mon-yyyy') reply_due_date"
SQL = SQL &  " from moc_inspection_requests ir, vpd_vessels_master v"
SQL = SQL &  " where ir.vessel_code = v.vessel_code"
SQL = SQL &  " and ir.status <> 'COMPLETED'"
SQL = SQL &  " and v.status = 'ACTIVE'"
'SQL = SQL &  " order by v.vessel_name, moc_short_name"
SQL = SQL &  " order by ir.inspection_date"

rs.Open SQL,connObj

%>
<HTML>
<HEAD>
<title>MOC Inspection - Pending Report</title>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css">
<style>
TD
{
	 font-size:10px;
	 padding-left:5px;
	 padding-right:5px;
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
	color:red;
	font-weight:bold;
}
</style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	dim i
	for i=0 to tab2.rows(0).cells.length-1
		set td2 = tab2.rows(0).cells(i)
		set td3 = tab3.rows(0).cells(i)
		if td2.offsetWidth > td3.offsetWidth then
			td2.style.width = td2.offsetWidth - 4
			td3.style.width = td2.offsetWidth - 4
		else
			td2.style.width = td3.offsetWidth - 4
			td3.style.width = td3.offsetWidth - 4
		end if
		'MsgBox td2.style.width & " : " & td3.style.width
	next
	for i=tab2.rows(0).cells.length-1 to 0 step -1
		set td2 = tab2.rows(0).cells(i)
		set td3 = tab3.rows(0).cells(i)
		if td2.offsetWidth > td3.offsetWidth then
			td2.style.width = td2.offsetWidth - 4
			td3.style.width = td2.offsetWidth - 4
		else
			td2.style.width = td3.offsetWidth - 4
			td3.style.width = td3.offsetWidth - 4
		end if
		'MsgBox td2.style.width & " : " & td3.style.width
	next
	for i=0 to 0
		tab1.rows(0).cells(i).style.width = tab2.rows(0).cells(i).style.width
	next
	
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

<table>
  <tr>
    <td align=center><h3 style="margin-bottom:0">MOC Inspection - Pending Report</h3>
    <br><%=dateTimeDisplay(now)%>
</table>
<br>
<table id=tab1 border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <caption nowrap><h4 nowrap style="margin:0;" class=txt>Vessels inspected and SIRE not received</h4></caption>
  <tr class=tableheader>
    <td style="font-size:12px;">Vessel
    <td style="font-size:12px;">MOC
    <td style="font-size:12px;">Inspected<br>on
    <td style="font-size:12px;">Port
<%dim sColor
rs.filter="insp_status='INSPECTED'"
while not rs.eof%>
<tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
  <td nowrap><%=rs("vessel_name")%>
  <td><%=rs("moc_short_name")%>
  <td nowrap><%=rs("inspection_date")%>
  <td><%=rs("inspection_port")%>
<%rs.movenext
wend%>
</table>
<br>
<table id=tab2 border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
	<caption><h4 style="margin:0;" class=txt>Pending reply to inspection report</h4></caption>
  <tr class=tableheader>
    <td style="font-size:12px;">Vessel
    <td style="font-size:12px;">MOC
    <td style="font-size:12px;">Inspected<br>on
    <td style="font-size:12px;">Port
    <td style="font-size:12px;">Report<br>received
    <td style="font-size:12px;">Reply<br>due
<%
rs.filter="insp_status='REPORT RECEIVED' or insp_status='SIRE REPORT RECEIVED' or insp_status='PENDING BASED SIRE' or insp_status='TECHNICAL HOLD'"
rs.sort="sire_recd_date2"
while not rs.eof%>
<tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
  <td nowrap><%=rs("vessel_name")%>
  <td><%=rs("moc_short_name")%>
  <td nowrap><%=rs("inspection_date")%>
  <td><%=rs("inspection_port")%>
  <td nowrap><%=rs("sire_recd_date")%>
  <td nowrap><%=rs("reply_due_date")%>
<%rs.movenext
wend%>
</table>
<br>
<table id=tab3 border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <caption><h4 style="margin:0;" class=txt>Report replied, pending acceptance by MOC</h4></caption>
  <tr class=tableheader>
    <td style="font-size:12px;">Vessel
    <td style="font-size:12px;">MOC
    <td style="font-size:12px;">Inspected<br>on
    <td style="font-size:12px;">Port
    <td style="font-size:12px;">Report<br>received
    <td style="font-size:12px;">Report<br>replied
<%
rs.filter="insp_status='REPORT REPLIED' or insp_status='SIRE REPORT REPLIED'"
'rs.sort="inspection_date2"
rs.sort="date_replied_to_SIRE2"
while not rs.eof%>
<tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
  <td nowrap><%=rs("vessel_name")%>
  <td><%=rs("moc_short_name")%>
  <td nowrap><%=rs("inspection_date")%>
  <td><%=rs("inspection_port")%>
  <td nowrap><%=rs("sire_recd_date")%>
  <td nowrap><%=rs("date_replied_to_sire")%>
<%rs.movenext
wend%>
</table>
<br>
</BODY>
</HTML>
