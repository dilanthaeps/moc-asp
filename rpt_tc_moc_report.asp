<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<%	'===========================================================================
	'	Template Name	:	MOC Status Report
	'	Template Path	:	rpt_StatusReport.asp
	'	Functionality	:	Current status report of MOC inspections
	'	Called By		:	ins_request_maint.asp
	'	Created By		:	Prashant Kumar
	'   Create Date		:	9 August, 2006
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
dim SQL,rs,rsDetails,lastVessel,sClassExp, sClassHilite
dim FLEET
FLEET = Request.QueryString("FLEET")
lastVessel = ""

set rs=Server.CreateObject("ADODB.Recordset")
with rs
	.CursorLocation=adUseClient
	.CursorType=adOpenDynamic
	.LockType=adLockReadonly
end with
set rsDetails=Server.CreateObject("ADODB.Recordset")
with rsDetails
	.CursorLocation=adUseClient
	.CursorType=adOpenDynamic
	.LockType=adLockReadonly
end with
SQL = " select mta.vessel_code,v.vessel_name,mtv.remarks,"
SQL = SQL &  " mta.time_charterer_id,mc.short_name tc_name,"
SQL = SQL &  " mta.moc_id,mm.short_name moc_name,decode(mta.mandatory,1,'<font color=maroon><b>*</b></font>','')mandatory"
SQL = SQL &  " from MOC_TC_MOC_ASGN mta,MOC_TC_VESSEL_ASGN mtv,"
SQL = SQL &  " wls_vw_vessels_new v,moc_master mm, moc_time_charterers mc"
SQL = SQL &  " where mtv.vessel_code=mta.vessel_code"
SQL = SQL &  " and mta.vessel_code=v.vessel_code"
SQL = SQL &  " and mta.time_charterer_id=mc.time_charterer_id"
SQL = SQL &  " and mta.moc_id=mm.moc_id"
if FLEET<>"" then
	SQL = SQL &  " and v.fleet_code = '" & FLEET & "'"
end if
SQL = SQL &  " order by v.vessel_name, mta.mandatory desc,moc_name"
rs.Open SQL,connObj


SQL = " Select ir.request_id,ir.vessel_code, v.vessel_name,moc_id, moc_fn_moc_short_name(moc_id)moc_short_name,"
SQL = SQL &  " ir.status, insp_status, to_char(inspection_date,'dd-Mon-yyyy')inspection_date, to_char(expiry_date,'dd-Mon-yyyy')expiry_date,"
SQL = SQL &  " basis_sire, basis_sire_moc"
SQL = SQL &  " from moc_inspection_requests ir, wls_vw_vessels_new v"
SQL = SQL &  " where ir.vessel_code = v.vessel_code"
if FLEET<>"" then
	SQL = SQL &  " and v.fleet_code = '" & FLEET & "'"
end if
SQL = SQL &  " and ir.status <> 'COMPLETED'"
SQL = SQL &  " order by v.vessel_name, moc_short_name, inspection_date"

rsDetails.open SQL,connObj



%>
<HTML>
<HEAD>
<title>MOC Inspections - Status Report</title>
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
	color:red;
	font-weight:bold;
}
.clsBasisMOC
{
	color:darkblue;
	font-size:10px;
	font-weight:bold;
}
</style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	cmbFleet.value = "<%=FLEET%>"
End Sub

Sub window_onbeforeprint
	divTopMenu.style.display = "none"
	cmbFleet.style.display = "none"
End Sub
Sub window_onafterprint
	divTopMenu.style.display = ""
	cmbFleet.style.display = ""
End Sub

sub cmbFleet_onchange
	window.location.href = "rpt_tc_moc_Report.asp?FLEET=" & cmbFleet.value
end sub
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
<table cellpadding=2 cellspacing=0>
  <tr>
    <td align=center><h3 style="margin-bottom:0">MOC Approval Status Report - T/C Ships
    <%if FLEET<>"" then%>
    - <%=FLEET%>
    <%end if%>
    </h3>
    
    <br>
      <%=dateTimeDisplay(now)%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <select name=cmbFleet class=menuHide>
        <option value="<%%>">All Fleets
        <option value="AFRAMAX">AFRAMAX
        <option value="FSO">FSO
        <option value="PRODUCT">PRODUCT
        <option value="SUEZMAX">SUEZMAX
        <option value="VLCC">VLCC
      </select>
</table>
<br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <caption class=num style="padding:0"><font size=-2 color=maroon><b>* mandatory</b></font></caption>
  <tr class=tableheader>
    <td style="font-size:12px">Vessel
    <td style="font-size:12px">MOC
    <td style="font-size:12px">Inspection Status
    <td style="font-size:12px">Inspected on
    <td style="font-size:12px">Expiry Date
<%
dim sColor

while not rs.eof
	rsDetails.filter="vessel_code='" & rs("vessel_code") & "' and moc_id='" & rs("moc_id") & "'"
	
	if lastVessel<>rs("vessel_name") then
      sColor=""
      lastVessel = rs("vessel_name")
%>
  <tr class="bgcolorlogin"><td colspan=5 style="height:15px">
  
  <tr bgcolor="<%=ToggleColor(sColor)%>">
    <td class=clsVessel><%=rs("vessel_name")%>
    <td colspan=4>Time chartered to - <b><%=rs("tc_name")%></b>
  <tr bgcolor="white">
    <td>
    <td colspan=4><%=rs("remarks")%>
	<%
	else
		if not rsDetails.EOF then
		while not rsDetails.eof
			if not IsNull(rsDetails("expiry_date")) then
				if cdate(rsDetails("expiry_date"))<now then
					sClassExp = " clsExpired"
				elseif cdate(rsDetails("expiry_date")) < (now+60) then
					sClassHilite = " clsHighlighted"
				end if
			end if
%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rsDetails("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td class=num><%=rs("mandatory")%>
    <td><%=rs("moc_name")%>
    <td class="<%=sClassExp%>">
    <%if sClassExp<>"" then%>
  EXPIRED
    <%else%>
  <%=rsDetails("insp_status")%>
      <%if rsDetails("basis_sire_moc")<>"" then%>
      / <span class=clsBasisMOC><%=rsDetails("basis_sire_moc")%></span>
      <%end if%>
    <%end if%>
    <td>&nbsp;<%=rsDetails("inspection_date")%>
    <td class="<%=sClassExp%> <%=sClassHilite%>">&nbsp;<%=rsDetails("expiry_date")%>
		<%	sClassExp = ""
			sClassHilite = ""
			rsDetails.movenext
		wend
		else%>
  <tr bgcolor="<%=ToggleColor(sColor)%>"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td class=num><%=rs("mandatory")%>
    <td><%=rs("moc_name")%>
    <td class="<%=sClassExp%>" colspan=3><font color=red><b>Not inspected yet</b></font>
		
	<%	end if
		rs.movenext
	end if
wend%>
</table>
</center>
<br>
</BODY>
</HTML>
