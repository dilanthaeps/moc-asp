<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
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
dim SQL,rs,lastVessel,sClassExp, sClassHilite,sClassHilite1
dim FLEET
FLEET = Request.QueryString("FLEET")
lastVessel = ""

SQL = " Select ir.request_id, vv.vessel_name, moc_fn_moc_short_name(moc_id)moc_short_name,"
SQL = SQL &  " ir.status, insp_status,to_char(inspection_date,'dd-Mon-yyyy')inspection_date,to_char(expiry_date,'dd-Mon-yyyy')expiry_date,inspection_date insp_date, expiry_date exp_date,"
SQL = SQL &  " basis_sire, basis_sire_moc"
SQL = SQL &  " from moc_inspection_requests ir, vpd_vessels_master v, vessels vv"
SQL = SQL &  " where ir.vessel_code = v.vessel_code"
SQL = SQL &  " and v.vessel_code = vv.vessel_code"
if FLEET<>"" then
	SQL = SQL &  " and vv.tech_manager = '" & FLEET & "'"
end if
SQL = SQL &  " and ir.status <> 'COMPLETED'"
SQL = SQL &  " and v.status = 'ACTIVE'"
SQL = SQL &  " order by v.vessel_name, insp_date desc, moc_short_name "

'response.write sql
'response.end
set rs = connObj.execute(SQL)

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
.clsHighlighted1
{
	color:red;
	
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
	window.location.href = "rpt_statusReport.asp?FLEET=" & cmbFleet.value
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
    <td align=center><h3 style="margin-bottom:0">Major Oil Companies - Inspection Status Report
    <%if FLEET<>"" then%>
    - <%=FLEET%>
    <%end if%>
    </h3>
    
    <br>
      <%=dateTimeDisplay(now)%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <select name=cmbFleet class=menuHide>
	       <option value="<%%>">All Fleets</option>
	       <option value="AFRAMAX">AFRAMAX</option>
	       <option value="FSO">FSO</option>
	       <option value="PRODUCT">PRODUCT</option>
	       <option value="SUEZMAX">SUEZMAX</option>
	       <option value="VLCC">VLCC</option>
	       <option value="BULK">BULK</option>
		   <option value="CONTAINER">CONTAINER</option>
		   <option value="PCTC">PCTC</option>
		   <option value="CHEMICAL">CHEMICAL</option>
		   <option value="LPG">LPG</option>
      </select>
</table>
<br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <tr class=tableheader>
    <td style="font-size:12px">Vessel
    <td style="font-size:12px">MOC
    <td style="font-size:12px">Inspection Status
    <td style="font-size:12px">Inspected on
    <td style="font-size:12px">Expiry Date
<%
dim sColor,sClassHilite2
while not rs.eof   
	if not isnull(rs("expiry_date")) then
		if cdate(rs("expiry_date"))<now then
			sClassExp = " clsExpired"
			sClassHilite2="clsHighlighted"
		'elseif rs("expiry_date") < (now+60) then
			'sClassHilite = " clsHighlighted"
		end if	
	end if
	if cdate(rs("inspection_date"))<(now-120) and (trim(rs("insp_status"))="ACCEPTED BASED SIRE" or trim(rs("insp_status"))="ACCEPTED")then		
		sClassHilite1 = "clsHighlighted1"
	end if
	

if lastVessel<>rs("vessel_name") then 
  sColor=""
%>
  <tr class="bgcolorlogin"><td colspan=5 style="height:15px">
  
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td class=clsVessel><%=rs("vessel_name")%>
<%else%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td>
<%end if%>
    <td><%=rs("moc_short_name")%>
    <td class="<%=sClassExp%>">
  <%if sClassExp<>"" then%>
  EXPIRED
  <%else%>
  <%=rs("insp_status")%>
    <%if rs("basis_sire_moc")<>"" then%>
      / <span class=clsBasisMOC><%=rs("basis_sire_moc")%></span>
    <%end if%>
  <%end if%>
    <td class="<%=sClassHilite1%>" nowrap="nowrap">&nbsp;<%=rs("inspection_date")%></td>
    <td class="<%=sClassHilite2%>" nowrap="nowrap">&nbsp;<%=rs("expiry_date")%>
<%lastVessel = rs("vessel_name")
  sClassExp = ""
  sClassHilite = ""
  sClassHilite1=""
  sClassHilite2=""
  rs.movenext
wend%>
</table>
</center>
<br>
</BODY>
</HTML>
