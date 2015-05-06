<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<%
dim SQL,rs,sColor,MODE,cnt,border,VESSELFILTER
MODE = Request.QueryString("MODE")
border=0
if MODE="EXCEL" then
	Response.ContentType = "application/vnd.ms-excel"
	border=1
end if

SQL = "select mir.request_id,a.vessel_code,a.vessel_name,to_char(a.inspection_date,'dd-Mon-yyyy')inspection_date,mir.moc_id,mm.short_name"
SQL = SQL &  " from"
SQL = SQL &  " (select mir.vessel_code,v.vessel_name,max(inspection_date)inspection_date"
SQL = SQL &  " from moc_inspection_requests mir, wls_vw_vessels_new v"
SQL = SQL &  " where mir.vessel_code=v.vessel_code"
SQL = SQL &  " and is_sire='Y'"
'commeted by sankar based on Capt. Mishra request SQL = SQL &  " and mir.ocimf_report_number is not null"
'SQL = SQL &  " and upper(mir.status)='ACTIVE'"
'SQL = SQL &  " and mir.insp_type='MOC'"
SQL = SQL &  " and mir.insp_status in ('INSPECTED','ACCEPTED','FAILED','SIRE REPORT REPLIED','SIRE REPORT RECEIVED','TECHNICAL HOLD','INSPN PROCESS COMPLT')"
SQL = SQL &  " and v.FLEET_CODE IN ('AFRAMAX','PRODUCT','SUEZMAX','VLCC','CHEMICAL','LPG','CONTAINER','BULK','PCTC')"
'SQL = SQL &  " and mir.insp_status not like '%SIRE%'"
SQL = SQL &  " group by mir.vessel_code,v.vessel_name)A,"
SQL = SQL &  " moc_inspection_requests mir, moc_master mm"
SQL = SQL &  " where mir.vessel_code=a.vessel_code"
SQL = SQL &  " and mir.inspection_date=a.inspection_date and mir.is_sire='Y'"
SQL = SQL &  " and mir.moc_id=mm.moc_id"
'commeted by sankar based on Capt. Mishra request SQL = SQL &  " and mir.ocimf_report_number is not null"
'SQL = SQL &  " and upper(mir.status)='ACTIVE'"
'SQL = SQL &  " and mir.insp_type='MOC'"
'SQL = SQL &  " and mir.insp_status in('INSPECTED','ACCEPTED')"
'SQL = SQL &  " and mir.insp_status not like '%SIRE%'"
SQL = SQL &  " order by a.inspection_date DESC,vessel_name"
'response.write SQL
set rs=connObj.execute(SQL)
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<title>Last SIRE Inspection</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css"></link>
<%if MODE<>"EXCEL" then%>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload

End Sub
sub ShowInspection(v_ins_request_id)
	dim adWindow,winStats
	winStats="toolbar=no,location=no,directories=no,menubar=no,scrollbars=yes," & _
		"resizable=yes,status=yes,left=50,top=10,width=770,height=650"
	
	set adWindow=window.open("ins_request_entry.asp?v_ins_request_id=" & v_ins_request_id,"moc_request_entry",winStats)
	'set adWindow=window.open("ins_request_def_maint.asp?v_ins_request_id=" & v_ins_request_id,"moc_request_entry",winStats)
	adWindow.focus
end sub
sub Hilite(obj)
	obj.style.backgroundColor = "palegoldenrod"
end sub
sub RemoveHilite(obj)
	obj.style.backgroundColor = ""
end sub
function outputExcel()
	window.open "rpt_last_sire_inspected.asp?MODE=EXCEL","lastsireinspected"
end function
-->
</SCRIPT>
<%end if%>
</HEAD>
<%if MODE<>"EXCEL" then%>
<BODY class="bgcolorlogin" style="text-align:center">
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<div>
<span style="float:right">
  <a href="javascript:outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
  <a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
</span>
</div><br><br>
<%else%>
<body>
<%end if%>
<table border=<%=border%> cellspacing=1 cellpadding=2 bgcolor=lightgrey>
  <caption><h3 style="margin-bottom:0">Last SIRE Inspected</h3></caption>
  <caption style="font-size:smaller"><span id=lblCount style="float:left"></span>
  </caption>
  <thead class=tableheader>
    <th align=left>Vessel</td>
    <th align=left>MOC</td>
    <th align=left>Date of last<br>physical inspection</td>
  </thead>
  <%
  cnt=0
  while not rs.eof%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td align=left><%=rs("vessel_name")%></td>
    <td align=left><%=rs("short_name")%></td>
    <td align=left><%=rs("inspection_date")%></td>
  </tr>
  <%cnt=cnt+1
    rs.movenext
  wend%>
</table>
<br>
</BODY>
</HTML>
<script>
lblCount.innerHTML = "<b>Number of records:</b> <%=cnt%>"
</script>