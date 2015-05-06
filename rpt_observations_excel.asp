<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<%
dim rs,rsVessel,rsMOC,rsInspector,SQL,sColor,cnt,ACTCODE,MOCType
dim SDATE1,SDATE2,FLEET,VID,MOC,INSPECTOR,STATUS,rsActionCode,rsMOCType
dim VIQ,KEYWORD,CHAPTER,FREQ,RISKFACTOR

SDATE1 = Request.QueryString("SDATE1")
SDATE2 = Request.QueryString("SDATE2")
FLEET = Request.QueryString("FLEET")
VID = Request.QueryString("VID")
MOC = Request.QueryString("MOC")
INSPECTOR = Request.QueryString("INSPECTOR")
STATUS = Request.QueryString("STATUS")

VIQ = Request.QueryString("VIQ")
KEYWORD = Request.QueryString("KEYWORD")
CHAPTER = Request.QueryString("CHAPTER")
FREQ = Request.QueryString("FREQ")
RISKFACTOR = Request.QueryString("RISKFACTOR")
ACTCODE = Request.QueryString("ACTCODE")
MOCType=Request.QueryString("MOCType")


Response.ContentType = "application/vnd.ms-excel"

SQL = "Select vessel_code, vessel_name, fleet_code from wls_vw_vessels_new where vessel_code ='" & VID & "'"
set rsVessel = connObj.execute(SQL)

SQL = "Select moc_id,short_name from moc_master where moc_id = '" & MOC & "'"
set rsMOC = connObj.execute(SQL)

SQL="select distinct insp_type from moc_inspection_requests where insp_type = '" & MOCType & "'"
set rsMOCType=connObj.execute(SQL)

SQL = "Select inspector_id,short_name from moc_inspectors where inspector_id = '" & INSPECTOR & "'"
set rsInspector = connObj.execute(SQL)

SQL = "select distinct code from moc_deficiency_action_codes "
SQL =SQL & " where  trim(code_type)='Deficiency Action Code' order by code"
set rsActionCode=connObj.execute(SQL)

SQL = "Select mir.request_id,v.vessel_name,moc.short_name moc_name,ins.short_name inspector_name,insp_type,action_code,"
SQL = SQL & " to_char(mir.inspection_date,'dd-Mon-yyyy')inspection_date,md.section,md.deficiency,md.status,md.risk_factor,"
SQL = SQL & " ('<u><b>Question:</b></u><br>' || vq.question_text)question_text,"
SQL = SQL & " ('<u><b>Reply:</b></u><br>' || md.reply)reply"
SQL = SQL & " from moc_inspection_requests mir, moc_deficiencies md, moc_master moc, moc_inspectors ins, vessels v, moc_viq_questions vq"
SQL = SQL & " where mir.request_id=md.request_id"
SQL = SQL & " and mir.moc_id=moc.moc_id"
SQL = SQL & " and mir.inspector_id=ins.inspector_id(+)"
SQL = SQL & " and mir.vessel_code=v.vessel_code"
'SQL = SQL & " and nvl(md.section,'MISSING')=nvl(vq.question_number,'MISSING')"
SQL = SQL & " and md.section=vq.question_number(+)"

if SDATE1="" and SDATE2="" and FLEET="" and VID="" and MOC="" and INSPECTOR="" and STATUS="" AND VIQ="" then
	SDATE1 = FormatDateTimeValues(now-180,2)
	SDATE2 = FormatDateTimeValues(now,2)
end if
if SDATE1<>"" then
	SQL = SQL & " and trunc(mir.inspection_date)>='" & SDATE1 & "'"
end if
if SDATE2<>"" then
	SQL = SQL & " and trunc(mir.inspection_date)<='" & SDATE2 & "'"
end if
if FLEET <>"" then
	SQL = SQL & " and v.tech_manager='" & FLEET & "'"
end if
if VID<>"" then
	SQL = SQL & " and mir.vessel_code ='" & VID & "'"
end if
if MOC<>"" then
	SQL = SQL & " and mir.moc_id ='" & MOC & "'"
end if
if INSPECTOR<>"" then
	SQL = SQL & " and mir.inspector_id ='" & INSPECTOR & "'"
end if
if STATUS<>"" then
	SQL = SQL & " and md.status ='" & STATUS & "'"
end if
if VIQ<>"" and FREQ="TRUE" then
	SQL = SQL & " and md.section = '" & VIQ & "'"
elseif VIQ<>"" then
	SQL = SQL & " and md.section like '" & VIQ & "%'"
end if
if KEYWORD<>"" then
	SQL = SQL & " and upper(md.deficiency) like '%" & ucase(KEYWORD) & "%'"
end if
if CHAPTER<>"" then
	SQL = SQL & " and vq.chapter = '" & CHAPTER & "'"
end if
if RISKFACTOR<>"" then
	SQL = SQL & " and md.risk_factor='" & RISKFACTOR & "'"
end if
if ACTCODE<>"" then
	SQL = SQL & " and action_code='" & ACTCODE & "'"
end if
if MOCType<>"" then
	SQL = SQL & " and insp_type='" & MOCType & "'"
end if

SQL = SQL & " order by v.tech_manager,v.vessel_name,moc.short_name,ins.short_name,mir.inspection_date desc"

set rs = Server.CreateObject("adodb.recordset")
with rs
	.CursorLocation = adUseClient
	.CursorType = adOpenDynamic
	.LockType = adLockReadonly
end with
rs.Open SQL,connObj

if STATUS="" then STATUS = "ALL"
%>
<HTML>
<HEAD>
<title>List of observations</title>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css">
<style>
TD
{
	 font-size:10px;
}
</style>
</HEAD>
<BODY>
<center>
<table width=100%>
  <tr>
    <td align=center colspan=4><h3 style="margin-bottom:0">Major Oil Companies - List of observations</h3>
    <tr>
    <td align=left colspan=4><b>Date from</b>: <%=SDATE1%>
    <tr>
    <td align=left colspan=4><b>Date to</b>: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=SDATE2%>
    <%if FLEET<>"" then%>
  <tr>
    <td align=left colspan=4><b>Fleet</b>: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=FLEET%>
    <%end if%>
    <%if not rsVessel.eof then%>
  <tr>
    <td align=left colspan=4><b>Vessel</b>: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsVessel("vessel_name")%>
    <%end if%>
    <%if not rsMOCType.eof then%>
  <tr>
    <td align=left colspan=4><b>MOCType</b>: &nbsp;&nbsp;<%=rsMOCType("insp_type")%>
    <%end if%>
    <%if not rsMOC.eof then%>
  <tr>
    <td align=left colspan=4><b>MOC</b>: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsMOC("short_name")%>
    <%end if%>
    <%if not rsInspector.eof then%>
  <tr>
    <td align=left colspan=4><b>Inspector</b>: &nbsp;<%=rsInspector("short_name")%>
    <%end if%>
  
  <tr>
    <td align="left" colspan=4><b>Status</b>: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=STATUS%>
  <tr><td>
  <tr><td align="left" colspan=4><b>Number of observations</b>: <%=rs.RecordCount%>
</table>
<table border=1 cellspacing=0 cellpadding=0 width=100%>
  <tr class=tableheader>
    <td>Vessel
    <td>Inspecting<br>Agency
    <td>Inspector
    <td>Insp Date
    <td>VIQ No.
    <td>Obs Details
    <td>Status
    <td>Action Code
  <%cnt=0
  if not rs.eof  then
  while not rs.eof%>
  <tr bgcolor="<%=ToggleColor(sColor)%>">
    <td nowrap><%=rs("Vessel_name")%>
    <td><%=rs("moc_name")%>
    <td><%=rs("inspector_name")%>
    <td nowrap><%=rs("inspection_date")%>
    <td><%=rs("section")%>
    <td><%=rs("deficiency")%>
    <!--<td><%=rs("reply")%>-->
    <td><%=rs("status")%>
    <td><%=rs("action_code")%>
  <%cnt=cnt+1
    rs.movenext
  wend
  else%>
  <tr><td>No records found
  <%
  end if%>
</table>
<br>
</BODY>
</HTML>