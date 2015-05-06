<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css">
<style>
TD
{
	 font-size:10px;
}
</style>
</HEAD>
<%
dim SQL,rs,rsVessel,rsMOC,rsInspector,rsChapters,sColor
dim FLEET,VID,MOC,INSPECTOR,STATUS,CHAPTER,VIQ,KEYWORD
dim SDATE1,SDATE2,TOP30

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
TOP30 = Request.QueryString("TOP30")
Response.ContentType = "application/vnd.ms-excel"

SQL = "Select * from("
SQL = SQL & " select count(*)cnt,md.section,vq.question_text"
SQL = SQL & " from moc_inspection_requests mir, moc_deficiencies md, moc_master moc, moc_inspectors ins, wls_vw_vessels_new v, moc_viq_questions vq"
SQL = SQL & " where mir.request_id=md.request_id"
SQL = SQL & " and mir.moc_id=moc.moc_id"
SQL = SQL & " and mir.inspector_id=ins.inspector_id"
SQL = SQL & " and mir.vessel_code=v.vessel_code"
SQL = SQL & " and md.section=vq.question_number(+)"
SQL = SQL & " and md.section is not null"

if SDATE1<>"" then
	SQL = SQL & " and trunc(mir.inspection_date)>='" & SDATE1 & "'"
end if
if SDATE2<>"" then
	SQL = SQL & " and trunc(mir.inspection_date)<='" & SDATE2 & "'"
end if
if FLEET <>"" then
	SQL = SQL & " and v.FLEET_CODE='" & FLEET & "'"
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
if VIQ<>"" then
	SQL = SQL & " and md.section like '" & VIQ & "%'"
end if
if KEYWORD<>"" then
	SQL = SQL & " and upper(md.deficiency) like '%" & ucase(KEYWORD) & "%'"
end if
if CHAPTER<>"" then
	SQL = SQL & " and vq.chapter = '" & CHAPTER & "'"
end if

SQL = SQL &  " group by md.section,vq.question_text"
SQL = SQL &  " order by cnt desc)"
if TOP30="TRUE" then
	SQL = SQL & " where rownum<31"
end if
set rs=connObj.execute(SQL)
%>
<BODY>
<FORM NAME="v_form" METHOD="post">
<table width="100%"><tr><td>
	<table border=1 cellspacing=0 cellpadding=0  bgcolor=lightgrey>
	  <tr class=tableheader>
	    <td nowrap>VIQ No.
	    <td>Question
	    <td>Freq
	  </tr>
	  <%
	  while not rs.eof%>
	  <tr bgcolor="<%=ToggleColor(sColor)%>">
	    <td align=left><%=rs("section")%>
	    <td><%=rs("question_text")%>
	    <td class=num><%=rs("cnt")%>
	  </tr>
	  <%rs.movenext
	  wend%>
	</table></td></tr>
</table>
</BODY>
</HTML>

