<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
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
'SQL = SQL &  " having count(*)>5"
SQL = SQL &  " order by cnt desc)"
if TOP30="TRUE" then
	SQL = SQL & " where rownum<61"
end if

'Response.Write sql
set rs = connObj.execute(SQL)

SQL = "Select vessel_code, vessel_name, fleet_code from wls_vw_vessels_new where vessel_code in (select distinct vessel_code from moc_inspection_requests) order by vessel_name"
set rsVessel = connObj.execute(SQL)

SQL = "Select moc_id,short_name from moc_master where moc_id in (select distinct moc_id from moc_inspection_requests) order by short_name"
set rsMOC = connObj.execute(SQL)

SQL = "Select inspector_id,short_name from moc_inspectors where inspector_id in (select distinct inspector_id from moc_inspection_requests) order by short_name"
set rsInspector = connObj.execute(SQL)

SQL = " Select chapter,substr(question_number,1,instr(question_number,'.',1)-1)chapter_number"
SQL = SQL &  " from MOC_VIQ_QUESTIONS"
SQL = SQL &  " group by chapter,substr(question_number,1,instr(question_number,'.',1)-1)"
set rsChapters = connObj.execute(SQL)
%>
<HTML>
<HEAD>
<title>Observation frequency</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css">
<style>
TD
{
	 font-size:10px;
}
</style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	cmbFleet_onchange	
	v_form.cmbFleet.value = "<%=FLEET%>"
	v_form.cmbVessel.value = "<%=VID%>"
	v_form.cmbMOC.value = "<%=MOC%>"
	v_form.cmbInspector.value = "<%=INSPECTOR%>"
	v_form.cmbStatus.value = "<%=STATUS%>"
	v_form.cmbChapter.value = "<%=CHAPTER%>"
	
End Sub
sub RefreshPage(top60)
	dim sUrl
	sUrl = "rpt_recurring_observations.asp?"
	sUrl = sUrl & "VID=" & v_form.cmbVessel.value
	sUrl = sUrl & "&FLEET=" & v_form.cmbFleet.value
	sUrl = sUrl & "&MOC=" & v_form.cmbMOC.value
	sUrl = sUrl & "&INSPECTOR=" & v_form.cmbInspector.value
	sUrl = sUrl & "&SDATE1=" & v_form.v_insp_from_date.value
	sUrl = sUrl & "&SDATE2=" & v_form.v_insp_to_date.value
	sUrl = sUrl & "&STATUS=" & v_form.cmbStatus.value
	sUrl = sUrl & "&CHAPTER=" & v_form.cmbChapter.value
	sUrl = sUrl & "&VIQ=" & v_form.txtVIQ.value
	sUrl = sUrl & "&KEYWORD=" &  Escape(v_form.txtKeyword.value)
	if top60 then
		sUrl = sUrl & "&TOP60=TRUE"				
	end if	
	window.location.href = sUrl
end sub
sub ShowObservationList(viq)
	dim adWindow,winStats
	winStats="toolbar=no,location=no,directories=no,menubar=no,scrollbars=yes," & _
		"resizable=yes,status=yes,left=50,top=10,width=770,height=650"	
	dim sUrl
	sUrl = "rpt_observations.asp?FREQ=TRUE"
	sUrl = sUrl & "&VID=<%=VID%>"
	sUrl = sUrl & "&FLEET=<%=FLEET%>"
	sUrl = sUrl & "&MOC=<%=MOC%>"
	sUrl = sUrl & "&INSPECTOR=<%=INSPECTOR%>"
	sUrl = sUrl & "&SDATE1=<%=SDATE1%>"
	sUrl = sUrl & "&SDATE2=<%=SDATE2%>"
	sUrl = sUrl & "&STATUS=<%=STATUS%>"
	sUrl = sUrl & "&CHAPTER=<%=CHAPTER%>"
	sUrl = sUrl & "&VIQ=" & viq
	sUrl = sUrl & "&KEYWORD=<%=KEYWORD%>"
	set adWindow=window.open(sUrl,"moc_request_entry",winStats)
	adWindow.focus
end sub
sub Hilite(obj)
	obj.style.backgroundColor = "palegoldenrod"
end sub
sub RemoveHilite(obj)
	obj.style.backgroundColor = ""
end sub
Sub window_onbeforeprint
	v_form.cmdRefresh.style.display="none"
	divTopMenu.style.display="none"
End Sub
Sub window_onafterprint
	v_form.cmdRefresh.style.display=""
	divTopMenu.style.display=""
End Sub
function outputExcel()
	dim sUrl	
	sUrl = "rpt_recurring_observations_excel.asp?TOP30=<%=TOP30%>"
	sUrl = sUrl & "&VID=" & v_form.cmbVessel.value
	sUrl = sUrl & "&FLEET=" & v_form.cmbFleet.value
	sUrl = sUrl & "&MOC=" & v_form.cmbMOC.value
	sUrl = sUrl & "&INSPECTOR=" & v_form.cmbInspector.value
	sUrl = sUrl & "&SDATE1=" & v_form.v_insp_from_date.value
	sUrl = sUrl & "&SDATE2=" & v_form.v_insp_to_date.value
	sUrl = sUrl & "&STATUS=" & v_form.cmbStatus.value
	sUrl = sUrl & "&CHAPTER=" & v_form.cmbChapter.value
	sUrl = sUrl & "&VIQ=" & v_form.txtVIQ.value
	sUrl = sUrl & "&KEYWORD=" &  Escape(v_form.txtKeyword.value)	
	window.open sUrl,"freqofobs"
end function
Sub cmbFleet_onchange
	lblVessel.innerHTML = cmbVessel(v_form.cmbFleet.selectedIndex).outerHTML
	lblVessel.children(0).style.display=""
End Sub

-->
</SCRIPT>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>
</HEAD>
<BODY class="bgcolorlogin">
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<center>
<%rsVessel.filter=0%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='AFRAMAX'"%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='FSO'"%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='PRODUCT'"%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='SUEZMAX'"%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='VLCC'"%>
<select name=cmbVessel style="display:none" class=menuHide>
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<FORM NAME="v_form" METHOD="post">
<input type=hidden name="top30" id="top30">
<table style="border:1px solid blue" border=0 cellpadding=2 cellspacing=0 width=100%>
  <caption><h3 style="margin-bottom:0">Major Oil Companies - Recurring observations
    <%if FLEET<>"" then%>
    - <%=FLEET%>
    <%end if%>
    </h3>
  </caption>
  <tr>
    <td>Fleet<br>
      <select name=cmbFleet class=menuHide>
        <option value="<%%>">All Fleets
        <option value="AFRAMAX">AFRAMAX
        <option value="FSO">FSO
        <option value="PRODUCT">PRODUCT
        <option value="SUEZMAX">SUEZMAX
        <option value="VLCC">VLCC
      </select>
    <td>Vessel<br>
      <span id=lblVessel>
      <select name=cmbVessel class=menuHide>
        <option value="<%%>">--Select vessel--</option>
      <%
      rsVessel.filter=0
      while not rsVessel.eof%>
        <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
      <%rsVessel.movenext
      wend%>
      </select>
      </span>
    <td>MOC<br>
      <select name=cmbMOC style="width:260px">
        <option value="<%%>">--Select MOC--</option>
      <%
      while not rsMOC.eof%>
        <option value="<%=rsMOC("moc_id")%>"><%=rsMOC("short_name")%>
      <%rsMOC.movenext
      wend%>
      </select>
    <td>Inspector<br>
      <select name=cmbInspector class=menuHide style="width:200px">
        <option value="<%%>">--Select inspector--</option>
      <%
      while not rsInspector.eof%>
        <option value="<%=rsInspector("inspector_id")%>"><%=rsInspector("short_name")%>
      <%rsInspector.movenext
      wend%>
      </select>
  <tr>
    <td colspan=4>
    <table width=100%>
      <tr>
        <td nowrap>Date from<br>
          <nobr>
          <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<%=SDATE1%>" SIZE="12"
				onblur="vbscript:valid_date v_form.v_insp_from_date,'Inspection Date From','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_insp_from_date',v_form.v_insp_from_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		  </nobr>
		<td nowrap>Date to<br>
		  <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_to_date" VALUE="<%=SDATE2%>" SIZE="12"
				onblur="vbscript:valid_date v_form.v_insp_to_date,'Inspection Date From','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_insp_to_date',v_form.v_insp_to_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		<td>Status<br>
		  <select name=cmbStatus>
			<option value="<%%>">--Select status--</option>
			<option value="Active">Active
			<option value="Completed">Completed
		  </select>
		<td><nobr>VIQ Chapter</nobr><br>
		  <select name=cmbChapter>
		    <option value="<%%>">--All chapters--
		    <%while not rsChapters.eof%>
		    <option value="<%=rsChapters("chapter")%>"><%=rsChapters("chapter")%>
		    <%rsChapters.movenext
		    wend%>
		  </select>
		<td><nobr>VIQ Number</nobr><br>
		  <input type=text name=txtVIQ value="<%=VIQ%>" size=5>
		<td><nobr>Key-phrase</nobr><br>
		  <input type=text name=txtKeyWord value="<%=KEYWORD%>">
	  <tr>
		<td colspan=6 align=center>
		  <span style="width:400px;color:maroon;font-weight:bold;float:left;text-align:left"><br>*  The VIQ Question No is based on SIRE VIQ 3rd Ed / 2005 format</span>
		  <span style="width:400px;float:right;cursor:hand;color:blue;font-weight:bold;text-align:right" onclick="RefreshPage(true)"><br>Show top 60 only</span>
		  <button id=cmdRefresh onclick="RefreshPage(false)">Refresh</button>
	</table>
</table>
</form>
<div>
<span width=100% style="float:right">
  <a href="javascript:outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
  <a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
</span></div>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey width=100%>
  <tr class=tableheader>
    <td nowrap>VIQ No.
    <td>Question
    <td>Freq
  </tr>
  <%
  while not rs.eof%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowObservationList '<%=rs("section")%>'"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td><%=rs("section")%>
    <td><%=rs("question_text")%>
    <td class=num><%=rs("cnt")%>
  </tr>
  <%rs.movenext
  wend%>
</table>
</BODY>
</HTML>