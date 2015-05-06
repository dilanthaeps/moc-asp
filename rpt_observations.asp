<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<!--#include file="common_procs.asp"-->
<%
dim rs,rsVessel,rsMOC,rsInspector,rsChapters,SQL,sColor,cnt,rsMOCType,rsActionCode,ACTCODE
dim SDATE1,SDATE2,FLEET,VID,MOC,INSPECTOR,STATUS,FORMAT,MOCType
dim VIQ,KEYWORD,CHAPTER,FREQ,RISKFACTOR

SDATE1 = Request.QueryString("SDATE1")
SDATE2 = Request.QueryString("SDATE2")
FLEET = Request.QueryString("FLEET")
VID = Request.QueryString("VID")
MOC = Request.QueryString("MOC")
INSPECTOR = Request.QueryString("INSPECTOR")
STATUS = Request.QueryString("STATUS")
FORMAT = Request.QueryString("FORMAT")
VIQ = Request.QueryString("VIQ")
KEYWORD = Request.QueryString("KEYWORD")
CHAPTER = Request.QueryString("CHAPTER")
FREQ = Request.QueryString("FREQ")
RISKFACTOR = Request.QueryString("RISKFACTOR")
ACTCODE = Request.QueryString("ACTCODE")
MOCType=Request.QueryString("MOCType")

if FORMAT="EXCEL" then
	Response.ContentType = "application/vnd.ms-excel"
end if

SQL = "Select vessel_code, vessel_name, fleet_code from wls_vw_vessels_new where vessel_code in (select distinct vessel_code from moc_inspection_requests) order by vessel_name"
set rsVessel = connObj.execute(SQL)

SQL = "Select moc_id,short_name from moc_master where moc_id in (select distinct moc_id from moc_inspection_requests) order by short_name"
set rsMOC = connObj.execute(SQL)

SQL="select distinct insp_type from moc_inspection_requests"
set rsMOCType=connObj.execute(SQL)

SQL = "Select inspector_id,short_name from moc_inspectors where inspector_id in (select distinct inspector_id from moc_inspection_requests) order by short_name"
set rsInspector = connObj.execute(SQL)

SQL = " Select chapter,substr(question_number,1,instr(question_number,'.',1)-1)chapter_number"
SQL = SQL &  " from MOC_VIQ_QUESTIONS"
SQL = SQL &  " group by chapter,substr(question_number,1,instr(question_number,'.',1)-1)"
set rsChapters = connObj.execute(SQL)

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

if SDATE1="" and SDATE2="" and FLEET="" and VID="" and MOC="" and INSPECTOR="" and STATUS="" AND VIQ="" and ACTCODE="" and MOCType="" then
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
if Request.QueryString("VIQ")<>"" then
    SQL = SQL & " and section='" & Request.QueryString("VIQ") & "'"
end if
SQL = SQL & " order by v.vessel_name,v.tech_manager,moc.short_name,ins.short_name,mir.inspection_date desc"
'Response.Write SQL
'Response.end
set rs = connObj.execute(SQL)
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
	v_form.cmbRiskFactor.value = "<%=RISKFACTOR%>"
	v_form.cmbActcode.value="<%=ACTCODE%>"
	v_form.cmbMOCType.value="<%=MOCType%>"
End Sub
sub RefreshPage
	dim sUrl
	sUrl = "rpt_observations.asp?"
	sUrl = sUrl & "VID=" & v_form.cmbVessel.value
	sUrl = sUrl & "&FLEET=" & v_form.cmbFleet.value
	sUrl = sUrl & "&MOC=" & v_form.cmbMOC.value
	sUrl = sUrl & "&INSPECTOR=" & v_form.cmbInspector.value
	sUrl = sUrl & "&SDATE1=" & v_form.v_insp_from_date.value
	sUrl = sUrl & "&SDATE2=" & v_form.v_insp_to_date.value
	sUrl = sUrl & "&STATUS=" & v_form.cmbStatus.value
	sUrl = sUrl & "&CHAPTER=" & v_form.cmbChapter.value
	sUrl = sUrl & "&RISKFACTOR=" & v_form.cmbRiskFactor.value
	sUrl = sUrl & "&VIQ=" & v_form.txtVIQ.value
	sUrl = sUrl & "&KEYWORD=" &  Escape(v_form.txtKeyword.value)
	sUrl = sUrl & "&ACTCODE=" & v_form.cmbActcode.value
	sUrl = sUrl & "&MOCType=" & v_form.cmbMOCType.value
	window.location.href = sUrl
end sub
sub ShowInspection(v_ins_request_id)
	dim adWindow,winStats
	winStats="toolbar=no,location=no,directories=no,menubar=no,scrollbars=yes," & _
		"resizable=yes,status=yes,left=50,top=10,width=770,height=650"	
	'set adWindow=window.open("ins_request_entry.asp?v_ins_request_id=" & v_ins_request_id,"moc_request_entry",winStats)
	set adWindow=window.open("ins_request_def_maint.asp?v_ins_request_id=" & v_ins_request_id,"moc_request_entry",winStats)
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
function outputExcel
	dim sUrl
	sUrl = "rpt_observations_excel.asp?"
	sUrl = sUrl & "VID=" & v_form.cmbVessel.value
	sUrl = sUrl & "&FLEET=" & v_form.cmbFleet.value
	sUrl = sUrl & "&MOC=" & v_form.cmbMOC.value
	sUrl = sUrl & "&INSPECTOR=" & v_form.cmbInspector.value
	sUrl = sUrl & "&SDATE1=" & v_form.v_insp_from_date.value
	sUrl = sUrl & "&SDATE2=" & v_form.v_insp_to_date.value
	sUrl = sUrl & "&STATUS=" & v_form.cmbStatus.value
	sUrl = sUrl & "&CHAPTER=" & v_form.cmbChapter.value
	sUrl = sUrl & "&RISKFACTOR=" & v_form.cmbRiskFactor.value
	sUrl = sUrl & "&VIQ=" & v_form.txtVIQ.value
	sUrl = sUrl & "&KEYWORD=" &  Escape(v_form.txtKeyword.value)
	sUrl = sUrl & "&ACTCODE=" & v_form.cmbActcode.value
	sUrl = sUrl & "&MOCType=" & v_form.cmbMOCType.value
	window.open sUrl,"listofobs"
end function
Sub cmbFleet_onchange
	lblVessel.innerHTML = cmbVessel(v_form.cmbFleet.selectedIndex).outerHTML
	lblVessel.children(0).style.display=""
End Sub
dim objTip,mx,my
Sub ShowTooltip(obj)
	set objTip = obj
	mx = window.event.clientX
	my = window.event.clientY
	if divTooltip.style.display="" then
		DisplayTooltip
	else
		setTimeout "DisplayTooltip",400,"vbscript"
	end if
End Sub

Sub HideTooltip
	divTooltip.style.display="none"
	divTooltip.innerHTML=""
	objTip = empty
End Sub
sub DisplayTooltip
	dim x,y

	if IsEmpty(objTip) then exit sub
	if objTip.children(0).innerText="" then exit sub
	x=mx + 20
	y=my + document.body.scrollTop
	'y=40 + objPortCallTip.offsetTop + objPortCallTip.parentElement.offsetTop
	

	with divTooltip
		.innerHTML = objTip.children(0).innerText
		.style.display=""
				
		if x+.offsetWidth+20 > document.body.offsetWidth then x=x-.offsetWidth
		if y+.offsetHeight+5 > document.body.offsetHeight + document.body.scrollTop then y=y-.offsetHeight
		.style.left = x
		.style.top = y
	end with
end sub
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
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='AFRAMAX'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='FSO'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='PRODUCT'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='SUEZMAX'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='VLCC'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='BULK'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='CONTAINER'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='PCTC'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='CHEMICAL'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<%rsVessel.filter="fleet_code='LPG'"%>
<select name=cmbVessel style="display:none">
  <option value="<%%>">--All vessels--</option>
<%while not rsVessel.eof%>
  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
<%rsVessel.movenext
wend%>
</select>
<FORM NAME="v_form" METHOD="post">
<table style="border:1px solid blue" border=0 cellpadding=2 cellspacing=0 width=100%>
  <caption><h3 style="margin-bottom:0">Major Oil Companies - List of observations
    <%if FLEET<>"" then%>
    - <%=FLEET%>
    <%end if%>
    </h3>
  </caption>
  <tr>
    <td>
      <table><tr><td>Date from<br>
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
		
    <td>Fleet<br>
      <select name=cmbFleet class=menuHide>
        <option value="<%%>">All Fleets
        <option value="AFRAMAX">AFRAMAX
        <option value="FSO">FSO
        <option value="PRODUCT">PRODUCT
        <option value="SUEZMAX">SUEZMAX
        <option value="VLCC">VLCC
        <option value="BULK">BULK
        <option value="CONTAINER">CONTAINER
        <option value="PCTC">PCTC
        <option value="CHEMICAL">CHEMICAL
        <option value="LPG">LPG
      </select>
    <td>Vessel<br>
      <span id=lblVessel>
      <select name=cmbVessel>
        <option value="<%%>">--Select vessel--</option>
      <%
      rsVessel.filter=0
      while not rsVessel.eof%>
        <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
      <%rsVessel.movenext
      wend%>
      </select>
      </span>
      <td>MOC Type<br>
      <select name=cmbMOCType style="width:140px">
        <option value="<%%>">-Select MOCType-</option>
      <%
      while not rsMOCType.eof%>
        <option value="<%=rsMOCType("insp_type")%>"><%=rsMOCType("insp_type")%>
      <%rsMOCType.movenext
      wend%>
      </select>
      
    <td>MOC<br>
      <select name=cmbMOC style="width:210px" class=menuHide>
        <option value="<%%>">--Select MOC--</option>
      <%
      while not rsMOC.eof%>
        <option value="<%=rsMOC("moc_id")%>"><%=rsMOC("short_name")%>
      <%rsMOC.movenext
      wend%>
      </select>   
    
    <td>Risk factor<br>
      <select name=cmbRiskFactor class=menuHide>
        <option value="<%%>">--Select risk factor--</option>
        <option value="GENERAL">General
        <option value="LOW">Low
        <option value="HIGH">High
      </select>
      </table>
  <tr>
    <td colspan=5>
    <table width=100%>
      <tr>
        <td>Inspector<br>
      <select name=cmbInspector class=menuHide style="width:200px">
        <option value="<%%>">--Select inspector--</option>
      <%
      while not rsInspector.eof%>
        <option value="<%=rsInspector("inspector_id")%>"><%=rsInspector("short_name")%>
      <%rsInspector.movenext
      wend%>
      </select>
      
      <td>VIQ Chapter<br>
		  <select name=cmbChapter class=menuHide>
		    <option value="<%%>">--All chapters--
		    <%while not rsChapters.eof%>
		    <option value="<%=rsChapters("chapter")%>"><%=rsChapters("chapter")%>
		    <%rsChapters.movenext
		    wend%>
		  </select>  
		<td><nobr>VIQ Number</nobr><br>
		  <input type=text name=txtVIQ value="<%=VIQ%>" size=5 style="width:70px">
		  		
		  <td align=center>Status<br>
		  <select name=cmbStatus>
			<option value="<%%>">--Select status--</option>
			<option value="Active">Active
			<option value="Pending">Pending
			<option value="Completed">Completed
		  </select>
		  <td>Action Code<br>
		    <select name=cmbActcode class=menuHide>
		      <option value="">-Select Action Code-</option>
		      <%while not rsActionCode.eof%>
		        <option value="<%=rsActionCode("code")%>"><%=rsActionCode("code")%>
		      <%rsActionCode.movenext
		        wend%>
		      </select>     
		   <td><nobr>Key-phrase</nobr><br>
		  <input type=text name=txtKeyWord value="<%=KEYWORD%>">
	  <tr>
		<td colspan=6 align=center>
		  <span style="color:maroon;font-weight:bold;float:right;visibility:hidden"><br>*  The VIQ Question No is based on SIRE VIQ 3rd Ed / 2005 format</span>
		  <span style="color:maroon;font-weight:bold;float:left"><br>*  The VIQ Question No is based on SIRE VIQ 3rd Ed / 2005 format</span>
		  <button id=cmdRefresh onclick="RefreshPage">Refresh</button>
	</table>
</table>
</form>
<div>
<span style="float:right">
  <a href="javascript:outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
  <a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
</span>
<span id=lblCount style="float:left"></span></div><br><br>
<table border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey width=100%>
  <tr class=tableheader>
    <td>Vessel
    <td>Inspecting Agency    
    <td>Insp Date
    <td>VIQ No.
    <td>Obs Details
    <td>Risk
    <td>Status
    <td>Action Code
    
  <%cnt=0
  if not rs.eof  then
  while not rs.eof%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" style="cursor:hand" onclick="ShowInspection(<%=rs("request_id")%>)"
	onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
    <td nowrap><%=rs("Vessel_name")%>
    <td><%=rs("moc_name")%>    
    <td nowrap><%=rs("inspection_date")%>
    <td onmousemove="ShowTooltip(me)" onmouseover="ShowTooltip(me)" onmouseout="HideTooltip()">
		<textarea id="qtext" style="display:none"><%=rs("question_text")%></textarea>
		<%=rs("section")%>
    <td onmousemove="ShowTooltip(me)" onmouseover="ShowTooltip(me)" onmouseout="HideTooltip()">
		<textarea id="reply" style="display:none"><%=rs("reply")%></textarea>
		<%=rs("deficiency")%>
    <td><%=rs("risk_factor")%>
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
<div id=divTooltip style="display:none;position:absolute;background-color:lightblue;font-size:10px;font-weight:normal;
	height:50px;width:350px;overflow:visible;text-align:left;padding:5px;border:1px solid midnightblue"></div>
</BODY>
</HTML>
<script>
lblCount.innerHTML = "<b>Number of observations:</b> <%=cnt%>"
</script>