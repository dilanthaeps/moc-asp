<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Vessel TC Assignment Entry/Edit/Listing
	'	Template Path	:	.../ves_tc_asgn.asp
	'	Functionality	:	To view/edit the MOC Vessel TC Assignment
	'	stored_proc		:	N/A
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = true
%>

<html>
<head>
<title>MOC - Vessel to Time Charterer Assignment</title>
</head>

<%
dim VID,TCID,TCREMARKS
dim SQL, rsVessel, rsMOC, rsTC, rsTCSel, rsMOCSel
dim v_mode,v_header
v_mode = "edit"  
v_header="Update MOC - Vessel Time Charterer Assignment"

VID = Request.QueryString("VID")
TCID = Request.QueryString("TCID")

if VID="" then
	VID="959"
	'Response.Write "Missing parameters"
	'Response.End
end if

SQL = "Select time_charterer_id,remarks"
SQL = SQL & " from MOC_TC_VESSEL_ASGN"
SQL = SQL & " where vessel_code='" & VID & "'"
set rsTCSel = connObj.execute(SQL)

if not rsTCSel.eof then
	TCID = rsTCSel(0)
	TCREMARKS = rsTCSel("remarks")
else
	TCID = "-9999"
end if

SQL = "Select vessel_code,vessel_name"
SQL = SQL & " from WLS_VW_VESSELS_NEW"
SQL = SQL & " order by vessel_name"
set rsVessel = connObj.execute(SQL)

SQL = "Select moc_id,short_name"
SQL = SQL & " from moc_master"
SQL = SQL & " where moc_id not in(Select moc_id from MOC_TC_MOC_ASGN where vessel_code='" & VID & "' and time_charterer_id=" & TCID & ")"
SQL = SQL & " and entry_type='MOC'"
SQL = SQL & " order by short_name"
set rsMOC = connObj.execute(SQL)

SQL = "Select time_charterer_id,short_name"
SQL = SQL & " from moc_time_charterers"
'SQL = SQL & " where time_charterer_id not in(Select time_charterer_id from MOC_TC_MOC_ASGN where time_charterer_id=" & TCID & ")"
SQL = SQL & " order by short_name"
set rsTC = connObj.execute(SQL)

SQL = "Select mta.moc_id,mm.short_name,mta.mandatory"
SQL = SQL & " from MOC_TC_MOC_ASGN mta,moc_master mm"
SQL = SQL & " where mta.moc_id=mm.moc_id"
SQL = SQL & " and mta.vessel_code='" & VID & "'"
SQL = SQL & " and mta.time_charterer_id=" & TCID
SQL = SQL & " order by mm.short_name"
set rsMOCSel = connObj.execute(SQL)

if TCID = "-9999" then TCID=""
%>
	
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css"></link>
<TITLE>Tanker Pacific - TC - MOC Assignment</TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
	form1.cmbVessel.value = "<%=VID%>"
	form1.cmbTC.value = "<%=TCID%>"
	form1.cmbVessel.focus
	
	setInterval "UpdateCharCount",500
End Sub
sub UpdateCharCount
	lblCount.innerText = len(form1.txtRemarks.value)
end sub
dim iTimer
iTimer=0
sub cmbVessel_onchange
	vesselChanged
end sub
Sub vesselChanged
	clearTimeout(iTimer)
	iTimer = setTimeout("ReloadWindow",1000)
end sub
Sub ReloadWindow
	iTimer=0
	window.location.href = "ves_tc_asgn.asp?VID=" & form1.cmbVessel.value
End Sub

function form1_onsubmit
	dim sMand,sOpt,i
	if form1.cmbTC.value<>"" and form1.lstMOCSelMand.options.length=0 then
		MsgBox "Please select MOC names for the selected time charterer",64
		form1_onsubmit = false
		exit function
	end if
	if len(form1.txtRemarks.value)>500 then
		MsgBox "Please restrict your remarks to 500 chars",64
		form1.txtRemarks.focus
		form1.txtRemarks.select
		form1_onsubmit = false
		exit function
	end if
	'mandatory MOCs
	for i=0 to form1.lstMOCSelMand.options.length-1
		sMand = sMand & form1.lstMOCSelMand.options(i).value & ","
	next
	IF sMand<>"" THEN
		sMand = left(sMand,len(sMand)-1)
	end if
	form1.txtSelMocMand.value = sMand
	'optional MOCs
	for i=0 to form1.lstMOCSelOpt.options.length-1
		sOpt = sOpt & form1.lstMOCSelOpt.options(i).value & ","
	next
	if sOpt<>"" then
		sOpt = left(sOpt,len(sOpt)-1)
	end if
	form1.txtSelMocOpt.value = sOpt
End function

Sub AddItem(list1,list2)
	dim obj,i,cnt,lastMoved
	cnt = list1.options.length-1
	for i = cnt to 0 step -1
		if list1.options(i).selected then
			set obj = document.createElement("OPTION")
			s = list1.options(i).text
			obj.text = s
			obj.value = list1.options(i).value
			for j=0 to list2.options.length-1
				if ucase(s) < ucase(list2.options(j).text) then exit for
			next
			list2.options.add obj,j
			list1.options.remove i
			
			lastMoved = i
		end if
	next
	if lastMoved > list1.options.length-1 then lastMoved = list1.options.length-1
	list1.selectedIndex = lastMoved
End Sub

Sub RemoveItem(list1,list2)
	dim obj,i,cnt,s,lastMoved
	cnt = list2.options.length-1
	for i = cnt to 0 step -1
		if list2.options(i).selected then
			set obj = document.createElement("OPTION")
			s = list2.options(i).text
			obj.text = s
			obj.value = list2.options(i).value
			for j=0 to list1.options.length-1
				if ucase(s) < ucase(list1.options(j).text) then exit for
			next
			list1.options.add obj,j
			list2.options.remove i
			
			lastMoved = i
		end if
	next
	if lastMoved > list2.options.length-1 then lastMoved = list2.options.length-1
	list2.selectedIndex = lastMoved
End Sub
-->
</SCRIPT>
<script language="Javascript" src="AutoComplete.js"></script>
</HEAD>
<BODY style="text-align:center" class="bgcolorlogin">
<!--#include file="menu_include.asp"-->
<h3><%= v_header %></h3>
<font color=red size=2><%= Request.QueryString("v_message")%></font>
<form name=form1  action=ves_tc_asgn_save.asp method=post>
<input type=hidden name=txtSelMocMand>
<input type=hidden name=txtSelMocOpt>
<table border=0 cellpadding=2 cellspacing=0 WIDTH=70%>
  <tr>
    <td><b>Vessel</b><br>
    <select class="menuHide" name=cmbVessel onkeypress="control_onkeypress():vesselChanged()" onblur="control_onblur()">
	<%
	while not rsVessel.eof%>
	  <option value="<%=rsVessel("vessel_code")%>"><%=rsVessel("vessel_name")%>
	<%rsVessel.movenext
	wend%>
	</select>
	<td width="150px">
	<td nowrap><b>Time charterer</b><br>
    <select class="menuHide" name=cmbTC onkeypress="control_onkeypress()" onblur="control_onblur()">
      <option value="<%%>">--Select Charterer--</option>
	<%
	while not rsTC.eof%>
	  <option value="<%=rsTC("time_charterer_id")%>"><%=rsTC("short_name")%>
	<%rsTC.movenext
	wend%>
	</select>
	&nbsp;&nbsp;&nbsp;<INPUT type="submit" value="Save" id=submit1 name=submit1>
  <tr><td colspan=3><b>Remarks</b>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<span>(<span id=lblCount>0</span> &nbsp;of max 500 chars)</span>
	<br>
    <textarea style="width:100%" name=txtRemarks><%=TCREMARKS%></textarea>
    <br>
  <tr>
    <td rowspan=2 style="vertical-align:top"><b>Un-assigned MOCs</b><br>
    <select size=21 name=lstMOC id=lstMOC style="width:285px"
		 onkeypress="control_onkeypress()" onblur="control_onblur()">
    <%
    while not rsMOC.eof%>
      <option value="<%=rsMOC("moc_id")%>"><%=rsMOC("short_name")%>
    <%rsMOC.movenext
    wend%>
    </select>
    <td style="text-align:center;">
		<button name=cmdAddItem1 onclick="AddItem form1.lstMOC,form1.lstMOCSelMand">&nbsp;&gt;&nbsp;</button><br>
		<br>
		<button name=cmdRemoveItem1 onclick="RemoveItem form1.lstMOC,form1.lstMOCSelMand">&nbsp;&lt;&nbsp;</button>
    <td style="vertical-align:top;"><b>Mandatory MOCs</b><br>
    <select size=15 name=lstMOCSelMand style="width:expression(form1.lstMOC.offsetWidth);height:expression(form1.lstMOC.offsetHeight/2 - 7);">
    <%
    rsMOCSel.filter = "mandatory=1"
    while not rsMOCSel.eof%>
      <option value="<%=rsMOCSel("moc_id")%>"><%=rsMOCSel("short_name")%>
    <%rsMOCSel.movenext
    wend%>
    </select>
  <tr>
    <td style="text-align:center;">
		<button name=cmdAddItem2 onclick="AddItem form1.lstMOC,form1.lstMOCSelOpt">&nbsp;&gt;&nbsp;</button><br>
		<br>
		<button name=cmdRemoveItem2 onclick="RemoveItem form1.lstMOC,form1.lstMOCSelOpt">&nbsp;&lt;&nbsp;</button>
    <td style="vertical-align:top;"><b>Optional MOCs</b><br>
    <select size=15 name=lstMOCSelOpt style="width:expression(form1.lstMOC.offsetWidth);height:expression(form1.lstMOC.offsetHeight/2 - 8);">
    <%
    rsMOCSel.filter = "mandatory=0"
    while not rsMOCSel.eof%>
      <option value="<%=rsMOCSel("moc_id")%>"><%=rsMOCSel("short_name")%>
    <%rsMOCSel.movenext
    wend%>
    </select>
</table>
<br>
<%
connObj.close
set connObj = nothing
%>
</FORM>
</BODY>
</HTML>