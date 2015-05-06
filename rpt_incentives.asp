<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->

<%	
	Response.Buffer = true	
	dim SDATE1, SDATE2, dt1
	dim status, detention, insp_type_disp, v_Page, v_mess, v_button_disabled
	dim v_filter, v_remarks_filter_condition, v_defi_filter_condition
	dim v_expiry_filter_condition, v_selected, v_ctr,inspector,inspector1
	dim rsObj_Vessel, rsObj_moc, rsObj_insp_status, rsObj_status, rsObj_insp_type,rsInspector
	dim rsObj_port, strSqlRemStatus, rsObjRemStatus, strSqlDefStatus
	dim rsObjDefStatus, class_color, v_style_start, v_style_end, v_diff_days
    
    
	if request("status")="" or isnull(request("status")) then
		status = "ACTIVE"
	else
		status = request("status")
	end if

	if request("status")="All"  then
		status=""
	end if
	
	detention = request("detention")	
	
	if request("insp_type")="" or request("insp_type")= null  then
		insp_type_disp=""
	else
		insp_type_disp=request("insp_type")
	end if

	if insp_type_disp = "All"  then
		insp_type_disp=""
	end if	

	if insp_type_disp = "All"  then
		insp_type_disp=""
	end if
	
	SDATE1 = request("from_inspection_date")
	SDATE2 = request("to_inspection_date")
	if SDATE1="" and SDATE2="" then
		dt1 = "1 " & monthname(month(now),true) & " " & year(now)
		dt1 = CDate(dt1)
		dt1 = DateAdd("m",-6,dt1)
		'SDATE1 = FormatDateTimeValues(dt1,1)
		SDATE1 = "1 Jul 2002"
		SDATE2 = FormatDateTimeValues(now + 60,1)
	else
		if SDATE1="" then SDATE1 = SDATE2
		if SDATE2="" then SDATE2 = SDATE1
	end if

	Function IIF(expr, trueValue, falseValue)

		If expr Then
			IIF = trueValue
		Else
			IIF = falseValue
		End If
		
	End Function
	

%>
<html>
<head>
<title>Vessel Inspections - Incentive Report</title>
<meta HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 2000 14:25:27 GMT">
<link REL="stylesheet" HREF="moc.css"></link>
<style>
.clsIncident
{
	font-family:webdings;
	font-size:15px;
	color:red;
	cursor:default;
}
.clsHilite
{
	color:red;
	font-size:12px;
	font-weight:bold;
	font-style:italic;
}
</style>
<script language="Javascript" src="js_date.js"></script>
<script language="Javascript" src="AutoComplete.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function fncall(v_ins_request_id)
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes,resizable=yes,status=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=50,top=10,width=770,height=650'
	}else{
		winStats+=',screenX=350,screenY=200,width=400,height=280'
	}
	adWindow=window.open("ins_request_entry.asp?v_ins_request_id="+v_ins_request_id,"moc_request_entry",winStats);
	adWindow.focus();
}
function fncall_remark(v_ins_request_id,vessel_code,vessel_name,moc_id,moc_name, insp_date, insp_port)
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes,resizable=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=50,top=10,width=750,height=450'
	}else{
		winStats+=',screenX=350,screenY=200,width=400,height=280'
	}
	adWindow=window.open("ins_request_remark_maint.asp?v_ins_request_id=" + v_ins_request_id + "&VID=" + vessel_code + "&vessel_name=" + vessel_name + "&moc_id=" + moc_id + "&moc_name=" + moc_name + "&v_insp_date=" + insp_date + "&v_insp_port=" + insp_port, "moc_request_remark_entry", winStats);
	adWindow.focus();

	return false;
}
function fncall_def(v_ins_request_id,vessel_code,vessel_name,moc_id,moc_name, insp_date, insp_port)
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes,resizable=yes,fullscreen=no'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=50,top=10,width=720,height=650'
	}else{
		winStats+=',screenX=350,screenY=200,width=400,height=280'
	}
	adWindow=window.open("ins_request_def_maint.asp?v_ins_request_id=" + v_ins_request_id + "&VID=" + vessel_code + "&vessel_name=" + vessel_name + "&moc_id=" + moc_id + "&moc_name=" + moc_name + "&v_insp_date=" + insp_date + "&v_insp_port=" + insp_port,"moc_request_def_entry",winStats);
	adWindow.focus();
	return false;
}
function fnnotify(v_ins_request_id)
{
        winStats='Location=no, top=250, left=300, Height=250, Width=400'
       	//winStats='Location=no, top=100, left=100, Height=750, Width=450'
		adWindow=window.open("ins_vessel_notification.asp?v_ins_request_id="+v_ins_request_id,"moc_notification",winStats);
	    adWindow.focus();
	    return false;			
}
function v_select(v_field)
{
//alert(v_field)
//var v_field_value=eval("document.form1."+v_field+".value")
//document.form1.action="ins_request_maint.asp?"+v_field+"="+v_field_value
//alert(document.form1.action)
//document.form1.submit();
}
function v_sort(v_sort_field,v_sort_order)
{
	document.form1.action="rpt_incentives.asp?item="+v_sort_field+"&order="+v_sort_order
	document.form1.submit();
}
function v_clear_all_filters()
{
	location.href="rpt_incentives.asp";
}

function outputExcel()
{
	var preAction = document.form1.action;
	var preTarget = document.form1.target;
	var qryString = "";	
	qryString += "?status=" + document.form1.status.value;
	qryString += "&fleet_code=" + document.form1.fleet_code.value;
	qryString += "&vessel_code=" + document.form1.vessel_code.value;
	qryString += "&insp_type=" + document.form1.insp_type.value;
	qryString += "&moc_id=" + document.form1.moc_id.value;
	qryString += "&insp_status=" + document.form1.insp_status.value;
	qryString += "&inspection_port=" + document.form1.inspection_port.value;
	qryString += "&cmdInspector=" + document.form1.cmbInspector.value;
	qryString += "&from_inspection_date=" + document.form1.from_inspection_date.value;
	qryString += "&to_inspection_date=" + document.form1.to_inspection_date.value;
	qryString += "&v_remarks_filter=" + document.form1.v_remarks_filter.value;
	qryString += "&v_defi_filter=" + document.form1.v_defi_filter.value;
	qryString += "&v_expiry_filter=" + document.form1.v_expiry_filter.value;	
	qryString += "&item=" + "<% =Request("item") %>";
	qryString += "&order=" + "<% =Request("order") %>";
	
	document.form1.action = "ins_request_maint_excel.asp" + qryString;
	document.form1.target = "_blank";
	document.form1.submit();
	
	document.form1.action = preAction;
	document.form1.target = preTarget;

	return false;
}

function window_onload() {
}

function fleet_code_onchange() {
	lblVessel.innerHTML = vessel_code(form1.fleet_code.selectedIndex).outerHTML
	lblVessel.children(0).style.display=""
}

//-->
</SCRIPT>
<script language="vbscript">
dim colIndex,arrColor,iDir
colIndex=0
iDir=1
arrColor = Array("violet","darkviolet","indigo","blueviolet","blue","aquamarine","green","greenyellow","yellow","gold","orange","orangered","red","firebrick")
setInterval "HighlightText",100,"vbscript"
sub HighlightText()
	dim coll
	set coll = document.getElementsByName("lblHighlight")
	for each obj in coll
		obj.style.color = arrColor(colIndex)
	next
	if colIndex=ubound(arrColor) then iDir=-1
	if colIndex=lbound(arrColor) then iDir=1
	colIndex = colIndex + iDir
	'window.status = colIndex
end sub

dim objX
sub UpdateNote()
	if document.form1.vessel_code.value<>"" then
		set objX = CreateObject("MSXML2.XMLHttp.3.0")
		objX.open "POST","SaveNote.asp",false
		objX.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objX.send document.form1.vessel_code.value & "~" & document.getElementById("txtNote").value
		window.status = objX.responseText
	end if
end sub

</script>
</head>
<body class="bgcolorlogin" LANGUAGE=javascript onload="return window_onload()">
<!--#include file="menu_include.asp"-->
<center>
<table border="0" width="100%">
<tr height="30pt" VALIGN="bottom">
		<td align="center">
		<h3 style="margin-bottom:0">Vessel Inspections - Incentive Report</h3>			
			<%  v_mess=Request.QueryString("v_message")
			if v_mess <> "" then
			%>
			<br>
			<font color="red" size="+2"><%=v_mess%></font>
			<br>
			<% end if%>

		</td>
</tr>
</table>
<%    
	v_button_disabled = "DISABLED"
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Then
	if UserIsAdmin then
		v_button_disabled = ""
	End If
	
    v_filter =""
    
    strSql = ""
    strSql = strSql &  " SELECT mir.request_id, mir.vessel_code, v.vessel_name,"
    strSql = strSql &  "        UPPER (mi.short_name) short_name, mir.inspector_id,"
    strSql = strSql &  "        def.def_count,"
    strSql = strSql &  "        adef.active_def_count, mir.moc_id, mm.short_name moc_name,"
    strSql = strSql &  "        mir.insp_status,"
    strSql = strSql &  "        TO_CHAR (mir.inspection_date, 'DD-Mon-YYYY') inspection_date1,"
    strSql = strSql &  "        mir.inspection_port,"
    strSql = strSql &  "        SUBSTR(moc_fn_basis_sire_short_name (mir.basis_sire),1,255) basis_sire_name,"    
    strSql = strSql &  "        TO_CHAR (mir.expiry_date, 'DD-Mon-YYYY') expiry_date1,"
    strSql = strSql &  "        NVL (TRUNC (expiry_date), TRUNC (SYSDATE)) - TRUNC (SYSDATE) diff_days,"
    strSql = strSql &  "        NVL (moc_fn_vessel_age_years (mir.vessel_code), 0) age,insp_type,"
    strSql = strSql &  "        app.status IncentiveStatus, app.total_amount IncentiveAmt"
    strSql = strSql &  "   FROM moc_inspection_requests mir,"
    strSql = strSql &  "        vessels v,"
    strSql = strSql &  "        moc_master mm,"
    strSql = strSql &  "        moc_inspectors mi,"
    strSql = strSql &  "        moc_approvals app,"
    strSql = strSql &  "        (SELECT   request_id, COUNT (*) remark_count"
    strSql = strSql &  "             FROM moc_request_remarks"
    strSql = strSql &  "         GROUP BY request_id) REM,"
    strSql = strSql &  "        (SELECT   request_id, COUNT (*) active_remark_count"
    strSql = strSql &  "             FROM moc_request_remarks"
    strSql = strSql &  "            WHERE remark_status = 'Active'"
    strSql = strSql &  "         GROUP BY request_id) arem,"
    strSql = strSql &  "        (SELECT   request_id, COUNT (*) def_count"
    strSql = strSql &  "             FROM moc_deficiencies"
    strSql = strSql &  "         GROUP BY request_id) def,"
    strSql = strSql &  "        (SELECT   request_id, COUNT (*) active_def_count"
    strSql = strSql &  "             FROM moc_deficiencies"
    strSql = strSql &  "            WHERE status = 'Pending'"
    strSql = strSql &  "         GROUP BY request_id) adef"
    strSql = strSql &  "  WHERE "
    strSql = strSql &  "    upper(mir.insp_type) = 'MOC'"
    strSql = strSql &  "    AND upper(mir.insp_status) <> 'ACCEPTED BASED SIRE'"
    strSql = strSql &  "    AND mir.vessel_code = v.vessel_code"
    strSql = strSql &  "    AND mir.moc_id = mm.moc_id"
    strSql = strSql &  "    AND mir.request_id = REM.request_id(+)"
    strSql = strSql &  "    AND mir.request_id = arem.request_id(+)"
    strSql = strSql &  "    AND mir.request_id = def.request_id(+)"
    strSql = strSql &  "    AND mir.request_id = adef.request_id(+)"
    strSql = strSql &  "    AND mir.inspector_id = mi.inspector_id(+)"
    strSql = strSql &  "    AND mir.request_id = app.request_id(+)"

	 
if request("vessel_code")<>"" then
	strSql=strSql & " and mir.vessel_code='" & request("vessel_code") & "'"
	v_filter="Vessel "
elseif Request("fleet_code")<>"" then
	strSql=strSql & " and MIR.vessel_code in(select vessel_code from wls_vw_vessels_new where fleet_code='" & request("fleet_code") &"')"
	v_filter="Fleet "
end if

if insp_type_disp="INCIDENT" then
	strSql=strSql & " and insp_type ='" & insp_type_disp & "'"
elseif insp_type_disp<>"" then
	strSql=strSql & " and mm.entry_type='" & insp_type_disp & "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Inspection Type "
	else
		v_filter=v_filter&" Inspection Type "
	end if
end if

if status<>"" then
	strSql=strSql & " and mir.status='" & status& "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Status "
	else
		v_filter=v_filter&" Status "
	end if
end if



if request("moc_id")<>"" then
	strSql=strSql & " and mir.moc_id=" & request("moc_id")
	if v_filter <> "" then
		v_filter=v_filter&" , MOC "
	else
		v_filter=v_filter&" MOC "
	end if
end if


if request("insp_status")<> "" then
    strSql=strSql & " and insp_status='" & request("insp_status")& "'"    
	if v_filter <> "" then
		v_filter=v_filter & " , Inspection Status "
	else
		v_filter=v_filter & " Inspection Status "
	end if
end if

if request("cboIncentiveStatus")<>"" then
	if trim(request("cboIncentiveStatus")) = "PENDING APPROVAL" THEN
	    strSql=strSql & " and app.status is NULL "    
	else
	    strSql=strSql & " and app.status='" & request("cboIncentiveStatus")& "'" 
	END IF

	
	if v_filter <> "" then
		v_filter=v_filter&" , Incentive "
	else
		v_filter=v_filter&" Incentive "
	end if
end if


if request("inspection_port")<>"" then
	strSql=strSql & " and inspection_port='" & request("inspection_port")& "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Inspection Port "
	else
		v_filter=v_filter&" Inspection Port "
	end if
end if

if request("cmbInspector")<>"" then
	strSql=strSql & " and mi.inspector_id='" & request("cmbInspector")& "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Inspector "
	else
		v_filter=v_filter&"  Inspector "
	end if
end if

strSql=strSql & " and trunc(inspection_date) between '" & SDATE1 & "' and '" & SDATE2 & "'"
if v_filter <> "" then
	v_filter = v_filter & " , Inspection Date "
else
	v_filter=v_filter & " Inspection Date "
end if


v_expiry_filter_condition = ""
If	Trim(Request("v_expiry_filter")) <> "" Then
	If Trim(Request("v_expiry_filter")) = "2months" Then
		v_expiry_filter_condition = " and nvl(trunc(expiry_date), trunc(sysdate)) - trunc(sysdate) between 0 and 60 and expiry_date is not null "
	ElseIf Trim(Request("v_expiry_filter")) = "overdue" Then
		v_expiry_filter_condition = " and nvl(trunc(expiry_date), trunc(sysdate)) - trunc(sysdate) < 0 "
	End If
End If

strSql = strSql & v_remarks_filter_condition & v_defi_filter_condition & v_expiry_filter_condition

If v_filter ="" Then
	v_filter = "No Filter (All records shown)"
End If

if request("item")<>"" then
	strSql=strSql & " order by " & request("item")& " " & request("order")
else
	strSql=strSql & " order by v.vessel_name, mm.short_name, INSPECTION_DATE "
end if

'Response.Write strSql 
'Response.End

if Request.TotalBytes=0 then
	Set rsObj=connObj.execute("Select * from moc_inspection_requests where 1=2")
else
	Set rsObj=connObj.execute(strSql)
end if

strSql="Select vessel_code, vessel_name, fleet_code from wls_vw_vessels_new where vessel_code in (select distinct vessel_code from moc_inspection_requests) order by vessel_name"

Set rsObj_vessel = connObj.execute(strSql)

strSql = "Select moc_id,short_name from moc_master where moc_id in (select distinct moc_id from moc_inspection_requests) order by short_name"
Set rsObj_moc=connObj.execute(strSql)

strSql = "Select distinct insp_status from moc_inspection_requests order by insp_status"
' Select List -      Inspection Status
strSql = "SELECT sys_para_id, para_desc insp_status,parent_id,sort_order "
strSql = strSql & "from moc_system_parameters "
strSql = strSql & "where parent_id = 'Status' "
strSql = strSql & "order by sort_order"
Set rsObj_insp_status=connObj.execute(strSql)

strSql = "Select distinct status from moc_inspection_requests order by status"
Set rsObj_status=connObj.execute(strSql)

'strSql = "Select distinct inspection_port from moc_inspection_requests order by inspection_port"
strSql="SELECT PORT_name ,port_country FROM PORTS_LIBRARY ORDER BY trim(upper(PORT_name)) "
Set rsObj_port=connObj.execute(strSql)

strSql="Select inspector_id,upper(short_name) short_name from moc_inspectors where inspector_id in "
strSql = strSql & "(select distinct inspector_id from moc_inspection_requests) order by short_name"
set rsInspector=connObj.execute(strSql)

dim rsTemp,mocNote,noteVessel
set rsTemp = connObj.execute("Select v.vessel_name, mvn.note from MOC_VESSEL_NOTES mvn, vessels v where mvn.vessel_code=v.vessel_code and mvn.vessel_code='" & request("vessel_code") & "'")
if not rsTemp.eof then
	noteVessel = rsTemp(0)
	mocNote = rsTemp(1)
end if
rsTemp.close
set rsTemp=nothing

%>

<!--start of hidden vessel combos for different fleets-->

<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='AFRAMAX'"%>
<select id="Select1" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='FSO'"%>
<select id="Select2" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='PRODUCT'"%>
<select id="Select3" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='SUEZMAX'"%>
<select id="Select4" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='VLCC'"%>
<select id="Select5" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<!--end of hidden vessel combos for different fleets-->
<!--h3 align="center">Selection Filter </h3-->
<form name="form1" method="post" action="rpt_incentives.asp">
<table WIDTH="70%" border="0" cellspacing="1" cellpadding="1" align="center">
  <tr>
    <td class="tableheader">From Inspection Date
    <td class="tableheader">To Inspection Date
	<td class="tableheader">Inspection Type
    <td class="tableheader">Select Status
    <td class="tableheader">Select Insp. Status      
    <td class="tableheader">Incentive Approval
  </tr>
  <tr>
    <td class="tabledata">
        <input type="text" id="from_inspection_date" value="<%=SDATE1%>" name="from_inspection_date" size="10" onblur="vbscript:checkDate from_inspection_date,'From Inspection Date','form1'" onchange="javascript:v_select('from_inspection_date');">&nbsp;<a HREF="javascript:show_calendar('form1.from_inspection_date',form1.from_inspection_date.value);"><img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0"></a>
      </td>
      <td class="tabledata">
        <input type="text" id="to_inspection_date" value="<%=SDATE2%>" name="to_inspection_date" size="10" onblur="vbscript:checkDate to_inspection_date,'To Inspection Date','form1'" onchange="javascript:v_select('to_inspection_date');">&nbsp;<a HREF="javascript:show_calendar('form1.to_inspection_date',form1.to_inspection_date.value);"><img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0"></a>
      </td>
    <td class="tabledata">
             <select name="insp_type">
				<option value="All">All Types</option>
				<option value="MOC" <%if insp_type_disp="MOC" then Response.Write " selected" %>>MOC</option>
				<option value="PSC" <%if insp_type_disp="PSC" then Response.Write " selected" %>>PSC</option>
				<option value="TMNL" <%if insp_type_disp="TMNL" then Response.Write " selected" %>>TMNL</option>
				<option value="TVEL" <%if insp_type_disp="TVEL" then Response.Write " selected" %>>TVEL</option>
				<option value="FLAG" <%if insp_type_disp="FLAG" then Response.Write " selected" %>>FLAG</option>
				<option value="INCIDENT" <%if insp_type_disp="INCIDENT" then Response.Write " selected" %>>INCIDENT</option>
				<option value="PRE-TC" <%if insp_type_disp="PRE-TC" then Response.Write " selected" %>>PRE-TC</option>
			</select>
      </td>
      <td class="tabledata">
             <select name="status">
				<option value="All">All</option>
				<%
					if not(rsObj_status.eof or rsObj_status.bof) then
						while not rsObj_status.eof
				%>
						<option value="<%=rsObj_status("status")%>" <%if status=rsObj_status("status") then Response.Write " selected" %>><%=rsObj_status("status")%></option>
				<%		rsObj_status.movenext
						wend
					end if
				%>
			</select>
      </td>

      <td class="tabledata">
          <select class="menuHide" id="insp_status" name="insp_status" onchange="javascript:v_select('insp_status');">
             <option value="<%%>">--Show All--</option>
				<%
					if not(rsObj_insp_status.eof or rsObj_insp_status.bof) then
						while not rsObj_insp_status.eof
				%>
						<option value="'<%=rsObj_insp_status("sys_para_id")%>'" <%if request("insp_status")=rsObj_insp_status("sys_para_id") then Response.Write " selected" %>><%=rsObj_insp_status("insp_status")%></option>
				<%		rsObj_insp_status.movenext
						wend
					end if
				%>
			</select>
      </td>
      <td class="tabledata">
          <select name="cboIncentiveStatus">
              <option value="">All</option>
              <option value="APPROVED" <%if request("cboIncentiveStatus")="APPROVED" then response.write " selected" %>>
                  APPROVED</option>
              <option value="DISAPPROVED" <%if request("cboIncentiveStatus")="DISAPPROVED" then response.write " selected" %>>
                  DISAPPROVED</option>
              <option value="PENDING APPROVAL" <%if request("cboIncentiveStatus")="PENDING APPROVAL" then response.write " selected" %>>
                  PENDING APPROVAL</option>
              <option value="VERIFY CREW" <%if request("cboIncentiveStatus")="VERIFY CREW" then response.write " selected" %>>
                  VERIFY CREW</option>
          </select>
      </td>
      
   </tr>
 </table>
<table border="0" cellspacing="1" cellpadding="1" align="center">
  <tr>
    <td class="tableheader">Select Fleet
    <td class="tableheader">Select Vessel
    <td class="tableheader">Select Inspecting Agency
    <td class="tableheader">Select Port
    <td class="tableheader">Select Inspector
  </tr>
  <tr>
      <td class="tabledata">
			 <select class="menuHide" id="fleet_code" name="fleet_code" LANGUAGE=javascript onchange="return fleet_code_onchange()">
			   <option value="<%%>">--All Fleets--</option>
			   <option value="AFRAMAX">AFRAMAX</option>
			   <option value="FSO">FSO</option>
			   <option value="PRODUCT">PRODUCT</option>
			   <option value="SUEZMAX">SUEZMAX</option>
			   <option value="VLCC">VLCC</option>
			 </select>
	  </td>
      <td class="tabledata">
             <span id=lblVessel>
             <select id="Select6" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" class="menuHide">
             <option value="<%%>">--All Vessels--</option>
				<%	rsObj_vessel.filter=0
					if not(rsObj_vessel.eof or rsObj_vessel.bof) then
						while not rsObj_vessel.eof
				%>
						<option value="<%=rsObj_vessel("vessel_code")%>" <%if request("vessel_code")=rsObj_vessel("vessel_code") then Response.Write " selected" %>><%=rsObj_vessel("vessel_name")%></option>
				<%		rsObj_vessel.movenext
						wend
					end if
				%>
			</select></span>
      </td>
      <td class="tabledata">
          <select id="moc_id" name="moc_id" onchange="javascript:v_select('moc_id');"]
			onkeypress="control_onkeypress()" onblur="control_onblur()" >
             <option value="<%%>">--All MOCs--</option>
				<%
					if not(rsObj_moc.eof or rsObj_moc.bof) then
						while not rsObj_moc.eof
				%>
						<option value="<%=rsObj_moc("moc_id")%>" <%if cstr(request("moc_id"))=cstr(rsObj_moc("moc_id")) then Response.Write " selected" %>><%=left(rsObj_moc("short_name"),20)%></option>
				<%		rsObj_moc.movenext
						wend
					end if
				%>
			</select>
      </td>
      <td class="tabledata">
          <select id="inspection_port" name="inspection_port" onchange="javascript:v_select('inspection_port');"
			onkeypress="control_onkeypress()" onblur="control_onblur()">
             <option value="<%%>">--All Ports--</option>
				<%if not(rsObj_port.eof or rsObj_port.bof) then
					while not rsObj_port.eof%>
						<option value="<%=rsObj_port("port_name")%>" <%if request("inspection_port")=rsObj_port("port_name") then Response.Write " selected" %>><%=rsObj_port("port_name")%></option>
				<%		rsObj_port.movenext
					wend
				end if%>
			</select>
      </td>
      <td class="tabledata">
          <select class="menuHide" id="cmbInspector" name="cmbInspector" onchange="javascript:v_select('inspector_id');"
			onkeypress="control_onkeypress()" onblur="control_onblur()">
             <option value="">--All Inspectors--</option>
				<% 
				while not rsInspector.eof				   
				    %>
						<option value="<%=rsInspector("inspector_id")%>" <%if cstr(request("cmbInspector"))=cstr(rsInspector("inspector_id")) then Response.Write "selected" %> ><%=rsInspector("short_name")%></option>
				<%		rsInspector.movenext
					wend%>				
			</select>
      </td>
  </tr>
  <tr>
	<td colspan="2" align="left">
		<input type="submit" value="Refresh" style="font-weight:bold" id="submit1" name="submit1" class="cmdButton">
		<!--<input type="button" value="Clear All Filters" onclick="Javascript:v_clear_all_filters();" class="cmdButton">-->
	</td>
	<TD colspan=3 rowspan=2>
	<table align=right width=24%><tr>
	<td align=right><br>	
		
	</td>
	<tr>
	<td ALIGN="right">
<!--		<a href="javascript:window.excelOut()" OnClick="return outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
		<a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>-->
	</td>
	</table>

  </tr>
  <tr>
	<td colspan="2"><div id=divCount align="left"></div>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
      <td class="tableheader" valign="top" style="white-space:nowrap">
          <a href="javascript:v_sort('vessel_name','asc');">
              <img src="Images/up.gif" alt="Sort Ascending by Vessel" border="0" width="15" align="top">
          </a>Vessel <a href="javascript:v_sort('vessel_name','desc');">
              <img src="Images/down.gif" alt="Sort Descending by Vessel" border="0" width="15"
                  align="top">
          </a>
      </td>
      <td class="tableheader" valign="top" style="white-space:nowrap">
          <a href="javascript:v_sort('moc_name','asc');">
              <img src="Images/up.gif" alt="Sort Ascending by MOC" border="0" width="15" align="top">
          </a>MOC <a href="javascript:v_sort('moc_name','desc');">
              <img src="Images/down.gif" alt="Sort Descending by MOC" border="0" width="15" align="top">
          </a>
      </td>
      <td class="tableheader" valign="top" style="white-space:nowrap">
          <a href="javascript:v_sort('insp_status','asc');">
              <img src="Images/up.gif" alt="Sort Ascending by Status" border="0" width="15" align="top">
          </a>Status <a href="javascript:v_sort('insp_status','desc');">
              <img src="Images/down.gif" alt="Sort Descending by Status" border="0" width="15" align="top">
          </a>
      </td>
      <td class="tableheader" valign="top" style="white-space:nowrap">
          <a href="javascript:v_sort('inspection_date','asc');">
              <img src="Images/up.gif" alt="Sort Ascending by Inspection Date" border="0" width="15" align="top">
          </a>Inspection Date <a href="javascript:v_sort('inspection_date','desc');">
              <img src="Images/down.gif" alt="Sort Descending by Inspection Date" border="0" width="15" align="top">
          </a>
      </td>
      <td class="tableheader" valign="top" style="white-space:nowrap">
          <a href="javascript:v_sort('inspection_port','asc');">
              <img src="Images/up.gif" alt="Sort Ascending by Inspection Port" border="0" width="15" align="top">
          </a>Inspection Port <a href="javascript:v_sort('inspection_port','desc');">
              <img src="Images/down.gif" alt="Sort Descending by Inspection Port" border="0" width="15" align="top">
          </a>
      </td>
    <%if request("cmbInspector")<>"" then %>
      <td class="tableheader"  valign="top">Inspector</td>
    <%end if %>
    
    
    <td class="tableheader"  valign="top">Incentive</td>
    
    <td class="tableheader"  valign="top">Amount(USD)</td>
  </tr>
</form>
<%
v_ctr=0
if not (rsObj.bof or rsObj.eof) then
while not rsObj.eof
v_ctr=v_ctr+1
	if (v_ctr mod 2) = 0 then
		class_color="columncolor2"
		else
		class_color="columncolor3"
	end if
%>
  
<tr valign="bottom">
    <td class="<%=class_color%>">
        <a name="1" href="javascript:fncall('<%=rsObj("request_id") %>');">
            <%=rsObj("vessel_name")%>
        </a>
    </td>

    
    <td class="<%=class_color%>">
    <input TYPE="image" SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit Observations !" OnClick="javascript:return fncall_def('<%=rsObj("request_id")%>','<%=rsObj("vessel_code")%>','<%=rsObj("vessel_name")%>','<%=rsObj("moc_id")%>','<%=replace(rsObj("moc_name"),"'","\'")%>', '<%=rsObj("inspection_date1")%>', '<%=escape(rsObj("inspection_port")&"")%>');" WIDTH="11" HEIGHT="16">
    <%=rsObj("moc_name") %>
    </td>
    <td class="<%=class_color%>" nowrap>  
<%
		v_style_start = ""
		v_style_end = ""
		v_diff_days=clng(rsObj("diff_days"))
		
		If v_diff_days < 0 then
			Response.Write "<span class='clsHilite'>Expired !</span>"
		Else
			If Isnull(rsObj("basis_sire_name")) Or rsObj("basis_sire_name") = Null Or rsObj("basis_sire_name") = "" Then
				if rsObj("insp_status")="TECHNICAL HOLD" then
					Response.Write "<span id='lblHighlight' style='font-weight:bold'>" & rsObj("insp_status") & "</span>"
				elseif rsObj("insp_status")="FAILED" then
					Response.Write "<span class='clsHilite'>Failed !</span>"
				else
					Response.Write v_style_start & replace(rsObj("insp_status"),"_"," ") & v_style_end
				end if
			Else
				Response.Write v_style_start & replace(rsObj("insp_status"),"_"," ") & " / <FONT SIZE='1' COLOR='darkblue'><B>" & rsObj("basis_sire_name") & "</B></FONT>" & v_style_end
			End If
		End If
%>
	</td>
    <td class="<%=class_color%>">
        <%
		v_style_start = ""
		v_style_end = ""

		v_diff_days = DateDiff("d",cdate(rsObj("inspection_date1")),now)
		v_diff_days = v_diff_days\30
		
		
		if csng(rsObj("age"))<10 and v_diff_days >= 6 and rsObj("insp_status") ="ACCEPTED" then
			v_style_start = "<FONT COLOR='red' title='Ship less than 10 years old and inspected more than 6 months ago'>"
			v_style_end = "</FONT>"
		elseif csng(rsObj("age"))>=10 and v_diff_days >= 4 and rsObj("insp_status") ="ACCEPTED" then
			v_style_start = "<FONT COLOR='red' title='Ship more than 10 years old and inspected more than 4 months ago'>"
			v_style_end = "</FONT>"
		end if
        %>
        <%=v_style_start & mid(rsObj("inspection_date1"),1,7) & mid(rsObj("inspection_date1"),10,2) & v_style_end %>
        &nbsp;</td>
    <td class="<%=class_color%>">
        <%=rsObj("inspection_port") %>
        &nbsp;</td>
    <%if request("cmbInspector")<>"" then %>
      <td class="<%=class_color%>"><%=rsObj("short_name")%>&nbsp;</td>
     <%end if%>
     
    
		
		
    <td class="<%=class_color%>" align="center">
	    <%= rsObj("IncentiveStatus")%>	    
    </td>
    <td class="<%=class_color%>" align="center">
	    <%IF rsObj("IncentiveStatus") = "APPROVED" then response.Write rsObj("IncentiveAmt")%>	    
    </td>

  </tr>
  
<%
  rsObj.movenext
wend
else
Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if

%>
</table>
<div id=div1 align="left"><strong>Number of inspections :</strong> <%=v_ctr%></div>
</body>
</html>
<%
rsObj.close
set rsObj = nothing

rsObj_vessel.close
set rsObj_vessel = nothing

rsObj_moc.close
set rsObj_moc = nothing

rsObj_insp_status.close
set rsObj_insp_status = nothing

rsObj_port.close
set rsObj_port = nothing


%>