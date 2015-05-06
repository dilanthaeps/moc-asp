<%@  language="VBScript" %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->

<%	'===========================================================================
	'	Template Name	:	Inspection Request Maintenance
	'	Template Path	:	ins_request_maint.asp
	'	Functionality	:	To show the list of requests
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	31st August, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = true	
	dim SDATE1, SDATE2, dt1, Is_Sire
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
    <title>Vessel Inspections - Tanker Pacific</title>
    <meta http-equiv="expires" content="Tue, 20 Aug 2000 14:25:27 GMT">
    <link rel="stylesheet" href="moc.css"></link>
    <style>
        .clsIncident {
            font-family: webdings;
            font-size: 15px;
            color: red;
            cursor: default;
        }

        .clsHilite {
            color: red;
            font-size: 12px;
            font-weight: bold;
            font-style: italic;
        } 
    </style>
    <script language="Javascript" src="js_date.js"></script>
    <script language="Javascript" src="AutoComplete.js"></script>
    <script language="VBScript" src="vb_date.vs"></script>
    <script id="clientEventHandlersJS" language="javascript">
<!--
        //add test comment for Github
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
        document.form1.action="ins_request_maint.asp?item="+v_sort_field+"&order="+v_sort_order
        document.form1.submit();
    }
    function v_clear_all_filters()
    {
        location.href="ins_request_maint.asp";
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
        var divCountZero = document.getElementById('divCountZero');
        var divCountFirst = document.getElementById('divCountFirst');

        <%if request.TotalBytes=0 then%>   
            divCountFirst.innerHTML = ""
        divCountZero.innerHTML = "<font color=red><b>No filters specified. Please specify filters and click refresh.</b></font>"
        form1.fleet_code.value = "ACTIVE_VESSELS"
    <%else%>
        divCountZero.innerHTML = divCountFirst.innerHTML
        form1.fleet_code.value = "<%=request("fleet_code")%>"
    <%end if%>
    <%if request("vessel_code")="" then%>
        fleet_code_onchange()
    <%end if%>
    }

    function fleet_code_onchange() {
        lblVessel.innerHTML = vessel_code(form1.fleet_code.selectedIndex).outerHTML
        lblVessel.children(0).style.display=""
    }

    //-->
    </script>
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
<body class="bgcolorlogin" language="javascript" onload="return window_onload()">
    <!--#include file="menu_include.asp"-->
    <center>
<table border="0" width="100%">
<tr height="30pt" VALIGN="bottom">
		<td align="center">
		<h3 style="margin-bottom:0">Vessel Inspections </h3>			
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

'strSql = "SELECT    MIR.REQUEST_ID , MOC_FN_INS_REMARK_COUNT(request_id) remark_count,"
'strSql = strSql & " moc_fn_ins_active_remark_count(request_id) active_remark_count,"
'strSql = strSql & " MIR.VESSEL_CODE ,v.vessel_name,"
'strSql = strSql & " moc_fn_ins_active_def_count(request_id) active_def_count, "
'strSql = strSql & " MOC_FN_DEF_COUNT(request_id) def_count,MIR.MOC_ID ,"
'strSql = strSql & " mm.short_name moc_name,"
'strSql = strSql & " MIR.INSP_STATUS, to_char(MIR.INSPECTION_DATE,'DD-Mon-YYYY') inspection_date1,"
'strSql = strSql & " MIR.INSPECTION_PORT, to_char(MIR.EXPIRY_DATE,'DD-Mon-YYYY') expiry_date1,"
'strSql = strSql & " nvl(trunc(expiry_date),trunc(sysdate))-trunc(sysdate) diff_days, insp_type,"
'strSql = strSql & " substr(moc_fn_basis_sire_short_name(mir.basis_sire), 1, 255) basis_sire_name,"
'strSql = strSql & " nvl(moc_fn_vessel_age_years(mir.vessel_code),0)age"
'strSql = strSql & " FROM MOC_INSPECTION_REQUESTS MIR, vessels v,moc_master mm"
'strSql = strSql & " where mir.vessel_code=v.vessel_code"
'strSql = strSql & " and mir.moc_id=mm.moc_id"

strSql = " SELECT   mir.request_id, mir.vessel_code, v.vessel_name,detention,upper(mi.short_name) short_name,mir.inspector_id,expences_in_usd, mir.Is_Sire,"
strSql = strSql &  " 		 nvl(rem.remark_count,0)remark_count,"
strSql = strSql &  "          nvl(arem.active_remark_count,0)active_remark_count,"
strSql = strSql &  "          def.def_count, adef.active_def_count,"
strSql = strSql &  "          mir.moc_id, mm.short_name moc_name, mir.insp_status,"
strSql = strSql &  "          TO_CHAR (mir.inspection_date, 'DD-Mon-YYYY') inspection_date1,"
strSql = strSql &  "          mir.inspection_port,"
strSql = strSql &  "          TO_CHAR (mir.expiry_date, 'DD-Mon-YYYY') expiry_date1,"
strSql = strSql &  "            NVL (TRUNC (expiry_date), TRUNC (SYSDATE))"
strSql = strSql &  "          - TRUNC (SYSDATE) diff_days, insp_type,"
strSql = strSql &  "          SUBSTR(moc_fn_basis_sire_short_name (mir.basis_sire),1,255) basis_sire_name,"
strSql = strSql &  "          NVL (moc_fn_vessel_age_years (mir.vessel_code), 0) age, app.STATUS appStatus"
strSql = strSql &  "     FROM moc_inspection_requests mir, vessels v, moc_master mm,moc_inspectors mi,moc_approvals app,"
strSql = strSql &  " 	(select request_id,count(*)remark_count from moc_request_remarks group by request_id)rem,"
strSql = strSql &  " 	(select request_id,count(*)active_remark_count from moc_request_remarks where remark_status='Active' group by request_id)arem,"
strSql = strSql &  " 	(select request_id,count(*)def_count from moc_deficiencies group by request_id)def,"
strSql = strSql &  " 	(select request_id,count(*)active_def_count from moc_deficiencies where status = 'Pending' group by request_id)adef"
strSql = strSql &  "    WHERE mir.vessel_code = v.vessel_code"
strSql = strSql &  "      AND mir.moc_id = mm.moc_id"
strSql = strSql &  " 	 AND mir.request_id = rem.request_id(+)"
strSql = strSql &  "  	 AND mir.request_id = arem.request_id(+)"
strSql = strSql &  "  	 AND mir.request_id = def.request_id(+)"
strSql = strSql &  "  	 AND mir.request_id = adef.request_id(+)"
strSql = strSql &  "     and mir.inspector_id=mi.inspector_id(+)"
strSql = strSql &  "     and mir.request_id = app.REQUEST_ID(+)"
	 
if request("vessel_code")<>"" then
	strSql=strSql & " and mir.vessel_code='" & request("vessel_code") & "'"
	v_filter="Vessel "
elseif Request("fleet_code")="ACTIVE_VESSELS" then
	'strSql=strSql & " and MIR.vessel_code in(select vessel_code from wls_vw_vessels_new where fleet_code in('AFRAMAX','PRODUCT','SUEZMAX','VLCC','BULK','CHEMICAL','LPG','CONTAINER','PCTC'))"
	strSql=strSql & " and MIR.vessel_code in(select vessel_code from wls_vw_vessels_new)"
	v_filter="Fleet "
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

if detention<>"" then
	strSql=strSql & " and detention='" & detention & "'"
end if

if request("moc_id")<>"" then
	strSql=strSql & " and mir.moc_id=" & request("moc_id")
	if v_filter <> "" then
		v_filter=v_filter&" , MOC "
	else
		v_filter=v_filter&" MOC "
	end if
end if

if request("insp_status")<>"" then
	strSql=strSql & " and insp_status='" & request("insp_status")& "'"
	if v_filter <> "" then
		v_filter=v_filter&" , Inspection Status "
	else
		v_filter=v_filter&" Inspection Status "
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


v_remarks_filter_condition = ""
If	Trim(Request("v_remarks_filter")) <> "" Then
	v_remarks_filter_condition = " and moc_fn_ins_active_remark_count(mir.request_id) > 0 "
End If

v_defi_filter_condition = ""
If	Trim(Request("v_defi_filter")) <> "" Then
	v_defi_filter_condition = " and moc_fn_ins_active_def_count(mir.request_id) > 0 "
End If

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

<%rsObj_vessel.filter="fleet_code='AFRAMAX' or fleet_code='FSO' or fleet_code='PRODUCT' or fleet_code='SUEZMAX' or fleet_code='VLCC' or fleet_code='BULK' or fleet_code='CONTAINER' or fleet_code='PCTC' or fleet_code='CHEMICAL' or fleet_code='LPG'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All ACTIVE vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter=0%>
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
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All AFRAMAX vessels--</option>
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
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All FSO vessels--</option>
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
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All PRODUCT vessels--</option>
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
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All SUEZMAX vessels--</option>
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
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All VLCC vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='BULK'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All BULK vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='CONTAINER'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All CONTAINER vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='PCTC'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All PCTC vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='CHEMICAL'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All CHEMICAL vessels--</option>
  <%
  while not rsObj_vessel.eof
  %>
  <option value="<%=rsObj_vessel("vessel_code")%>"><%=rsObj_vessel("vessel_name")%>
  <%
	rsObj_vessel.movenext
  wend
  %>
</select>
<%rsObj_vessel.filter="fleet_code='LPG'"%>
<select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
				onkeypress="control_onkeypress()" onblur="control_onblur()" style="display:none">
  <option value="<%%>">--All LPG vessels--</option>
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
<form name="form1" method="post" action="ins_request_maint.asp">
<table WIDTH="70%" border="0" cellspacing="1" cellpadding="1" align="center">
  <tr>
    <td class="tableheader">From Inspection Date
    <td class="tableheader">To Inspection Date
	<td class="tableheader">Select Type
    <td class="tableheader">Select Status
    <td class="tableheader">Detention 
    <td class="tableheader">Select Insp. Status      
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
            <select name="detention" style="width:100%">
				<option value="<%%>">All</option>
				<option value="YES" <%if detention="YES" then Response.Write " selected"%>>Yes
				<option value="NO" <%if detention="NO" then Response.Write " selected"%>>No
			</select>
      </td>
      <td class="tabledata">
          <select class="menuHide" id="insp_status" name="insp_status" onchange="javascript:v_select('insp_status');">
             <option value="<%%>">--Show All--</option>
				<%
					if not(rsObj_insp_status.eof or rsObj_insp_status.bof) then
						while not rsObj_insp_status.eof
				%>
						<option value="<%=rsObj_insp_status("sys_para_id")%>" <%if request("insp_status")=rsObj_insp_status("sys_para_id") then Response.Write " selected" %>><%=rsObj_insp_status("insp_status")%></option>
				<%		rsObj_insp_status.movenext
						wend
					end if
				%>
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
			   <option value="ACTIVE_VESSELS">ACTIVE VESSELS</option>
			   <option value="<%%>">ALL VESSELS</option>
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
	  </td>
      <td class="tabledata">
             <span id=lblVessel>
             <select id="vessel_code" name="vessel_code" onchange="javascript:v_select('vessel_code');"
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
		<input type="button" value="Create New Inspection" <% =v_button_disabled %> NAME="v_new_record" onclick="javascript:fncall('0');" class="cmdButton" style="width:150px">
	</td>
	<tr>
	<td ALIGN="right">
		<a href="javascript:window.excelOut()" OnClick="return outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
		<a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
	</td>
	</table>
	<%if UserIsAdmin then%>
		<div style="float:left;width:75%"><span style="font-size:11px;font-weight:bold">Notes for <%=noteVessel%></span><br>
		<textarea id="txtNote" style="width:100%;border:1px solid lightgrey;overflow:hidden;background-color:white;
		 font-family:arial;font-size:11px;font-weight:bold;color:maroon" rows=3
		 onblur="javascript:UpdateNote()"><%=mocNote%></textarea></span>
	<%end if%>
  </tr>
  <tr>
	<td colspan="2"><div id=divCountZero align="left"></div>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td class="tableheader" VALIGN="top" nowrap>Vessel<br>
		<table width="100%" border="0"><tr HEIGHT="2pt"></tr><tr>
		<td align="left">
			<a href="javascript:v_sort('vessel_name','asc');">
			<img SRC="Images/up.gif" ALT="Sort Ascending by Vessel" border="0" width="15" align="top">
			</a>
		</td>
		<td align="right">
		<a href="javascript:v_sort('vessel_name','desc');">
		<img SRC="Images/down.gif" ALT="Sort Descending by Vessel" border="0" width="15" align="top">
		</a>
		</td>
		</tr></table>
	</td>
    <td class="tableheader" valign="top">Followups<br>
		<table width="100%" border="0"><tr HEIGHT="2pt"></tr><tr>
		<td align="center">
			<select NAME="v_remarks_filter" CLASS="ddlist" STYLE="width:35pt" OnChange="document.form1.submit();">
<%
				strSqlRemStatus = "select sys_para_id, para_desc, upper(para_desc) sort_column "
				strSqlRemStatus = strSqlRemStatus & "from moc_system_parameters "
				strSqlRemStatus = strSqlRemStatus & "where upper(trim(parent_id)) = 'REMARK_STATUS' "
				strSqlRemStatus = strSqlRemStatus & "and upper(trim(sys_para_id)) <> 'COMPLETED' "
				strSqlRemStatus = strSqlRemStatus & "union "
				strSqlRemStatus = strSqlRemStatus & "select '' sys_para_id, 'All' para_desc, 'AAAAAAA' sort_column "
				strSqlRemStatus = strSqlRemStatus & "from dual "
				strSqlRemStatus = strSqlRemStatus & "order by 3"
				
				Response.Write strSqlRemStatus & "<BR>"
				Set rsObjRemStatus = connObj.Execute(strSqlRemStatus)
				
				While Not rsObjRemStatus.EOF
					
					v_selected = ""
					If Trim(Request("v_remarks_filter")) = Trim(rsObjRemStatus("sys_para_id")) Then
						v_selected = "SELECTED"
					End If
%>
					<option VALUE="<% =Trim(rsObjRemStatus("sys_para_id")) %>" <% =v_selected %>><% =Trim(rsObjRemStatus("para_desc")) %></option>
<%
					rsObjRemStatus.MoveNext
				Wend

				rsObjRemStatus.Close
				Set rsObjRemStatus = Nothing
%>
			</select>

		</td>
		</tr></table>
    </td>
    <td class="tableheader" valign="top">Observations<br>
		<table width="100%" border="0"><tr HEIGHT="2pt"></tr><tr>
		<td align="center">
			<select NAME="v_defi_filter" CLASS="ddlist" STYLE="width:50pt" OnChange="document.form1.submit();">
<%
				strSqlDefStatus = "select sys_para_id, para_desc, upper(para_desc) sort_column "
				strSqlDefStatus = strSqlDefStatus & "from moc_system_parameters "
				strSqlDefStatus = strSqlDefStatus & "where upper(trim(parent_id)) = 'DEFICIENCY_STATUS' "
				strSqlDefStatus = strSqlDefStatus & "and upper(trim(sys_para_id)) <> 'COMPLETED' "
				strSqlDefStatus = strSqlDefStatus & "union "
				strSqlDefStatus = strSqlDefStatus & "select '' sys_para_id, 'All' para_desc, 'AAAAAAA' sort_column "
				strSqlDefStatus = strSqlDefStatus & "from dual "
				strSqlDefStatus = strSqlDefStatus & "order by 3"
				
				'Response.Write strSqlDefStatus & "<BR>"
				Set rsObjDefStatus = connObj.Execute(strSqlDefStatus)
				
				While Not rsObjDefStatus.EOF
					
					v_selected = ""
					If Trim(Request("v_defi_filter")) = Trim(rsObjDefStatus("sys_para_id")) Then
						v_selected = "SELECTED"
					End If
%>
					<option VALUE="<% =Trim(rsObjDefStatus("sys_para_id")) %>" <% =v_selected %>><% =Trim(rsObjDefStatus("para_desc")) %></option>
<%
					rsObjDefStatus.MoveNext
				Wend

				rsObjDefStatus.Close
				Set rsObjDefStatus = Nothing
%>
			</select>

		</td>
		</tr></table>
    </td>
    <td class="tableheader" valign="top">MOC<br>
		<table width="100%"><tr HEIGHT="3pt"></tr><tr>
		<td align="left">
		<a href="javascript:v_sort('moc_name','asc');">
		<img SRC="Images/up.gif" ALT="Sort Ascending by MOC" border="0" width="15">
		</a>
		</td>
		<td align="right">
		<a href="javascript:v_sort('moc_name','desc');">
		<img SRC="Images/down.gif" ALT="Sort Descending by MOC" border="0" width="15">
		</a>
		</td>
		</tr></table>
    </td>

    <td class="tableheader" valign="top">Is Sire</Td>


    <td class="tableheader" valign="top">Status <br>
		<table width="100%"><tr HEIGHT="3pt"></tr><tr>
		<td align="left">
		<a href="javascript:v_sort('insp_status','asc');">
		<img SRC="Images/up.gif" ALT="Sort Ascending by Status" border="0" width="15">
		</a>
		</td>
		<td align="right">
		<a href="javascript:v_sort('insp_status','desc');">
		<img SRC="Images/down.gif" ALT="Sort Descending by Status" border="0" width="15">
		</a>
		</td>
		</tr></table>
    </td>
    <td class="tableheader" valign="top">Inspection Date<br>
		<table width="100%"><tr HEIGHT="3pt"></tr><tr>
		<td align="left">
		<a href="javascript:v_sort('inspection_date','asc');">
		<img SRC="Images/up.gif" ALT="Sort Ascending by Inspection Date" border="0" width="15">
		</a>
		</td>
		<td align="right">
		<a href="javascript:v_sort('inspection_date','desc');">
		<img SRC="Images/down.gif" ALT="Sort Descending by Inspection Date" border="0" width="15">
		</a>
		</td>
		</tr></table>
    </td>
    <td class="tableheader" valign="top">Inspection Port<br>
		<table width="100%"><tr HEIGHT="3pt"></tr><tr>
		<td align="left">
		<a href="javascript:v_sort('inspection_port','asc');">
		<img SRC="Images/up.gif" ALT="Sort Ascending by Inspection Port" border="0" width="15">
		</a>
		</td>
		<td align="right">
		<a href="javascript:v_sort('inspection_port','desc');">
		<img SRC="Images/down.gif" ALT="Sort Descending by Inspection Port" border="0" width="15">
		</a>
		</td>
		</tr></table>
    </td>
    <%if request("cmbInspector")<>"" then %>
      <td class="tableheader"  valign="top">Inspector</td>
    <%end if %>
    
    <td class="tableheader" valign="top">Expiry<br>
		<table width="100%" border="0"><tr HEIGHT="2pt"></tr><tr>
		<td align="center" class="menuHide">
			<select NAME="v_expiry_filter" CLASS="ddlist" STYLE="width:60pt" OnChange="document.form1.submit();">
				<option VALUE <% =IIF(Trim(Request("v_expiry_filter")) = "", "SELECTED", "") %>>All</option>
				<option VALUE="2months" <% =IIF(Trim(Request("v_expiry_filter")) = "2months", "SELECTED", "") %>>Due in 2 Months</option>
				<option VALUE="overdue" <% =IIF(Trim(Request("v_expiry_filter")) = "overdue", "SELECTED", "") %>>Overdue</option>
			</select>
		</td>
		</tr></table>
    </td>
    <!--<%if UserIsAdmin then %>
    <td class="tableheader"  valign="top">Incentive</td>
    <%end if %>-->
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
    <td class="<%=class_color%>"><a name="1" href="javascript:fncall('<%=rsObj("request_id") %>');"><%=rsObj("vessel_name")%></a></td>
    <td class="<%=class_color%>" align="center">
		<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
		<tr VALIGN="middle"><td VALIGN="middle" ALIGN="right">
	    <% if rsObj("active_remark_count")>"0" then
		%>
		<img SRC="Images/red_triangle.gif" title="Total Remarks = <%=rsObj("remark_count")%> &amp; Pending Remarks = <%=rsObj("active_remark_count")%>" WIDTH="12" HEIGHT="9">
		<%
		elseif rsObj("remark_count")>"0"	then
		%>
		<font FACE="Wingdings" SIZE="4" COLOR="green" title="Total Remarks =  <%=rsObj("remark_count")%> &amp; No Pending Remarks"><b><% =Chr(252) %></b></font>
		<%
		else
		%>
		&nbsp;
		<%
		end if
		%>
		</td>
		<td VALIGN="bottom" ALIGN="right">
			<input TYPE="image" SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit Remarks !" OnClick="javascript:return fncall_remark('<%=rsObj("request_id")%>','<%=rsObj("vessel_code")%>','<%=rsObj("vessel_name")%>','<%=rsObj("moc_id")%>','<%=replace(rsObj("moc_name"),"'","\'")%>', '<%=rsObj("inspection_date1")%>', '<%=escape(rsObj("inspection_port")&"")%>');" WIDTH="11" HEIGHT="16">
			</td>
		</tr>
		</table>
	</td>
    <td class="<%=class_color%>" align="center">
		<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
		<tr VALIGN="middle"><td VALIGN="middle" ALIGN="right">
	<% if rsObj("active_def_count")>"0" then
		%>
		<img SRC="Images/red_triangle.gif" title="Total Observations = <%=rsObj("def_count")%> &amp; Pending Observations = <%=rsObj("active_def_count")%>" WIDTH="12" HEIGHT="9">
		<%
		elseif rsObj("def_count")>"0"	then
		%>
		<font FACE="Wingdings" SIZE="4" COLOR="green" title="Total Observations =  <%=rsObj("Def_count")%> &amp; No Pending Observations "><b><% =Chr(252) %></b></font>
		<%
		else
		%>
		&nbsp;
		<%
		end if
		%>
		</td>
		<td VALIGN="bottom" ALIGN="right">
			<input TYPE="image" SRC="Images/click_to_open.gif" TITLE="Click to Add / Edit Observations !" OnClick="javascript:return fncall_def('<%=rsObj("request_id")%>','<%=rsObj("vessel_code")%>','<%=rsObj("vessel_name")%>','<%=rsObj("moc_id")%>','<%=replace(rsObj("moc_name"),"'","\'")%>', '<%=rsObj("inspection_date1")%>', '<%=escape(rsObj("inspection_port")&"")%>');" WIDTH="11" HEIGHT="16">
			</td>
		</tr>
		</table>
    </td>

    <td class="<%=class_color%>"><%=rsObj("moc_name") %>&nbsp;
    <%if rsOBj("insp_type")="INCIDENT" then%>
    <span class=clsIncident title="Incident Report"><%=chr(105)%></span>
    <%end if%>
    <%if rsOBj("detention")="YES" then%>
    <span class=clsIncident title="Detention"><%=chr(110)%></span>
    <%end if%>
    </td>

    <td class="<%=class_color%>" Align=center><%=rsObj("Is_Sire") %></td>

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
		
		'highlight only for MOC inspections
		'if rsObj("insp_type")<>"MOC" then v_diff_days=1
		
		'if csng(rsObj("age"))<=10 and v_diff_days >= 6 then
			'v_style_start = "<FONT COLOR='red' title='Ship less than 10 years old and inspected more than 6 months ago'>"
			'v_style_end = "</FONT>"
		'elseif csng(rsObj("age"))>10 and v_diff_days >= 4 then
			'v_style_start = "<FONT COLOR='red' title='Ship more than 10 years old and inspected more than 4 months ago'>"
			'v_style_end = "</FONT>"
		'end if
		
		'if rsObj("insp_type")<>"MOC" then v_diff_days=1
		
		if csng(rsObj("age"))<10 and v_diff_days >= 6 and rsObj("insp_status") ="ACCEPTED" then
			v_style_start = "<FONT COLOR='red' title='Ship less than 10 years old and inspected more than 6 months ago'>"
			v_style_end = "</FONT>"
		elseif csng(rsObj("age"))>=10 and v_diff_days >= 4 and rsObj("insp_status") ="ACCEPTED" then
			v_style_start = "<FONT COLOR='red' title='Ship more than 10 years old and inspected more than 4 months ago'>"
			v_style_end = "</FONT>"
		end if
%>		
    <%=v_style_start & mid(rsObj("inspection_date1"),1,7) & mid(rsObj("inspection_date1"),10,2) & v_style_end %>&nbsp;</td>
    <td class="<%=class_color%>"><%=rsObj("inspection_port") %>&nbsp;</td>
    <%if request("cmbInspector")<>"" then %>
      <td class="<%=class_color%>"><%=rsObj("short_name")%>&nbsp;</td>
     <%end if%>
     
    <td class="<%=class_color%>" nowrap>
<%
		v_style_start = ""
		v_style_end = ""
		v_diff_days=clng(rsObj("diff_days"))
		if v_diff_days < 0 then
			v_style_start = "<FONT SIZE='2' COLOR='red' title='Expired'><B><I>"
			v_style_end = "</I></B></FONT>"
		elseif v_diff_days <= 60 then
			v_style_start = "<FONT SIZE='2' COLOR='red' title='Due to expire in 2 months'><B><I>"
			v_style_end = "</I></B></FONT>"
		end if
%>		
		<%=v_style_start & mid(rsObj("expiry_date1"),1,7) & mid(rsObj("expiry_date1"),10,2) & v_style_end %>&nbsp;
		
		</td>		
        <!--Modified by Sankar<%if UserIsAdmin then %>		
            <td class="<%=class_color%>" align="center">
	            <%= rsObj("appstatus")%>	    
            </td>
        <%end if %>-->
  </tr>

<%
  rsObj.movenext
wend
else
Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if

%>
</table>
<div id=divCountFirst align="left"><strong>Number of inspections :</strong> <%=v_ctr%></div>
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