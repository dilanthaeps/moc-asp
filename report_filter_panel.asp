<%@ Language=VBScript %>
<!--#include file="common_dbconn.asp"-->
<%
	v_tab_width = "90%"
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Reports Panel</TITLE>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>
<SCRIPT>
function clearFilters()
{
	document.v_form.v_fleet_code.value = "";
	document.v_form.v_insp_type.value = "";
	document.v_form.v_insp_status.value = "";
	document.v_form.v_vess_code.value = "";
	document.v_form.v_inspection_port.value = ""
	document.v_form.v_moc_id.value = ""
	document.v_form.v_inspector_id.value = "";
	document.v_form.v_agent_id.value = "";
	document.v_form.v_time_charterer_id.value = "";
	document.v_form.v_defi_status.value = "";
	document.v_form.v_insp_from_date.value = "";
	document.v_form.v_insp_to_date.value = "";
	document.v_form.v_expr_from_date.value = "";
	document.v_form.v_expr_to_date.value = "";
	submitFormFleet();

	return false;
}

function submitFormFleet()
{
	document.v_form.v_vess_code.value = "";
	document.v_form.submit();
}

function substStar(inVal)
{
	if (inVal == "") return "*"; else return inVal;
}

function prepareReportDate(inputDate)
{
	if (inputDate == "") return "";
	
	var dateValue = inputDate;
	var dayValue, monthValue, yearValue;
	var monthNames = new Array('Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec');

	dayValue = dateValue.substr(0, 2);
	monthValue = dateValue.substr(3, 3);
	yearValue = dateValue.substr(7, 4);

	for(i = 1; i <= 12; i ++)
		if (monthValue == monthNames[i]) monthValue = i;

	if (monthValue < "10") monthValue = "0" + monthValue;
	
	return dayValue + "/" + monthValue + "/" + yearValue;
}

function callReport()
{
	var params;
	var reportName = document.v_form.v_report_name;
	var reportFileName = reportName.value;

	if (reportName.value == "")
	{
		alert("Please choose a Report Name to View !");
		reportName.focus();
		return false;
	}
	
	if (document.v_form.v_insp_from_date.value != "" && document.v_form.v_insp_to_date.value == "")
	{
		alert("Please mention Inspection Date To when Inspection Date From is mentioned !");
		document.v_form.v_insp_to_date.focus();
		return false;
	}

	if (document.v_form.v_insp_to_date.value != "" && document.v_form.v_insp_from_date.value == "")
	{
		alert("Please mention Inspection Date From when Inspection Date To is mentioned !");
		document.v_form.v_insp_from_date.focus();
		return false;
	}

	if (document.v_form.v_expr_from_date.value != "" && document.v_form.v_expr_to_date.value == "")
	{
		alert("Please mention Expiry Date To when Expiry Date From is mentioned !");
		document.v_form.v_expr_to_date.focus();
		return false;
	}

	if (document.v_form.v_expr_to_date.value != "" && document.v_form.v_expr_from_date.value == "")
	{
		alert("Please mention Expiry Date From when Expiry Date To is mentioned !");
		document.v_form.v_expr_from_date.focus();
		return false;
	}

	if (reportFileName.indexOf(".asp") != -1 || reportFileName.indexOf(".ASP") != -1)
	{
		windName = String(parseInt(String(Math.random() * 1000)));

		switch(reportFileName)
		{
			case "rpt_moc_followup.asp" : reportFileName = reportFileName + "?VID=" + document.v_form.v_vess_code.value + "&FLEET=" + document.v_form.v_fleet_code.value + "&MOC=" + document.v_form.v_moc_id.value + "&INSP_DATE_FROM=" + document.v_form.v_insp_from_date.value + "&INSP_DATE_TO=" + document.v_form.v_insp_to_date.value + "&INSP_TYPE=" + document.v_form.v_insp_type.value + "&INSP_STATUS=" + document.v_form.v_insp_status.value + "&REMARK_STATUS=";break;
			
		}
		document.v_form.action = reportFileName;
		document.v_form.target = windName;
		document.v_form.submit();
		
		document.v_form.action = "";
		document.v_form.target = "";

		return false;		
	}

	if (reportFileName.indexOf(".rpt") != -1 || reportFileName.indexOf(".RPT") != -1)
	{
		subRep1 = "";
		subRep2 = "";
		subRep3 = "";
		
		if (reportFileName == "moc_weekly_forecast.rpt")
		{
			subRep1 = "moc_insp_requested_subrep.rpt";
			subRep2 = "moc_insp_planned_subrep.rpt";
		}

		if (reportFileName == "moc_inspections_pending.rpt")
		{
			subRep1 = "moc_pending_reply_subrep.rpt";
			subRep2 = "moc_vessels_inspected_subrep.rpt";
			subRep3 = "moc_pending_accept_subreport.rpt";
		}

		if (reportFileName == "moc_office_alert.rpt")
		{
			subRep1 = "moc_off_alert_not_pending_subrep.rpt";
			subRep2 = "moc_off_alert_pending_subrep.rpt";
		}
		
		filters_list = "";
		filters_list += document.v_form.v_fleet_code(document.v_form.v_fleet_code.selectedIndex).text + "~~";
		filters_list += document.v_form.v_insp_type(document.v_form.v_insp_type.selectedIndex).text + "~~";
		filters_list += document.v_form.v_insp_status(document.v_form.v_insp_status.selectedIndex).text + "~~";
		filters_list += document.v_form.v_vess_code(document.v_form.v_vess_code.selectedIndex).text + "~~";
		filters_list += document.v_form.v_inspection_port(document.v_form.v_inspection_port.selectedIndex).text + "~~";
		filters_list += document.v_form.v_moc_id(document.v_form.v_moc_id.selectedIndex).text + "~~";
		filters_list += document.v_form.v_inspector_id(document.v_form.v_inspector_id.selectedIndex).text + "~~";
		filters_list += document.v_form.v_agent_id(document.v_form.v_agent_id.selectedIndex).text + "~~";
		filters_list += document.v_form.v_time_charterer_id(document.v_form.v_time_charterer_id.selectedIndex).text + "~~";
		filters_list += document.v_form.v_defi_status(document.v_form.v_defi_status.selectedIndex).text + "~~";

		params = "repcallnew20.asp?rr1=" + document.v_form.v_report_name.value;
		params += "&rp1=" + substStar(document.v_form.v_fleet_code.value);
		params += "&rp2=" + substStar(document.v_form.v_insp_type.value);
		params += "&rp3=" + substStar(document.v_form.v_insp_status.value);
		params += "&rp4=" + substStar(document.v_form.v_vess_code.value);
		params += "&rp5=" + substStar(document.v_form.v_inspection_port.value);
		params += "&rp6=" + substStar(document.v_form.v_moc_id.value);
		params += "&rp7=" + substStar(document.v_form.v_inspector_id.value);
		params += "&rp8=" + substStar(document.v_form.v_agent_id.value);
		params += "&rp9=" + substStar(document.v_form.v_time_charterer_id.value);
		params += "&rp10=" + substStar(document.v_form.v_defi_status.value);
		params += "&rp11=" + substStar(prepareReportDate(document.v_form.v_insp_from_date.value));
		params += "&rp12=" + substStar(prepareReportDate(document.v_form.v_insp_to_date.value));
		params += "&rp13=" + substStar(prepareReportDate(document.v_form.v_expr_from_date.value));
		params += "&rp14=" + substStar(prepareReportDate(document.v_form.v_expr_to_date.value));
		params += "&rp15=" + substStar(filters_list);
		params += "&rp16=" + substStar("");
		params += "&rp17=" + substStar("");
		params += "&rp18=" + substStar("");
		params += "&rp19=" + substStar("");
		params += "&rp20=" + substStar("");
		params += "&rs1=" + substStar(subRep1);
		params += "&rs2=" + substStar(subRep2);
		params += "&rs3=" + substStar(subRep3);
		params += "&rs4="
		params += "&rs5="
		params += "&rs6="
		params += "&rs7="

		winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
		winStats += 'scrollbars=yes,resizable=yes'

		if (navigator.appName.indexOf("Microsoft") >= 0) 
		{
			winStats += ',left=0,top=0,width=' + (screen.width - 10) + ',height=' + (screen.height - 30)
		}
		else
		{
			winStats += ',screenX=350,screenY=200,width=350,height=180'
		}

		windName = String(parseInt(String(Math.random() * 1000)));

		//prompt("checking", params);
		//return false;

		repWind = window.open(params, windName, winStats);
		repWind.focus();

		return false;
	}

	return false;
}
</SCRIPT>
</HEAD>
<BODY CLASS="bcolor">
<!--#include file="menu_include.asp"-->
<H3>Reports Panel</H3>
<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="1" BORDER="0">
<FORM NAME="v_form" METHOD="post">
	<TR HEIGHT="20pt">
		<TD BGCOLOR="lightblue"><FONT SIZE="1"><B>Filters</B></FONT></TD>
	</TR>
	<TR HEIGHT="10pt"></TR>
</TABLE>
<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="1" BORDER="0">
	<TR HEIGHT="20pt">
		<TD WIDTH="10%" CLASS="tableheader">Vessel Group</TD>
		<TD WIDTH="20%" CLASS="tabledata">&nbsp;
			<SELECT NAME="v_fleet_code" OnChange="submitFormFleet();" 
				CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select fleet_code, fleet_name, upper(fleet_name) sort_column"
				strSql = strSql & " from wls_vw_fleet_list"
				strSql = strSql & " union"
				strSql = strSql & " select NULL fleet_code, ' ' fleet_name, ' AAAAAAA' sort_column from dual"
				strSql = strSql & " order by 3"

				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_fleet_code")) = rsObj("fleet_code") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("fleet_code") & "' " & v_selected & ">"
					Response.Write rsObj("fleet_name")
					Response.Write "</OPTION>"

					rsObj.MoveNext
				Wend				

				rsObj.Close
				Set rsObj = Nothing
%>		
			</SELECT>
		</TD>
		<TD WIDTH="10%" CLASS="tableheader">Inspection Type</TD>
		<TD WIDTH="20%" CLASS="tabledata">&nbsp;
			<SELECT NAME="v_insp_type" CLASS="ddlist" STYLE="width:135pt">
				<option value="<%%>"></option>
<%
				strSql = "select sys_para_id, para_desc, sort_order sort_column"
				strSql = strSql & " from moc_system_parameters"
				strSql = strSql & " where upper(trim(parent_id)) = 'INSPECTION_TYPE'"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				v_insp_type = ""
				If Trim(Request("v_insp_type")) <> "" Then
					v_insp_type = Trim(Request("v_insp_type"))
				End If

				While Not rsObj.EOF

					v_selected = ""
					If v_insp_type = rsObj("sys_para_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("sys_para_id") & "' " & v_selected & ">"
					Response.Write rsObj("para_desc")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">Vessel</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_vess_code" CLASS="ddlist" STYLE="width:135pt">
<%
				v_vess_condition = ""

				If Request("v_fleet_code") <> "" Then
					v_vess_condition = " where fleet_code = '" & trim(Request("v_fleet_code")) & "'"
				End If

				strSql = "select vessel_code, vessel_name, upper(vessel_name) sort_column"
				strSql = strSql & " from wls_vw_vessels_new"
				strsql = strSql & v_vess_condition
				strSql = strSql & " union"
				strSql = strSql & " select NULL vessel_code, ' ' vessel_name, 'AAAAAAA' sort_column from dual"
				strSql = strSql & " order by 3"

				'Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

	
				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_vess_code")) = Trim(rsObj("vessel_code")) Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("vessel_code") & "' " & v_selected & ">"
					Response.Write rsObj("vessel_name")
					Response.Write "</OPTION>"

					rsObj.MoveNext
				Wend				

				rsObj.Close
				Set rsObj = Nothing
%>		
			</SELECT>
		</TD>
		<TD CLASS="tableheader">Inspection Status</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_insp_status" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select sys_para_id, para_desc, sort_order sort_column"
				strSql = strSql & " from moc_system_parameters"
				strSql = strSql & " where upper(trim(parent_id)) = 'STATUS'"
				strSql = strSql & " union"
				strSql = strSql & " select NULL sys_para_id, ' ' para_desc, 0 sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_insp_status")) = rsObj("sys_para_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("sys_para_id") & "' " & v_selected & ">"
					Response.Write rsObj("para_desc")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">MOC</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_moc_id" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select to_char(moc_id) moc_id, short_name, upper(short_name) sort_column"
				strSql = strSql & " from moc_master"
				strSql = strSql & " union"
				strSql = strSql & " select NULL moc_id, ' ' short_name, 'AAAAAAA' sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_moc_id")) = rsObj("moc_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("moc_id") & "' " & v_selected & ">"
					Response.Write rsObj("short_name")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
		<TD CLASS="tableheader">Inspection Port</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_inspection_port" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select port port_code, port port_name, upper(port) sort_column"
				strSql = strSql & " from port"
				strSql = strSql & " union"
				strSql = strSql & " select ip.inspection_port port_code, ip.inspection_port port_name, upper(ip.inspection_port) sort_column"
				strSql = strSql & " from (select distinct inspection_port from moc_inspection_requests"
				strSql = strSql & " where upper(trim(inspection_port)) not in (select upper(trim(port)) from port)) ip"
				strSql = strSql & " union"
				strSql = strSql & " select NULL port_code, ' ' port_name, ' AAAAAAA' sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_inspection_port")) = rsObj("port_code") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("port_code") & "' " & v_selected & ">"
					Response.Write rsObj("port_name")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">Agent</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_agent_id" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select to_char(agent_id) agent_id, short_name, upper(short_name) sort_column"
				strSql = strSql & " from moc_agents_master"
				strSql = strSql & " union"
				strSql = strSql & " select NULL agent_id, ' ' short_name, '     ' sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_agent_id")) = rsObj("agent_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("agent_id") & "' " & v_selected & ">"
					Response.Write rsObj("short_name")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
		<TD CLASS="tableheader">Inspector</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_inspector_id" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select to_char(inspector_id) inspector_id, short_name, upper(short_name) sort_column"
				strSql = strSql & " from moc_inspectors"
				strSql = strSql & " union"
				strSql = strSql & " select NULL inspector_id, ' ' short_name, ' AAAAAAA' sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_inspector_id")) = rsObj("inspector_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("inspector_id") & "' " & v_selected & ">"
					Response.Write rsObj("short_name")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">Time Charterer</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_time_charterer_id" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select to_char(time_charterer_id) time_charterer_id, short_name, upper(short_name) sort_column"
				strSql = strSql & " from moc_time_charterers"
				strSql = strSql & " union"
				strSql = strSql & " select NULL time_charterer_id, ' ' short_name, 'AAAAAAA' sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_time_charterer_id")) = rsObj("time_charterer_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("time_charterer_id") & "' " & v_selected & ">"
					Response.Write rsObj("short_name")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
		<TD CLASS="tableheader">Deficiency Status</TD>
		<TD CLASS="tabledata">&nbsp;
			<SELECT NAME="v_defi_status" CLASS="ddlist" STYLE="width:135pt">
<%
				strSql = "select sys_para_id, para_desc, sort_order sort_column"
				strSql = strSql & " from moc_system_parameters"
				strSql = strSql & " where upper(trim(parent_id)) = 'DEFICIENCY_STATUS'"
				strSql = strSql & " union"
				strSql = strSql & " select NULL sys_para_id, ' ' para_desc, 0 sort_column"
				strSql = strSql & " from dual"
				strSql = strSql & " order by 3"
				
				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					v_selected = ""
					If Trim(Request("v_defi_status")) = rsObj("sys_para_id") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj("sys_para_id") & "' " & v_selected & ">"
					Response.Write rsObj("para_desc")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
			</SELECT>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">Inspection Date From</TD>
		<TD CLASS="tabledata">&nbsp;
			<INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<% =Request("v_insp_from_date") %>" SIZE="12"
				onblur="vbscript:valid_date v_insp_from_date,'Inspection Date From','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_insp_from_date',v_form.v_insp_from_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		</TD>
		<TD CLASS="tableheader">Expiry Date From</TD>
		<TD CLASS="tabledata">&nbsp;
			<INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_expr_from_date" VALUE="<% =Request("v_expr_from_date") %>" SIZE="12"
				onblur="vbscript:valid_date v_expr_from_date,'Expiry Date From','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_expr_from_date',v_form.v_expr_from_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		</TD>
	</TR>
	<TR HEIGHT="20pt">
		<TD CLASS="tableheader">Inspection Date To</TD>
		<TD CLASS="tabledata">&nbsp;
			<INPUT TYPE="text" CLASS="textbox" NAME="v_insp_to_date" STYLE="background-color:white" VALUE="<% =Request("v_insp_to_date") %>" SIZE="12"
				onblur="vbscript:valid_date v_insp_to_date,'Inspection Date To','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_insp_to_date',v_form.v_insp_to_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		</TD>
		<TD CLASS="tableheader">Expiry Date To</TD>
		<TD CLASS="tabledata">&nbsp;
			<INPUT TYPE="text" CLASS="textbox" NAME="v_expr_to_date" STYLE="background-color:white" VALUE="<% =Request("v_expr_to_date") %>" SIZE="12"
				onblur="vbscript:valid_date v_expr_to_date,'Expiry Date To','v_form'">
				<A HREF="javascript:show_calendar('v_form.v_expr_to_date',v_form.v_expr_to_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		</TD>
	</TR>	
</TABLE>
<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="0">
	<TR HEIGHT="10pt"></TR>
	<TR HEIGHT="20pt">
		<TD WIDTH="50%" BGCOLOR="lightblue">&nbsp;
			<SELECT NAME="v_report_name" STYLE="width:185pt;background-color:lightblue" CLASS="ddlist">
				<OPTION VALUE="">Select a Report</OPTION>
				<!--<OPTION VALUE="vessel_inspections_report.asp">Vessel Inspections</OPTION>-->
				<OPTION VALUE="moc_inspections_report.asp">Major Oil Companies Inspections</OPTION>
				<!--<OPTION VALUE="inspection_cost_report.asp">Inspections Cost</OPTION>-->
				<!--<OPTION VALUE="moc_list_of_defs.rpt">List of Deficiencies</OPTION>-->
				<!--<OPTION VALUE="moc_inspection_status.rpt">Status Report</OPTION>-->
				<!--<OPTION VALUE="moc_weekly_forecast.rpt">Weekly Report</OPTION>-->
				<!--<OPTION VALUE="rpt_StatusReport.asp">Status Report</OPTION>-->
				<!--<OPTION VALUE="rpt_WeeklyReport.asp">Weekly Report</OPTION>-->
				<!--<OPTION VALUE="moc_inspections_pending.rpt">Inspection - Pending Report</OPTION>-->
				<!--<OPTION VALUE="moc_office_alert.rpt">Office Alert Report</OPTION>-->
				<!--<OPTION VALUE="moc_tc_status.rpt">TC Status Report</OPTION>-->
				<!--<OPTION VALUE="moc_follow_up.rpt">MOC Follow Up Report</OPTION>-->
				<!--<OPTION VALUE="rpt_moc_followup.asp">MOC Follow Up Report</OPTION>-->
			</SELECT>
			<INPUT TYPE="reset" VALUE="View Report" NAME="v_view_report" CLASS="cmdbutton"
				OnClick="return callReport();">&nbsp;&nbsp;
		</TD>
		<TD WIDTH="50%" ALIGN="right" BGCOLOR="lightblue">
			<INPUT TYPE="reset" VALUE="Clear Filters" NAME="v_clear" CLASS="cmdbutton"
				OnClick="return clearFilters();" TITLE="Clear Filters">&nbsp;
		</TD>
	</TR>
	<TR HEIGHT="10pt"></TR>
</TABLE>
</FORM>
</BODY>
</HTML>
