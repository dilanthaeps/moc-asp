<%@ Language=VBScript %>
<!--#include file="common_dbconn.asp"-->
<%
	v_tab_width = "60%"
	'Response.Write "v_insp_type - " & Request("v_insp_type") & "<BR>"
	'Response.Write "v_insp_type - " & Request("v_insp_from_date") & "<BR>"
	'Response.Write "v_insp_type - " & Request("v_insp_to_date") & "<BR>"

	Function IIF(expr, trueValue, falseValue)

		If expr Then
			IIF = trueValue
		Else
			IIF = falseValue
		End If

	End Function

	Function blankIfZero(val)

		If val = "0" Then
			blankIfZero = "&nbsp;"
		Else
			blankIfZero = val
		End If

	End Function

	strSql = "select para_desc "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where upper(trim(parent_id)) = 'INSPECTION_TYPE' "
	strSql = strSql & "and upper(trim(sys_para_id)) = '" & UCase(Trim(Request("v_insp_type"))) & "' "

	'Response.Write strSql & "<BR>"
	Set rsObj = connObj.Execute(strSql)
	
	v_insp_type_name = ""
	If rsObj.EOF = False Then
		v_insp_type_name = rsObj("para_desc")
	End If

	rsObj.Close
	Set rsObj = Nothing
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Vessel Inspections Report</TITLE>
</HEAD>
<BODY>
<P>
<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="0">
	<TR HEIGHT="70pt">
		<TD WIDTH="95%" CLASS="reportheading">Vessel Inspections</TD>
		<TD WIDTH="5%" ALIGN="right">
			<INPUT TYPE="button" NAME="v_close" CLASS="cmdbutton" VALUE="Close" OnClick="self.close();">
		</TD>
	</TR>
	<TR HEIGHT="30pt">
		<TD COLSPAN="2" CLASS="reportsubheading"><% =IIF(Request("v_insp_type") <> "", "Inspection Type: " & v_insp_type_name, "") %></TD>
	</TR>
	<TR>
		<TD COLSPAN="2" CLASS="reportsubsubheading">
			<% =IIF(Request("v_insp_from_date") <> "" And Request("v_insp_from_date") <> "", "Period From: " & Request("v_insp_from_date") & " To: " & Request("v_insp_to_date"), "") %>
		</TD>
	</TR>
	<TR HEIGHT="30pt"></TR>
</TABLE>
<%
	v_insp_type_condition = ""
	If Request("v_insp_type") <> "" Then
		v_insp_type_condition = "and upper(trim(mir.insp_type)) = '" & UCase(Trim(Request("v_insp_type"))) & "' "
	End If

	v_insp_date_condition = ""
	If Request("v_insp_from_date") <> "" And Request("v_insp_to_date") <> "" Then
		v_insp_date_condition = "and decode(upper(trim(mir.insp_status)), 'ACCEPTED BASED SIRE', mir.date_accepted, mir.inspection_date) between to_date('" & Request("v_insp_from_date") & "', 'dd-Mon-yyyy') and to_date('" & Request("v_insp_to_date") & "', 'dd-Mon-yyyy') "
	End If

	strSql = "select mir.moc_id, min(mm.short_name) short_name, "
	strSql = strSql & "sum(decode(upper(trim(mir.insp_status)), 'ACCEPTED BASED SIRE', 0, 1)) insp_count, "
	strSql = strSql & "sum(decode(upper(trim(mir.insp_status)), 'ACCEPTED BASED SIRE', 1, 0)) sire_count "
	strSql = strSql & "from moc_inspection_requests mir, moc_master mm "
	strSql = strSql & "where mir.moc_id = mm.moc_id "
	strSql = strSql & "and upper(trim(mir.insp_status)) in (select upper(trim(sys_para_id)) from moc_system_parameters "
	strSql = strSql & "where upper(trim(parent_id)) = 'VESS INSP REP STATUS') "
	strSql = strSql & v_insp_type_condition
	strSql = strSql & v_insp_date_condition
	strSql = strSql & "group by mir.moc_id "
	strSql = strSql & "order by 2"

	'Response.Write strSql & "<BR>"
	'Response.End
	Set rsObj = connObj.Execute(strSql)
	
	v_row_count = 1
	v_insp_total = 0
	v_sire_total = 0

	If UCase(Trim(Request("v_insp_type"))) = "MOC" Then

		If rsObj.EOF = False Then	'if there are records
%>
			<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="1" BORDERCOLOR="lightgrey">
				<TR HEIGHT="30pt">
					<TD WIDTH="50%" CLASS="reporttableheading" ALIGN="center">MOC</TD>
					<TD WIDTH="25%" CLASS="reporttableheading" ALIGN="center">Inspections</TD>
					<TD WIDTH="25%" CLASS="reporttableheading" ALIGN="center">Based on SIRE</TD>
				</TR>
<%
			While Not rsObj.EOF
%>
				<TR HEIGHT="20pt">
					<TD CLASS="reporttabledata">&nbsp;<% =rsObj("short_name") %></TD>
					<TD CLASS="reporttabledata" ALIGN="center"><% =blankIfZero(rsObj("insp_count")) %></TD>
					<TD CLASS="reporttabledata" ALIGN="center"><% =blankIfZero(rsObj("sire_count")) %></TD>
				</TR>
<%
				v_row_count = v_row_count + 1
				v_insp_total = v_insp_total + CDbl(rsObj("insp_count"))
				v_sire_total = v_sire_total + CDbl(rsObj("sire_count"))
				rsObj.MoveNext
			Wend
%>
				<TR HEIGHT="30pt" VALIGN="bottom">
					<TD CLASS="reporttableheading">&nbsp;Total</TD>
					<TD CLASS="reporttableheading" ALIGN="center"><% =v_insp_total %></TD>
					<TD CLASS="reporttableheading" ALIGN="center"><% =v_sire_total %></TD>
				</TR>
			</TABLE>
<%
		Else	'if there are no records
			Response.Write "<SPAN CLASS='reportsubheading'>No matching records !</SPAN><BR>"
		End If

	Else	'UCase(Trim(Request("v_insp_type"))) = "MOC"

		If rsObj.EOF = False Then	'if there are records
%>
			<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="1" BORDERCOLOR="lightgrey">
				<TR HEIGHT="30pt">
					<TD WIDTH="50%" CLASS="reporttableheading" ALIGN="center">MOC</TD>
					<TD WIDTH="25%" CLASS="reporttableheading" ALIGN="center">Inspections</TD>
				</TR>
<%
			While Not rsObj.EOF
%>
				<TR HEIGHT="20pt">
					<TD CLASS="reporttabledata">&nbsp;<% =rsObj("short_name") %></TD>
					<TD CLASS="reporttabledata" ALIGN="center"><% =rsObj("insp_count") %></TD>
				</TR>
<%
				v_row_count = v_row_count + 1
				v_insp_total = v_insp_total + CDbl(rsObj("insp_count"))
				rsObj.MoveNext
			Wend
%>
				<TR HEIGHT="30pt" VALIGN="bottom">
					<TD CLASS="reporttableheading">&nbsp;Total</TD>
					<TD CLASS="reporttableheading" ALIGN="center"><% =v_insp_total %></TD>
				</TR>
			</TABLE>
<%
		Else	'if there are no records
			Response.Write "<SPAN CLASS='reportsubheading'>No matching records !</SPAN><BR>"
		End If

	End If	'UCase(Trim(Request("v_insp_type"))) = "MOC"

	rsObj.Close
	Set rsObj = Nothing
%>
</BODY>
</HTML>
