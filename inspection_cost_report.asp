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

		If val = "0" Or CDbl(val) = 0 Then
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
		v_insp_date_condition = "and mir.inspection_date between to_date('" & Request("v_insp_from_date") & "', 'dd-Mon-yyyy') and to_date('" & Request("v_insp_to_date") & "', 'dd-Mon-yyyy') "
	End If

	strSql = "select mir.moc_id, count(mir.moc_id) cnt, min(mm.short_name) short_name, "
	strSql = strSql & "sum(nvl(expences_in_usd, 0.0)) cost_in_usd "
	strSql = strSql & "from moc_inspection_requests mir, moc_master mm "
	strSql = strSql & "where mir.moc_id = mm.moc_id "
	strSql = strSql & v_insp_type_condition
	strSql = strSql & v_insp_date_condition
	strSql = strSql & "group by mir.moc_id "
	strSql = strSql & "order by 2"

	'Response.Write strSql & "<BR>"
	'Response.End
	Set rsObj = connObj.Execute(strSql)
	
	v_row_count = 1
	v_cost_total = 0

	If rsObj.EOF = False Then	'if there are records
%>
		<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="1" BORDERCOLOR="lightgrey">
			<TR HEIGHT="30pt">
				<TD WIDTH="75%" CLASS="reporttableheading" ALIGN="center">MOC</TD>
				<TD WIDTH="25%" CLASS="reporttableheading" ALIGN="center">Cost USD</TD>
			</TR>
<%
		While Not rsObj.EOF
%>
			<TR HEIGHT="20pt">
				<TD CLASS="reporttabledata">&nbsp;<% =rsObj("short_name") %></TD>
				<TD CLASS="reporttabledata" ALIGN="right"><% =blankIfZero(FormatNumber(rsObj("cost_in_usd"), 2)) %>&nbsp;</TD>
			</TR>
<%
			v_row_count = v_row_count + 1
			v_cost_total = v_cost_total + CDbl(rsObj("cost_in_usd"))
			rsObj.MoveNext
		Wend
%>
			<TR HEIGHT="30pt" VALIGN="bottom">
				<TD CLASS="reporttableheading">&nbsp;Total</TD>
				<TD CLASS="reporttableheading" ALIGN="right"><% =FormatNumber(v_cost_total, 2) %>&nbsp;</TD>
			</TR>
		</TABLE>
<%
	Else	'if there are no records
		Response.Write "<SPAN CLASS='reportsubheading'>No matching records !</SPAN><BR>"
	End If

	rsObj.Close
	Set rsObj = Nothing
%>
</BODY>
</HTML>
