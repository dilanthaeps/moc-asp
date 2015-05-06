<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	
'===========================================================================
'	Template Name	:	MOC Inspection Followup report
'	Template Path	:	rpt_moc_followup.asp
'	Functionality	:	MOC inspection followup report
'	Called By		:	report_filter_panel.asp
'	Created By		:	Prashant Kumar
'   Create Date		:	18th  February, 2006
'	Update History	:
'						1.
'						2.
'===========================================================================
Response.Buffer = true

dim SQL,rs
dim VID,FLEET,MOC,INSP_DATE_FROM,INSP_DATE_TO,INSP_TYPE,INSP_STATUS
dim REMARK_STATUS

VID = Request.QueryString("VID")
if Request.QueryString("ACTION")="VALUE" then
  FLEET=Request.QueryString("v_fleet")
else
  FLEET = Request.QueryString("cmbFleet")  
end if
if REMARK_STATUS="" then
	REMARK_STATUS = "ACTIVE"
end if

SQL = " SELECT mir.request_id, mir.vessel_code, v.vessel_name, v.fleet_code,"
SQL = SQL &  "        TO_CHAR (mir.moc_id) moc_id,"
SQL = SQL &  "        SUBSTR (moc_fn_moc_short_name (mir.moc_id), 1, 255) moc_name,"
SQL = SQL &  "        mir.inspection_date, TO_CHAR (mir.inspection_date, 'dd-Mon-yyyy')inspection_date,"
SQL = SQL &  "        mir.inspection_port,"
SQL = SQL &  "        SUBSTR (moc_fn_sys_para_desc (mir.insp_status, 'Status'),"
SQL = SQL &  "                1,"
SQL = SQL &  "                255"
SQL = SQL &  "               ) insp_status_name,"
SQL = SQL &  "        SUBSTR (moc_fn_sys_para_desc (mir.insp_type, 'Inspection_Type'),"
SQL = SQL &  "                1,"
SQL = SQL &  "                255"
SQL = SQL &  "               ) insp_type_name,"
SQL = SQL &  "        SUBSTR (moc_fn_basis_sire_short_name (mir.basis_sire),"
SQL = SQL &  "                1,"
SQL = SQL &  "                255"
SQL = SQL &  "               ) basis_sire_name,"
SQL = SQL &  "        mir.insp_type, mrr.subject, mrr.remarks,"
SQL = SQL &  "        SUBSTR (wls_fn_user_name (mrr.remark_pic), 1, 255) remark_pic_name,"
SQL = SQL &  "        mrr.remark_target_date,"
SQL = SQL &  "        TO_CHAR (mrr.remark_target_date, 'dd-Mon-yyyy')remark_target_date"
SQL = SQL &  " FROM moc_inspection_requests mir,"
SQL = SQL &  " 	wls_vw_vessels_new v,"
SQL = SQL &  " 	moc_request_remarks mrr"
SQL = SQL &  " WHERE mir.request_id = mrr.request_id"
SQL = SQL &  " AND mir.vessel_code = v.vessel_code"
SQL = SQL &  " AND UPPER (TRIM (mrr.remark_status)) ='ACTIVE'"

if FLEET<>"" then
	SQL = SQL & " and v.fleet_code='" & FLEET & "'"
end if

set rs = connObj.execute(SQL)
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="vbscript">
  Sub window_onbeforeprint
	divTopMenu.style.display="none"	
	Hide.style.display="none"	
End Sub

Sub window_onafterprint
	divTopMenu.style.display=""	
	Hide.style.display=""	
End Sub

sub cmbFleet_onchange() 
   dim sUrl,str
   str=frm1.cmbFleet.value   
 sUrl = "rpt_moc_followup.asp?ACTION=VALUE&v_fleet=" & str 
	window.location.href = sUrl	
end sub
</script>
<link REL="stylesheet" HREF="moc.css"></link>
</HEAD>
<BODY bgcolor=white>
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<form name="frm1">
<table width=100% align=center border=0>
  <tr>
    <td id="Hide"><font color=blue><b>FLEET</b></font>
    &nbsp;&nbsp;<select name=cmbFleet class="menuHide">
        <option value="">All Fleets
        <option value="AFRAMAX" <%if FLEET="AFRAMAX" then Response.Write("selected")%>>AFRAMAX
        <option value="FSO" <%if FLEET="FSO" then Response.Write("selected")%>>FSO
        <option value="PRODUCT" <%if FLEET="PRODUCT" then Response.Write("selected")%>>PRODUCT
        <option value="SUEZMAX" <%if FLEET="SUEZMAX" then Response.Write("selected")%>>SUEZMAX
        <option value="VLCC" <%if FLEET="VLCC" then Response.Write("selected")%>>VLCC
      </select></td>
    <td width="40%"><h3>MOC Follow-up Report</h3>
    <td class=num nowrap><%=FormatDateTimeValues(now,2)%> 
  <tr><td>&nbsp;</td><td><blockquote><blockquote><h4 style="margin:0"><b>
      <%if FLEET="" then 
          response.write("All Fleets")
        else%>
          <%=FLEET%>
        <%end if%>   <b></h4></blockquote><blockquote></td></tr>         
</table>
<br>
<table width=100%>
<%
  if  rs.eof then
    Response.Write("<b>No Data Found</b>")
  else
		while not rs.eof%>
		  <tr>
		    <td style="font-size:15px;font-weight:bold"><%=rs("vessel_name")%> <font size=-2>(<%=rs("insp_type_name")%>)</font>
		  <tr style="font-weight:bold">
		    <td><%=rs("moc_name")%>
		    <td colspan=2 class=txt><%=rs("inspection_date")%> - <%=rs("inspection_port")%>
		    <td><%=rs("insp_status_name")%>
		  <tr bgcolor=lightgrey><td colspan=4>
		  <tr>
		    <td colspan=2><b>Responsibility:</b> <%=rs("remark_pic_name")%>
		    <td class=num><b>Target Date:</b>
		    <td><%=rs("remark_target_date")%>
		  <tr><td>&nbsp;
		  <tr>
		    <td colspan=4><b><%=rs("subject")%></b>
		  <tr>
		    <td colspan=4 style="border:1px solid lightgrey;"><%=ToHTML(rs("remarks"))%><br>
		  <tr>
		    <td colspan=4>
		<hr color=gray>
		<%
			rs.movenext
		wend
  end if%>
</table>
</form>
</BODY>
</HTML>
