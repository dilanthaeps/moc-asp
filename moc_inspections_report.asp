<%@ Language=VBScript %>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<%
	Response.Buffer = False
	v_tab_width = "70%"		
	
	Function IIF(expr, trueValue, falseValue)

		If expr Then
			IIF = trueValue
		Else
			IIF = falseValue
		End If

	End Function  
	
	 dim v_FLEET,SDATE1,SDATE2
	 
	 if Request.QueryString("ACTION")="VALUE" then
	   v_FLEET=Request.QueryString("FLEET")
	   SDATE1 = Request.QueryString("SDATE1")
       SDATE2 = Request.QueryString("SDATE2")	   
	 else 
       v_FLEET =Request.QueryString("cmbFleet") 
       SDATE1 = Request.QueryString("v_insp_from_date")
       SDATE2 = Request.QueryString("v_insp_to_date")           
    end if     
    
   if SDATE1="" then
	   dt = cdate("1 " & MonthName(month(now),true) & " " & year(now))
	   SDATE2 = FormatDateTimeValues(dt-1,2)
	   SDATE1 = FormatDateTimeValues(DateAdd("m",-1,dt),2)
    end if
       
   strSql = "select mir.moc_id, min(mm.short_name) short_name, "
   strSql = strSql & " sum(MOC_FN_MONTH_INSP_COUNT(mir.insp_status)) insp_mon_count, "
   strSql = strSql & " sum(MOC_FN_MONTH_SIRE_COUNT(mir.insp_status)) sire_mon_count "
   strSql = strSql & " from moc_inspection_requests mir, moc_master mm ,wls_vw_vessels_new wvn "
   strSql = strSql & " where mir.moc_id = mm.moc_id "
   strSql = strSql & " and mir.vessel_code=wvn.vessel_code"
   strSql = strSql & " and mir.insp_type = 'MOC'"
   strSql = strSql & " and  mir.inspection_date BETWEEN '" & SDATE1 & "' AND '" & SDATE2 & "'" 
   
    if v_FLEET<>"" then
        strSql = strSql & " and wvn.fleet_code='" & v_FLEET & "'"
    end if
    strSql = strSql & " group by mir.moc_id,short_name order by short_name" 

   Set rsObj = connObj.Execute(strSql)	
%>
<HTML>
<HEAD>
<link REL="stylesheet" HREF="moc.css"></link>
<style>
.clsIncident
{
	font-family:webdings;
	font-size:15px;
	color:red;
	cursor:default;
}
</style>
<script language="Javascript" src="js_date.js"></script>
<script language="Javascript" src="AutoComplete.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<script language="vbscript">	
sub cmbFleet_onchange()   
   dim sUrl,str
    str=frm1.cmbFleet.value     
	sd1=frm1.v_insp_from_date.value
	sd2=frm1.v_insp_to_date.value
	if IsDate(sd1) then
		s1 = day(sd1) & "-" & monthname(month(sd1),true) & "-" & year(sd1)
	else
		MsgBox "Please enter a valid from date",vbInformation,"MOC"
		frm1.v_insp_from_date.focus
		exit sub
	end if
	if IsDate(sd2) then
		s2 = day(sd2) & "-" & monthname(month(sd2),true) & "-" & year(sd2)
	else
		MsgBox "Please enter a valid to date",vbInformation,"MOC"
		frm1.v_insp_to_date.focus
		exit sub
	end if
	sUrl = "moc_inspections_report.asp?ACTION=VALUE&FLEET=" & str & "&SDATE1=" & s1 & "&SDATE2=" & s2
	window.location.href = sUrl	
end sub
Sub window_onbeforeprint
    trHide.style.display="none"	
	divTopMenu.style.display="none"
	'Hide.style.display=""
	Hide1.style.display=""
End Sub

Sub window_onafterprint
	trHide.style.display=""
	divTopMenu.style.display=""
	'Hide.style.display="none"
	Hide1.style.display="none"
End Sub

</script>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<body>
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
</HEAD>
<form name=frm1 style="margin-bottom:0">
<table cellpadding=2 cellspacing=0 border=0 align=center>
  <tr>
    <td align=center colspan=4><h3 style="margin-bottom:0">Major Oil Companies Inspections </h3>
    <tr id=Hide1 style="font-size:11px;display:none" align=center>
    <td> <%=SDATE1%> To <%=SDATE2%> For
    <%if v_FLEET="" then 
        Response.Write("All Fleets")
     else %>
    <%=v_FLEET%> 
    <%end if%></td>
    <td>
  </tr>    
   <tr><td>&nbsp;</td></tr>   
  <tr id=trHide>
    <td nowrap><b>Date from</b><br>
      <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<%=SDATE1%>" SIZE="12"
				onblur="vbscript:checkDate frm1.v_insp_from_date,'Inspection Date From','frm1'">
				<A HREF="javascript:show_calendar('frm1.v_insp_from_date',frm1.v_insp_from_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
	<td nowrap><b>Date to</b><br>
	  <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_to_date" VALUE="<%=SDATE2%>" SIZE="12"
				onblur="vbscript:checkDate frm1.v_insp_to_date,'Inspection Date To','frm1'">
				<A HREF="javascript:show_calendar('frm1.v_insp_to_date',frm1.v_insp_to_date.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>

    <td><b>Fleet</b><br>
      <select name="cmbFleet" id="cmbFleet" >
			<option value="">All Fleets
			<option value="AFRAMAX" <%if cstr(v_FLEET)="AFRAMAX" then Response.Write("selected")%>>AFRAMAX
			<option value="FSO" <%if cstr(v_FLEET)="FSO" then Response.Write("selected")%>>FSO
			<option value="PRODUCT" <%if cstr(v_FLEET)="PRODUCT" then Response.Write("selected")%>>PRODUCT
			<option value="SUEZMAX" <%if cstr(v_FLEET)="SUEZMAX" then Response.Write("selected")%>>SUEZMAX
			<option value="VLCC" <%if cstr(v_FLEET)="VLCC" then Response.Write("selected")%>>VLCC
      </select>
    <td>
    <td> <input type="button" name="but1" id="but1" value="Refresh" onclick="cmbFleet_onchange()"></td></tr>
</table><br>
</form>

<% 
	v_insp_mon_total = 0	
	v_sire_mon_total = 0
	If rsObj.EOF = False Then	'if there are records
%>
		<TABLE WIDTH="<% =v_tab_width %>" CELLPADDING="0" CELLSPACING="0" BORDER="1" BORDERCOLOR="lightgrey" align=center>
			<TR class=tableheader style="font-size:11px">
				<TD>MOC</TD>
				<TD align=center>Inspections</TD>
				<TD align=center>Based on SIRE</TD>
			</TR>		
<%
		While Not rsObj.EOF
           if  clng(rsObj("insp_mon_count"))<>0 or clng(rsObj("sire_mon_count"))<>0 then%>
			<TR HEIGHT="20pt">
				<TD CLASS="reporttabledata">&nbsp;<b><% =rsObj("short_name")%></b></TD>
				<TD CLASS="reporttabledata" ALIGN="center"><%=rsObj("insp_mon_count")%></TD>			  
				<TD CLASS="reporttabledata" ALIGN="center"><%=rsObj("sire_mon_count")%></TD>
			  	
			</TR>			
<%           end if			
			v_insp_mon_total = v_insp_mon_total + CDbl(rsObj("insp_mon_count"))			
			v_sire_mon_total = v_sire_mon_total + CDbl(rsObj("sire_mon_count"))			
			rsObj.MoveNext
		Wend%>
		<TR HEIGHT="30pt" VALIGN="bottom" bgcolor="Lightsteelblue">
				<TD CLASS="reporttableheading">&nbsp;<font color="red">TOTAL</font></TD>
				<TD CLASS="reporttableheading" ALIGN="center"><font color="red"><% =v_insp_mon_total %></font></TD>				
				<TD CLASS="reporttableheading" ALIGN="center"><font color="red"><% =v_sire_mon_total %></font></TD>				
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
