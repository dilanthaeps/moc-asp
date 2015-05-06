<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<%
Response.Expires=-1
dim SQL,rs,rsVessels,FLEET,SDATE1,SDATE2
dim sTemp,sColor,dt
FLEET = Request.QueryString("FLEET")
SDATE1 = Request.QueryString("SDATE1")
SDATE2 = Request.QueryString("SDATE2")
if SDATE1="" then
	dt = cdate("1 " & MonthName(month(now),true) & " " & year(now))
	SDATE2 = FormatDateTimeValues(dt-1,2)
	SDATE1 = FormatDateTimeValues(DateAdd("m",-1,dt),2)
end if
set rs=Server.CreateObject("ADODB.Recordset")
with rs
	.CursorLocation=adUseClient
	.CursorType=adOpenStatic
	.LockType=adLockReadonly
end with
SQL = " SELECT ir.request_id, ir.vessel_code, v.vessel_name,mm.short_name moc_name,inspection_grade,count(md.request_id) count,"
 SQL = SQL &  "           (CASE WHEN insp_type = 'MOC' AND basis_sire IS NULL THEN 1 ELSE 0 END) moc,"
 SQL = SQL &  "           (CASE WHEN insp_type = 'MOC' AND basis_sire IS NULL AND insp_status = 'FAILED' THEN 1 ELSE 0 END) moc_failed,"
 SQL = SQL &  "           (CASE WHEN insp_type = 'MOC' AND basis_sire IS NOT NULL THEN 1 ELSE 0 END) moc_based_sire,"
 SQL = SQL &  "          (CASE WHEN insp_type = 'PSC' THEN 1 ELSE 0 END) psc,"
 SQL = SQL &  " 			(CASE WHEN insp_type = 'TMNL' THEN 1 ELSE 0 END) tmnl,"
 SQL = SQL &  "          (CASE WHEN insp_type = 'TVEL' THEN 1 ELSE 0 END) uscg_tvel"
 SQL = SQL &  "  FROM moc_inspection_requests ir, wls_vw_vessels_new v, moc_master mm,moc_deficiencies md"
 SQL = SQL &  "  WHERE ir.vessel_code = v.vessel_code"
 SQL = SQL &  "      AND ir.moc_id=mm.moc_id(+)"
 SQL = SQL &  " 	  and ir.request_id=md.request_id(+)"
 SQL = SQL &  "      AND insp_status IN"
 SQL = SQL &  "             ('INSPECTED',"
 SQL = SQL &  "             'ACCEPTED',"
 SQL = SQL &  " 			 'FAILED',"
 SQL = SQL &  "              'PENDING BASED SIRE',"
 SQL = SQL &  "              'REPLIED BASED SIRE',"
 SQL = SQL &  "              'ACCEPTED BASED SIRE',"
 SQL = SQL &  "              'REPORT RECEIVED',"
 SQL = SQL &  "              'SIRE REPORT RECEIVED',"
 SQL = SQL &  "              'REPORT REPLIED',"
 SQL = SQL &  "             'SIRE REPORT REPLIED'"
 SQL = SQL &  "            )"
 SQL = SQL &  "      AND trunc(inspection_date) BETWEEN '" & SDATE1 & "' AND '" & SDATE2 & "'"
 if FLEET<>"" then
    SQL = SQL &  " 	  and fleet_code='" & FLEET & "'"
 end if
 SQL = SQL &  " 	  group by ir.request_id, ir.vessel_code, v.vessel_name,mm.short_name ,inspection_grade,insp_type,basis_sire,insp_status"
 SQL = SQL &  "  ORDER BY vessel_name"

rs.Open SQL,connObj

SQL = "Select vessel_code,vessel_name from wls_vw_vessels_new"
if FLEET<>"" then
	SQL = SQL & " where fleet_code='" & FLEET & "'"
end if
set rsVessels=connObj.execute(SQL)
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" HREF="moc.css"></link>
<style>
TD{vertical-align:top;}
</style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	dim i,j,rowTot,colTot(5),grandTot,nValue
	
	frm1.cmbFleet.value = "<%=FLEET%>"
	
	grandTot=0
	for i=2 to tabData.rows.length-2
		rowTot = 0
		for j=1 to tabData.rows(2).cells.length-2
			if trim(tabData.rows(i).cells(j).getAttribute("nval"))<>"" then
				nValue = cint(trim(tabData.rows(i).cells(j).getAttribute("nval")))
				rowTot = rowTot + nValue
				colTot(j-1) = colTot(j-1) + nValue
			end if
		next
		tabData.rows(i).cells(j).innerText = rowTot
		grandTot = grandTot + rowTot
	next
	for j=1 to tabData.rows(2).cells.length-2
		tabData.rows(i).cells(j).innerText = colTot(j-1)
	next
	tabData.rows(i).cells(j).innerText = grandTot
End Sub

sub cmbFleet_onchange
	dim sUrl
	sd1=frm1.v_insp_from_date.value
	sd2=frm1.v_insp_to_date.value
	if IsDate(sd1) then
		s1 = day(sd1) & " " & monthname(month(sd1),true) & " " & year(sd1)
	else
		MsgBox "Please enter a valid from date",vbInformation,"MOC"
		frm1.v_insp_from_date.focus
		exit sub
	end if
	if IsDate(sd2) then
		s2 = day(sd2) & " " & monthname(month(sd2),true) & " " & year(sd2)
	else
		MsgBox "Please enter a valid to date",vbInformation,"MOC"
		frm1.v_insp_to_date.focus
		exit sub
	end if
	sUrl = "rpt_ScoreCard.asp?FLEET=" & frm1.cmbFleet.value & "&SDATE1=" & s1 & "&SDATE2=" & s2
	window.location.href = sUrl
end sub

sub Hilite(tr)
	tr.style.backgroundColor = "lightgreen"
end sub

sub RemoveHilite(tr)
	tr.style.backgroundColor = ""
end sub
Sub window_onbeforeprint
	trHide.style.display="none"
	divTopMenu.style.display="none"
	Hide1.style.display=""
	Hide2.style.display="none"
End Sub

Sub window_onafterprint
	trHide.style.display=""
	divTopMenu.style.display=""
	Hide1.style.display="none"
	Hide2.style.display=""
End Sub
-->
</SCRIPT>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>
</HEAD>
<BODY>
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<center>
<form name=frm1 style="margin-bottom:0">
<table cellpadding=2 cellspacing=0 border=0 width=100%>
  <tr>
    <td rowspan=3 id=Hide2>
      <table style="border:1px solid black;">
        <tr bgcolor="khaki">
		  <td style="font-size:9px"><font color=red><b>       
                1-Unacceptable<br>
                2-Poor<br>
                3-Average<br>
                4-Good<br>
                5-Excellent</b></font></td></tr>
      </table>      
    </td> 
    <td align=center colspan=4><h3 style="margin-bottom:0">Major Oil Companies - Inspection Summary
    <%if FLEET<>"" then%>
    - <%=FLEET%>
    <%end if%>
    <td rowspan=3><table style="border:1px solid black;visibility:hidden">
          <tr  bgcolor="khaki">
		    <td style="font-size:9px"><font color=red><b>     
                1-Unacceptable<br>
                2-Poor<br>
                3-Average<br>
                4-Good<br>
                5-Excellent</b></font></td></tr>
      </table>
    </td>
  <tr id=Hide1 style="font-size:11px;display:none" align=center>
    <td> <%=SDATE1%> To <%=SDATE2%> For
    <%if FLEET="" then 
        Response.Write("All Fleets")
     else %>
    <%=FLEET%> 
    <%end if%></td>
    <td>
  </tr>   
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
      <select name=cmbFleet>
        <option value="<%%>">All Fleets
        <option value="AFRAMAX">AFRAMAX
        <option value="FSO">FSO
        <option value="PRODUCT">PRODUCT
        <option value="SUEZMAX">SUEZMAX
        <option value="VLCC">VLCC
      </select>
    <td><br><button onclick=cmbFleet_onchange>Refresh</button>
</table>
</form>
<br>
<table id=tabData border=0 bgcolor=lightgrey cellspacing=1>
  <tr class=tableheader style="font-size:11px">
    <td rowspan=2>Vessel
    <td colspan=2 align=center>MOC 
    <td rowspan=2>SIRE referral
    <td rowspan=2>Port State Control
    <td rowspan=2>Terminal
    <td rowspan=2>USCG-TVEL
    <td rowspan=2 class=tot>Total
  </tr>
  <tr class=tableheader style="font-size:11px">
    <td>Cleared
    <td>Failed
  </tr>
  <%
  while not rsVessels.eof%>
  <tr bgcolor="<%=ToggleColor(sColor)%>" onmouseover="Hilite(me)" onmouseout="RemoveHilite(me)">
     <td nowrap><a title="Click to open vessel particulars" href='http://webserve2/vid/create_data_file.asp?vessel_code=<%=rsVessels("vessel_code")%>&questionnaire_id=10000380' target='vessel_particulars'><span style='text-align:center;font-size:9;color:blue'><%=rsVessels("vessel_name")%></span></a></td>   
    <%
    sTemp="&nbsp;"
    rs.filter="moc=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>MOC (<b>" & rs.RecordCount & "</b>)"
    sTemp = sTemp & "	 <td style='font-size:9px'>Grade</tr>"
     sTemp = sTemp & "	 <td style='font-size:9px'>Obs</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td colspan=3>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      if clng(rs("count"))<>0 then      
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'><a href='ins_request_def_maint.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open list of observations'>" & rs("count")& "</a>"
      else
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("count")
      end if  
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <%
    sTemp="&nbsp;"
    rs.filter="moc_failed=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>MOC"
    sTemp = sTemp & "<td style='font-size:9px'>Grade</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td colspan=2>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      'sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <%
    sTemp="&nbsp;"
    rs.filter="moc_based_sire=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>Referral (<b>" & rs.RecordCount & "</b>)"
    'sTemp = sTemp & "<td style='font-size:9px'>Grade</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      'sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <%
    sTemp="&nbsp;"
    rs.filter="psc=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>PSC (<b>" & rs.RecordCount & "</b>)"
    'sTemp = sTemp & "<td style='font-size:9px'>Grade</tr>"
    sTemp = sTemp & "<td style='font-size:9px'>Obs</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td colspan=2>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      'sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      if clng(rs("count"))<>0 then      
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'><a href='ins_request_def_maint.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open list of observations'>" & rs("count")& "</a>"
      else
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("count")
      end if  
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <%
    sTemp="&nbsp;"
    rs.filter="tmnl=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>Terminal (<b>" & rs.RecordCount & "</b>)"
    'sTemp = sTemp & "<td style='font-size:9px'>Grade</tr>"
    sTemp = sTemp & "<td style='font-size:9px'>Obs</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td colspan=2>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      'sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      if clng(rs("count"))<>0 then      
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'><a href='ins_request_def_maint.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open list of observations'>" & rs("count")& "</a>"
      else
        sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("count")
      end if  
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <%
    sTemp="&nbsp;"
    rs.filter="uscg_tvel=1 and vessel_code='" & rsVessels("vessel_code") & "'"
    if not rs.EOF then
    'sTemp = "<b>Total: " & rs.recordcount & "</b><br>"
    sTemp = "<table border=0 width='100%'>"
    sTemp = sTemp & "<tr><td style='font-size:9px'>TVEL (<b>" & rs.RecordCount & "</b>)"
    'sTemp = sTemp & "<td style='font-size:9px'>Grade</tr>"
    sTemp = sTemp & "<tr bgcolor=lightgrey><td>"
    while not rs.eof
      sTemp = sTemp & "<tr><td style='font-size:9px'><a href='ins_request_entry.asp?v_ins_request_id=" & rs("request_id") & "' target='insp_details' title='Click to open inspection report'>" & rs("moc_name") & "</a>"
      'sTemp = sTemp & "<td style='font-size:9px;text-align:center'>" & rs("inspection_grade")
      rs.movenext
    wend
    sTemp = sTemp & "</table>"
    end if%>
    <td nowrap nVal="<%=rs.recordcount%>"><%=sTemp%>
    <td class=tot>
  <%rsVessels.movenext
  wend%>
  <tr class=tot bgcolor="<%=ToggleColor(sColor)%>">
    <td>Total Insp
    <td>&nbsp;
    <td>&nbsp;
    <td>&nbsp;
    <td>&nbsp;
    <td>&nbsp;
    <td>&nbsp;
    <td>&nbsp;
</table>
</center>
</BODY>
</HTML>
