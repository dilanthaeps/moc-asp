<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<html>
<head>
 <meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
 <link REL="stylesheet" HREF="moc.css">
 <style> 
TD
{
	 font-size:10px;
}
.clsVessel
{
	font-size:12px;
	font-weight:bold;
	
}
.clsExpired
{
	color:red;
	font-weight:bold;
}
.clsHighlighted
{
	color:red;
	font-weight:bold;
}
.clsBasisMOC
{
	color:darkblue;
	font-size:10px;
	font-weight:bold;
}
</style>
<SCRIPT LANGUAGE="Javascript" SRC="js_date.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="vb_date.vs"></SCRIPT>
<script language="vbscript">
sub MOC_change()
     if frm1.cmbselect.value="MOC" then			    
	 	 window.location.href="rpt_observations_MOC.asp"		   
	 elseif frm1.cmbselect.value="FLEET" then		 
	     window.location.href="rpt_observations_fleet.asp"
	 end if       	
end sub    
Sub window_onbeforeprint
	divTopMenu.style.display = "none"
	Hide.style.display = "none"	
	hide1.style.display =""	
End Sub
Sub window_onafterprint
	divTopMenu.style.display = ""
	Hide.style.display = ""
	hide1.style.display ="none"
End Sub
sub RefreshPage()
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
	sUrl = "rpt_observations_fleet.asp?SDATE1=" & s1 & "&SDATE2=" & s2
	window.location.href = sUrl
end sub
</script>
</head>
<% 
  dim SDATE1,SDATE2,dt,vessel_code
  SDATE1 = Request.QueryString("SDATE1")
  SDATE2 = Request.QueryString("SDATE2")
  if SDATE1="" then
	dt = cdate("1 " & MonthName(month(now),true) & " " & year(now))
	SDATE2 = FormatDateTimeValues(dt-1,2)
	SDATE1 = FormatDateTimeValues(DateAdd("m",-1,dt),2)
  end if
%>   
<BODY class="bgcolorlogin">
<div id=divTopMenu>
<!--#include file="menu_include.asp"-->
<br>
</div>
<div id="hide1" style="display:none">
     <center><h3 style="margin-bottom:0" id="no_hide">Observations by Fleet - From 1 Jan 2006<br></center>
</div>
<form name="frm1">
	<div id="Hide">    
       <blockquote><blockquote>
      <table><tr><td style="font-size=9pt"><b>Group by</b>
					 <select id="cmbselect" name="cmbselect" width="100px" onchange="MOC_change()" class=menuHide>
					   <option value="MOC">MOC</option>
					   <option value="FLEET" selected>FLEET</option>
					</select>
			    <td style="width:80px">&nbsp;
			    <td colspan=3><h3 style="margin-bottom:0">Observations by Fleet - From 1 Jan 2006</td>
			 <tr><td>&nbsp;</td> <td style="width:80px">&nbsp;
			     <td nowrap align=center><b>Date from</b><br>
			         <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_from_date" VALUE="<%=SDATE1%>" SIZE="12"
						onblur="vbscript:checkDate frm1.v_insp_from_date,'Inspection Date From','frm1'">
						<A HREF="javascript:show_calendar('frm1.v_insp_from_date',frm1.v_insp_from_date.value);">
						<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
		         <td>&nbsp;&nbsp;&nbsp;<b>Date to</b><br>
			         <INPUT TYPE="text" CLASS="textbox" STYLE="background-color:white" NAME="v_insp_to_date" VALUE="<%=SDATE2%>" SIZE="12"
						onblur="vbscript:checkDate frm1.v_insp_to_date,'Inspection Date To','frm1'">
						<A HREF="javascript:show_calendar('frm1.v_insp_to_date',frm1.v_insp_to_date.value);">
						<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0"></A>
				<td><input type=button value="refresh" id=button1 name=button1 onclick=RefreshPage()>		  
       </table></blockquote></blockquote>   
  </div>      	                      
	<%dim fleetcode,rsFleetCode,moc_id
	  strSql="select distinct fleet_code from wls_vw_vessels_new where vessel_code in"
	  strSql=strSql & " (select distinct vessel_code from moc_inspection_requests) order by fleet_code"
	  Set rsFleetCode=connObj.execute(strSql)
	  while not  rsFleetCode.eof 
				   fleetcode=rsFleetCode("fleet_code") %>		   
				   <center><h4 style="margin:0"><b><%=fleetcode%></center><br>
				   <table align=center border=0 cellspacing=1 cellpadding=2 bgcolor=lightgrey width="80%">
				     <tr class=tableheader><td align=center><b>Section</b></td><td align=center><b>Question</b></td>
				       <%dim rsVessel,rsDetails,rsDeficiency
				         dim sect,str2,count2
				         set rsVessel=server.CreateObject("ADODB.Recordset")                 
				         strSql="select distinct mir.vessel_code,vessel_name"
				         strSql=strSql & " from wls_vw_vessels_new wvn,moc_inspection_requests mir"
				         strSql=strSql & " where  mir.vessel_code=wvn.vessel_code"
				         strSql=strSql & " and fleet_code='" & fleetcode & "' order by vessel_name"
				         rsVessel.CursorLocation=adUseClient
				         rsVessel.Open strSql,connObj,adOpenDynamic,adLockReadOnly
				         while not rsVessel.EOF%>
				           <td style="writing-mode:tb-rl;filter:flipv fliph"><b><%=rsVessel("vessel_name")%></b></td>
				           <%
				            rsVessel.MoveNext
				         wend
				         set rsDetails=server.CreateObject("ADODB.Recordset")
				         strSql="select distinct section,question_text,mvq.sort_order "
				         strSql=strSql & " from moc_deficiencies md,moc_viq_questions mvq ,"
				         strSql=strSql & " moc_inspection_requests mir,wls_vw_vessels_new wvn"
				         strSql=strSql & " where md.section=mvq.question_number"
				         strSql=strSql & " and md.request_id=mir.request_id"  
				         strSql=strSql & " and wvn.fleet_code='" & fleetcode & "'"
				         strSql=strSql & " and mir.vessel_code=wvn.vessel_code"
				         if sDate1<>"" then
			                  strSql=strSql & " and trunc(mir.inspection_date) between '" & sDate1 & "' and '" & sDate2 & "'"
                         end if
				         strSql=strSql & " order by mvq.sort_order"
				         Set rsDetails=connObj.execute(strSql)
				         
				         set rsDeficiency=server.CreateObject("ADODB.Recordset")
				         strSql= " select section,question_text,vessel_name,count(*) count1,mvq.sort_order,mm.moc_id,mir.vessel_code"
				         strSql=strSql & " from moc_deficiencies md,moc_viq_questions mvq ,moc_inspection_requests mir,"
				         strSql=strSql & " moc_master mm,wls_vw_vessels_new wvn"
				         strSql=strSql & " where md.section=mvq.question_number "
				         strSql=strSql & " and mir.moc_id=mm.moc_id"
				         strSql=strSql & " and md.request_id=mir.request_id"
				         strSql=strSql & " and mir.vessel_code=wvn.vessel_code"
				         strSql=strSql & " and wvn.fleet_code='" & fleetcode & "'"
				         if sDate1<>"" then
			                  strSql=strSql & " and trunc(mir.inspection_date) between '" & sDate1 & "' and '" & sDate2 & "'"
                        end if
				         strSql=strSql & " group by vessel_name,section,question_text,mvq.sort_order,mm.moc_id,mir.vessel_code"
				         strSql=strSql & " order by mvq.sort_order,vessel_name"
				         rsDeficiency.CursorLocation=adUseClient
				         rsDeficiency.Open strSql,connObj,adOpenDynamic,adLockReadOnly
				         while not rsDetails.eof
								 sect=rsDetails("section")%>
								<tr bgcolor=white>
								  <td align="left"><%=rsDetails("section")%></td>
								  <td><%=rsDetails("question_text")%></td>
								  <%rsDeficiency.Filter="section='" & sect & "'"
								  rsVessel.movefirst
								  while not rsVessel.EOF 
										 str2=""
										 count2="&nbsp;"
										  if not rsDeficiency.EOF then
										    if rsVessel("vessel_name")=rsDeficiency("vessel_name") then
										        count2=rsDeficiency("count1")
										        moc_id=rsDeficiency("moc_id")
										        vessel_code=rsDeficiency("vessel_code")
										        rsDeficiency.movenext
										    end if						   
										end if%>
										 <td align="center"><a href="rpt_observations.asp?VID=<%=vessel_code%>&VIQ=<%=sect%>&MOC=<%=moc_id%>&sDate1=<%=sDate1%>&sDate2=<%=sDate2%>"><%=count2%></a></td>
								      <%rsVessel.MoveNext
								  wend%>
								</tr>
							<%rsDetails.MoveNext 
				         wend
				         %>				                        
				</table><br><br>			    
         <%rsFleetCode.movenext
      wend
    %>
</body>
</html>