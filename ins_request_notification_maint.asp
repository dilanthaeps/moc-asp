<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<html>
<head>
<title>Vessel Notifications - Tanker Pacific</title>
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
</style>
<script language="Javascript" src="js_date.js"></script>
<script language="Javascript" src="AutoComplete.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<script language="vbscript">
Sub window_onbeforeprint	
	divTopMenu.style.display="none"	
End Sub

Sub window_onafterprint	
	divTopMenu.style.display=""	
End Sub
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function fncall(v_ins_request_id,vess_name)
{
	
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes,resizable=yes,status=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=160,top=10,width=700,height=650'
	}else{
		winStats+=',screenX=350,screenY=200,width=400,height=280'
	}	
	adWindow=window.open("ins_vessel_notification.asp?v_ins_request_id="+v_ins_request_id+"&VessNameVal="+vess_name,"moc_notification",winStats);
	adWindow.focus();
}
function fnnotify(v_ins_request_id)
{
        winStats='toolbar=no,location=no,directories=no,menubar=no,'
	    winStats+='scrollbars=yes,resizable=yes,status=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=160,top=10,width=700,height=650'
	}else{
		winStats+=',screenX=350,screenY=200,width=400,height=280'
	}	
		adWindow=window.open("ins_vessel_notification.asp?v_ins_request_id="+v_ins_request_id,"moc_notification",winStats);
	    adWindow.focus();	    			
}
function v_sort(v_sort_field,v_sort_order)
{
	document.form1.action="ins_request_notification_maint.asp?item="+v_sort_field+"&order="+v_sort_order
	document.form1.submit();
}
function v_clear_all_filters()
{
	location.href="ins_request_notification_maint.asp";
}
function outputExcel()
{
	var preAction = document.form1.action;
	var preTarget = document.form1.target;
	var qryString = "";
	
	qryString += "?vessel_code=" + document.form1.vessel_code.value;	
	qryString += "&notification_type=" + document.form1.notify_type.value;	
	document.form1.action = "ins_notification_maint_excel.asp" + qryString;
	document.form1.target = "_blank";
	document.form1.submit();
	
	document.form1.action = preAction;
	document.form1.target = preTarget;

	return false;
}

function window_onload() {
<%if request.TotalBytes=0 then%>
	divCount(1).innerHTML = ""
	divCount(0).innerHTML = "<font color=red><b>No filters specified. Please specify filters and click refresh.</b></font>"
<%else %>
	divCount(0).innerHTML = divCount(1).innerHTML
<%end if%>	
}

//-->
</SCRIPT>
</head>
<body class="bgcolorlogin" LANGUAGE=javascript onload="return window_onload()">
<div id=divTopMenu>
  <!--#include file="menu_include.asp"-->
</div>
<center>
<table border="0" width="100%">
<tr height="30pt" VALIGN="bottom">
		<td align="center">		
			<h3 style="margin-bottom:0">Vessel Notifications </h3>		
		</td>
</tr>
</table>
<% 
    dim v_button_disabled,rsObj_vessel,notify_type,vessel_code,str2,notification
	v_button_disabled = "DISABLED"
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Then
	if UserIsAdmin then
		v_button_disabled = ""
	End If	 
   strSql = " select distinct vessel_name,mwn.vessel_code,upper(notification_type) notification_type,delivery_port,delivery_date,notification_id" 
   strSql = strSql &  " from moc_vessel_notification mwn,vessels wn"
   strSql = strSql &  " where mwn.vessel_code=wn.vessel_code "   
   if request("vessel_code")<>"" then
		 strSql = strSql &  " and mwn.vessel_code='" & request("vessel_code") & "'"
   end if
   if request("notify_type")<>"" then
		strSql = strSql &  " and trim(notification_type) ='" & request("notify_type") & "'"
   end if 
     strSql = strSql & " order by mwn.vessel_code"
 'Response.Write strSql
  if Request.TotalBytes=0 then
	Set rsObj=connObj.execute("Select * from moc_vessel_notification where 1=2")	
  else
	Set rsObj=connObj.execute(strSql)
  end if
   strSql="Select vessel_code, vessel_name, tech_manager fleet_code from vessels order by vessel_name"
' where vessel_code in (select distinct vessel_code from moc_inspection_requests) order by vessel_name"
  Set rsObj_vessel = connObj.execute(strSql)
%>
<form name="form1" method="post" action="ins_request_notification_maint.asp">
<table  width="20%" border="0" cellspacing="1" cellpadding="1" align="center">
  <tr>    
    <td class="tableheader" >Select Vessel
    <td class="tableheader">Select Type   
  </tr>
  <tr>     
      <td class="tabledata">             
             <select id="vessel_code" name="vessel_code" >
             <option value="<%%>">--All Vessels--</option>
				<%	rsObj_vessel.filter=0
					if not(rsObj_vessel.eof or rsObj_vessel.bof) then
						while not rsObj_vessel.eof
				%>
						<option value="<%=rsObj_vessel("vessel_code")%>" <%if request("vessel_code")=rsObj_vessel("vessel_code") then Response.Write "selected"%>><%=rsObj_vessel("vessel_name")%></option>
				<%		rsObj_vessel.movenext
						wend
					end if
				%>
			</select>
      </td>
      <td class="tabledata">
          <select id="notify_type" name="notify_type">         
             <option value="">--Select Type--</option>
             <option value="sale" <%if request("notify_type")="sale" then Response.Write("selected")%>> SALE </option>
			 <option value="acquisition" <%if request("notify_type")="acquisition" then Response.Write("selected")%>> ACQUISITION </option>
			 <option value="scrap" <%if request("notify_type")="scrap" then Response.Write("selected")%>> SCRAP </option>
			 <option value="others" <%if request("notify_type")="others" then Response.Write("selected")%>> OTHERS </option>						
		 </select></td>      
  </tr><table>
  <table border="0" cellspacing="1" cellpadding="1" align="center"><td>
	<input type="submit" value="Refresh" style="font-weight:bold" id="submit1" name="submit1" class="cmdButton"></td>	
	<td><input type="button" value="Create New Notification" <% =v_button_disabled %> NAME="v_notify" onclick="javascript:fnnotify('0');" class="cmdButton" style="width:185px"></td></tr>
 </table><br><br>
 <table WIDTH="80%" BORDER="0">
	<tr HEIGHT="40pt" VALIGN="bottom">
		<td><div id=divCount align="left"></div>
		<td WIDTH="50%" ALIGN="right">
			<a href="javascript:window.excelOut()" OnClick="return outputExcel();"><img src="Images/EXCEL.ICO" border="0" alt="Export this Page to Excel"></a>&nbsp;
			<a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
		</td>
	</tr>
</table>
<table width="80%" border="0">
  <tr>
    <td class="tableheader" VALIGN="top" nowrap>Vessel</td>	
	<td class="tableheader" VALIGN="top" nowrap>Notification Type</td>			
	<td class="tableheader" valign="top">Delivery Date</td>    
    <td class="tableheader" valign="top">Delivery Port</td>	      
  </tr>
 </form>
<%dim v_ctr,class_color,v_code,date_format
v_ctr=0
if not (rsObj.bof or rsObj.eof) then
	while not rsObj.eof		                
	     v_ctr=v_ctr+1
		if (v_ctr mod 2) = 0 then
			class_color="columncolor2"
		else
			class_color="columncolor3"
	    end if
	    if rsObj("delivery_date")<>"" then
	     date_format=day(rsObj("delivery_date")) & "-" & MonthName(Month(rsObj("delivery_date")),true) & "-" & year(rsObj("delivery_date"))	   
	    end if
	 %>
	  <tr valign="bottom">	     
	     <td class="<%=class_color%>"><a name="1" href="javascript:fncall('<%=rsObj("notification_id") %>','<%=rsObj("vessel_name")%>');"><%=rsObj("vessel_name")%></a></td>
	     <td class="<%=class_color%>" align="left"><%=rsObj("notification_type")%></td>
	     <td class="<%=class_color%>" align="left"><%=date_format%></td>
	     <td class="<%=class_color%>" align="left"><%=rsObj("delivery_port")%></td>
	  <%rsObj.movenext%></tr>
   <%wend
else   
   Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if

%>
</table>
<blockquote><blockquote><blockquote>
<div id=divCount align="left" style="display:none"><strong>Number of inspections :</strong> <%=v_ctr%></div></blockquote></blockquote></blockquote>
</body>
</html>
<%
rsObj.close
set rsObj = nothing

rsObj_vessel.close
set rsObj_vessel = nothing

%>