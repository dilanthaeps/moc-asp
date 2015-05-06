<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	
    Response.Buffer = true	 
	dim vess_code,notify_type,v_ctr,class_color		
	if request("vessel_code")<>""  then
		vess_code=request("vessel_code")    		
	end if
	if request("notification_type")=""  then
	   notify_type=request("notification_type")		
	end if	
	Response.ContentType = "application/vnd.ms-excel"
%>
<html>
<head>
<TITLE>Vessel Notification - Tanker Pacific</TITLE>
<META HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 2000 14:25:27 GMT">
<LINK REL="stylesheet" HREF="http://webserve2/wls/moc/moc.css"></LINK>
<style>
.clsIncident
{
	font-family:webdings;
	font-size:15px;
	color:red;
	cursor:default;
}
</style>
</head>
<body>
<table border=0 width=100%>
<tr>
	<td  align=center colspan=1>
		<p>
		<font size=4 style="font-weight:bold" >Notification Requests </font>
	</td>
</tr>
</table>
<%
   strSql = " select distinct vessel_name,mwn.vessel_code,notification_type,delivery_port,to_char(delivery_date,'DD-Mon-YYYY') delivery_date1,notification_id" 
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
   Set rsObj=connObj.execute(strSql)
%>
<TABLE WIDTH=100% BORDER="1" CELLPADDING="0" CELLSPACING="1">
  <tr>
    <td class="tableheader" nowrap>Vessel</td>
    <td class="tableheader" valign=top>Notification Type</td>    
    <td class="tableheader" valign=top>Delivery Date</td>
    <td class="tableheader" valign=top>Delivery Port</td>    
  </tr>
</TABLE>
<TABLE WIDTH=100% BORDER="1" CELLPADDING="0" CELLSPACING="1">
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
			  <tr valign=bottom>
			    <td class="<%=class_color%>"><%=rsObj("vessel_name") %>&nbsp;</td>
			    <td class="<%=class_color%>"><%=rsObj("notification_type") %>&nbsp;</td>
			    <td class="<%=class_color%>"><%=mid(rsObj("delivery_date1"),1,7)&mid(rsObj("delivery_date1"),10,2) %>&nbsp;</td>
			    <td class="<%=class_color%>"><%=rsObj("delivery_port") %>&nbsp;</td>			   
			  </tr>
			<%
			rsObj.movenext
	   wend
else
   Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if 
%>
</table>
</body>
</html>
<%
rsObj.close
set rsObj=nothing
%>