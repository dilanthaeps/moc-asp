<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<%  
    dim rsObj_vessels,rsnotification_id, notify_id, notify_type,ves_code
    dim rs_vesselcode,vess_name
    dim rsdetails,SaleMailBody,sBody,AcquMailBody,SrapMailBody
    dim v_port, v_imo_nr, v_insp_date ,saleSubject,acquSubject,scrapSubject  
		vess_name=Request.QueryString("vessel_code")		
		notify_type=Request.QueryString("notification_type")
				
   'Vessel List
   if vess_name<>"" then	
			strSql ="select vessel_code from vessels where initcap(vessel_name)='"& vess_name & "'"
			set rs_vesselcode=connObj.Execute(strSql)
			if not rs_vesselcode.eof then
			  ves_code=rs_vesselcode("vessel_code")
			end if
	
			strSql = "select nvl(max(notification_id),0)+1 id from MOC_VESSEL_NOTIFICATION"	
			set rsnotification_id=connObj.Execute(strSql)
			if not rsnotification_id.eof then
			   notify_id=rsnotification_id("id")
			end if	
			strSql="select mir.inspection_port,MOC_FN_VESSEL_IMO_NO(v.vessel_code)imo_number,"
			strSql = strSql & " to_char(inspection_date, 'DD-Mon-YYYY') inspection_date_disp,mir.vessel_code,request_id"
			strSql = strSql & " from moc_inspection_requests mir, moc_master mm, "
			strSql = strSql & " vessels v "
			strSql = strSql & " where mir.moc_id = mm.moc_id "
			strSql = strSql & " and mir.vessel_code = v.vessel_code "
			strSql = strSql & " and mir.vessel_code='" & ves_code & "'" 
			strSql = strSql & " and request_id='10388'"			
			set rsdetails=connObj.Execute(strSql)
			if not rsdetails.eof then
			    v_port=rsdetails("inspection_port")	
			    v_imo_nr=rsdetails("imo_number")
			    v_insp_date=rsdetails("inspection_date_disp")       
			end if 
	  
	   end if      
			 saleSubject ="Update on Fleet - MT "& vess_name & "( IMO No. <IMO> )"		
			 SaleMailBody= "To 	: " & vbcrlf
			 SaleMailBody=SaleMailBody & "Attn  :" & v_port & vbcrlf & vbcrlf 
		     SaleMailBody=SaleMailBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf
			 SaleMailBody=SaleMailBody & " Dear Mr " & vbcrlf & vbcrlf
			 SaleMailBody=SaleMailBody & " Please be advised that effective " & v_insp_date & "the captioned vessel is no longer under our management."
			 SaleMailBody=SaleMailBody & "  Kindly delete this vessel from our records accordingly."
			 SaleMailBody=SaleMailBody & "  For smooth and prompt response," 
			 SaleMailBody=SaleMailBody & " we would appreciate it if all vetting-related communications are directed to our vessel "
			 SaleMailBody=SaleMailBody & " inspections group email : mocinspectionteam@tanker.com.sg."        
			 	 
			acquSubject = "New Acquisition Update - MT "& vess_name & "( IMO No. <IMO> )"
			AcquMailBody=" To 	: " & vbcrlf
			AcquMailBody=AcquMailBody & "Attn  :"  & vbcrlf & vbcrlf
			AcquMailBody=AcquMailBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf
			AcquMailBody=AcquMailBody & " Dear Mr " & vbcrlf & vbcrlf
			AcquMailBody=AcquMailBody & " This is to advise you that on behalf of the Owners, Tanker Pacific Management (Singapore) Pte Ltd "
			AcquMailBody=AcquMailBody & " has taken over the full management of MT < vess_name >  on < v_insp_date> at <DELIVERY PORT>.  We attach herewith "
			AcquMailBody=AcquMailBody & " details of the vessel for your records." &vbcrlf 
           AcquMailBody=AcquMailBody & " For smooth and prompt response, we would appreciate it if all vetting-related" 
           AcquMailBody=AcquMailBody & " communications are directed to our vessel inspections group email : mocinspectionteam@tanker.com.sg." 
			sBody =AcquMailBody
      
      
			scrapSubject = "New Acquisition Update - MT "& vess_name & "( IMO No. <IMO> )"
			SrapMailBody=" To 	: " & vbcrlf
			SrapMailBody=SrapMailBody & "Attn  :"  & vbcrlf & vbcrlf
			SrapMailBody=SrapMailBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf
			SrapMailBody=SrapMailBody & " Dear Mr " & vbcrlf & vbcrlf			
			SrapMailBody=SrapMailBody & " This is to advise that our managed vessel, MT <VESSEL> ( IMO No. <IMO> ),"
			SrapMailBody=SrapMailBody & " has been delivered to the Shipbreakers at <DELIVERY PORT> on <DELIVERY DATE>."
            SrapMailBody=SrapMailBody & " The vessel is scheduled to be beached on/arnd <DATE>.  Therefore, we request you "
            SrapMailBody=SrapMailBody & " to delete this vessel from your records." & vbcrlf 
            SrapMailBody=SrapMailBody & " For smooth and prompt response, we would appreciate it if all vetting-related" 
            SrapMailBody=SrapMailBody & " communications are directed to our vessel inspections group email : mocinspectionteam@tanker.com.sg." 
		    sBody =SrapMailBody	
          %>
          <%Response.Write(vess_name)%>

 <HTML>
<head>
<meta HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 1996 14:25:27 GMT">
<link REL="stylesheet" HREF="moc.css"></link>
<style>
.clsFile
{
	font-size:9px;
}
</style>
<script language="vbscript" >
<!--
function insertvalues()    
    dim sBody,sSubject
    if form1.vessel_code.value="" then        
        msgbox("Please select a vessel")
        exit function
    end if 
    if form1.notification_type.value="" then
        msgbox("Please select notification type")        
        exit function
    end if       
      form1.action = "ins_vessel_notification.asp"     
      form1.submit
    if   form1.vessel_code.value<>"" then
       if form1.notification_type.value="sale" then
         sBody=form1.txtsaleMailBody.value      
         sSubject=form1.txtsaleMailSubject.value
          
      elseif  form1.notification_type.value="acquisition"  then
          sBody = form1.txtacquMailBody.value          
	      sSubject = form1.txtacquMailSubject.value	      
	   else 
	     sBody = form1.txtscrapMailBody.value      
	     sSubject = form1.txtscrapMailSubject.value
	   end if 
	   
	     form1.mail.displayMailClient "","","",cstr(sSubject),cstr(sBody),"" 
	   end if    
end function 
-->  
</script>
</HEAD>
<BODY class=bcolor><br>
<form name=form1>
<textarea name="txtsaleMailBody" class=textareah><%=SaleMailBody%></textarea>
<textarea name="txtsaleMailSubject" class=textareah><%=saleSubject%></textarea>
<textarea name="txtacquMailBody" class=textareah><%=AcquMailBody%></textarea>
<textarea name="txtacquMailSubject" class=textareah><%=acquSubject%></textarea>
<textarea name="txtscrapMailBody" class=textareah><%=SrapMailBody%></textarea>
<textarea name="txtscrapMailSubject" class=textareah><%=scrapSubject%></textarea>
<input type=hidden id="check" name="check" value="<%=v_port%>">
<center><span style="font-size:20px;font-weight:bold;color:maroon;padding:5px;">Create Vessel Notification</span></center>
<object id="mail" style="LEFT: 0px; TOP: 0px" name="mail"   codebase="../../../MailClient.CAB" classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" width="0" height="0" VIEWASTEXT>
	<param NAME="_ExtentX" VALUE="26">
	<param NAME="_ExtentY" VALUE="26">
</object>
 <br>
  <table width="0%" border="0" cellspacing="1" cellpadding="0" align=center>
    <tr>
      <td class="tableheader" align=right>Vessel</td>
      <td>		 
          <select name="vessel_code" id="vessel_code" onkeypress="control_onkeypress">
			<option value="<%%>">--Select Vessel--</option>
			<%	
			    strSql = "SELECT vessel_code, initcap(vessel_name) vessel_name"
	            strSql = strSql & " from vessels "
	             strSql = strSql & " order by vessel_name "	
	             set rsObj_vessels = connObj.Execute(strSql)				
				while not rsObj_vessels.eof%>				
			<option value="<%=rsObj_vessels("vessel_name")%>"><%=rsObj_vessels("vessel_name")%></option>
			<%			  rsObj_vessels.movenext
				wend%>			
        </select> </td></tr>  
	  <tr><td class="tableheader" align=right>Notification Type</td>      
      <td>
          <select name="notification_type" id="notification_type" onkeypress="control_onkeypress">
			<option value="<%%>">--Select Notification--</option>
			<option value="sale">SALE</option>
			<option value="acquisition">ACQUISITION</option>
			<option value="scrap">SCRAP</option>			
        </select></td>
       </tr>
       <tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr>       
       <tr><td align=center colspan=2><input type=button name="notify" value="Create Notification" onclick="javascript:return insertvalues()"></td></tr></table>
       
     </BODY>
</HTML>
