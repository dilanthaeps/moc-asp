<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<!--#include file="ado.inc"-->
<%  
Dim rsObj_vessels, rsnotification_id, notify_id, notify_type, ves_name
Dim rs_vesselcode, rsimo, imo_nr, vess_code, rsinsert, VessNameVal, v_name
Dim sBody, vstr, vess_code1, v_port, v_imo_nr, v_insp_date, v_notification, filter   
Dim Idval, v_header, v_mode, del_date, del_port, v_tem, rsDocs, v_button_disabled, str2

    v_button_disabled = "DISABLED"	
	If UserIsAdmin Then v_button_disabled = ""

	Idval = Request.QueryString("v_ins_request_id")	
	If Request.QueryString("VALUE") = "SAVE" Then Idval = Request.QueryString("v_ins_request")

	VessNameVal=Request.QueryString("VessNameVal")	  

	if Idval <> "0" then
		v_mode = "edit"
		filter = Idval
		v_header = "Vessel Notification Details"
		if request("v_read_mode") = "Yes" then
		   v_header = v_header & "(Read Only)"
		end if				
	else
		v_header="Create New Notification"
	end if
	vess_code = Request.QueryString("vesselcode")			 		
	notify_type = Request.QueryString("notify_type")
	del_port = Request.queryString("delivery_port")	
	del_date = Request.QueryString("delivery_date")
	v_name = Request.QueryString("vess_name")							
		
	strSql = "SELECT initcap(vessel_name) vessel_name,mvn.vessel_code,trim(notification_type) notification_type,delivery_date,delivery_port,creation_date,mvn.created_by,notification_id"
	strSql = strSql & " from moc_vessel_notification mvn,vessels wv"
	strSql = strSql & " where mvn.vessel_code=wv.vessel_code"
	strSql = strSql & " and notification_id='" & filter & "'"		  
	Set rsObj = connObj.Execute(strSql)
%>				 
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
<script language="VBScript" runat="server">
 function SFIELD(fname)     
	if v_mode="edit" then	  
		if not rsObj.eof then     	
	 		v_tem = rsObj(cstr(fname))	  		 	 	
			if fname="delivery_date" and v_tem<>""then
			   v_tem=day(v_tem) & "-" & MonthName(Month(v_tem),true) & "-" & year(v_tem)
			end if 
			SFIELD=v_tem 
	   end if 	
	else
	   SFIELD = ""
	end if	
End function

function ROB ' Read Only Button
	if request("v_read_mode") = "Yes" then
		ROB = " disabled "
	else
		ROB = ""
	end if
end function
</script>
<SCRIPT LANGUAGE="vbscript">
Sub TPMDnDCommonCtrl1_OnFileAdded(index)
	set obj = form1.TPMDnDCommonCtrl1
	dim oTr,oTd,sFile,sHTML,sName,sNewName,sFeatures,vnotify,v_name
	vnotify=form1.notification_type.value
	v_name=form1.vname.value
	sFile = form1.TPMDnDCommonCtrl1.GetFileNameAt(index)
	
	for each tr in tbFiles.rows
		if s = tr.cells(1).children(0).GetAttribute("filepath") then
			ShowMessage "File already exists in upload list"
			exit sub
		end if
	next	

	'sName = window.showModalDialog("SelectFileType.htm","","dialogHeight:156px;dialogWidth:165px;resizable:yes;center:yes;status:no;scroll:no;")	
	'sName =v_name & "/" & vnotify & " /" & sName	
	sName =v_name & "/" & vnotify & " /" & "Notification" 
	
	sNewName = InputBox("Please specify a name for the file you are uploading" & vbcrlf & vbcrlf & _
			"This will be used to describe the document in all displays", "Fileuploader", sName)
	if sNewName<>"" then sName = sNewName
	
	set oTr = tbFiles.insertRow
	oTr.style.backgroundColor="white"
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	sHTML = "<a href='" & sFile & "' filepath='" & sFile & "' target='moc_document'>" & sName & "</a>"
	sHTML = sHTML & "<input name=txtDocID type=hidden value=''><input name=txtDocName type=hidden value='" & sName & "'><input name=txtDocPath type=hidden value='" & sFile & "'>"
	oTd.innerHTML = sHTML
	
	set oTd = oTr.insertCell
	oTd.className = "clsFile"
	oTd.innerHTML = "<span style='cursor:hand;color:blue' onclick='RemoveFile(me.parentElement.parentElement)'>delete</span>"
End Sub
sub RemoveFile(objTr)
	form1.TPMDnDCommonCtrl1.RemoveFileNameFromList(objTr.cells(1).children(0).getAttribute("filepath"))
	tbFiles.deleteRow(objTr.sectionRowIndex)
end sub
sub ShowMessage(s)
	window.status = s
	setTimeout "ClearMessage",2000
end sub

sub ClearMessage
	window.status=""
end sub
</script>
<script language="vbscript">
function SendMail()    
	   dim sBody,sSubject,imr_nr,format_date,str,v_vessel,str2
	   dim v_port,v_del_date,v_notify	   
	   v_vessel=form1.vessel_code.value
	   str=split(v_vessel,"~~")	   
	   'v_name= form1.vessel_code.options(form1.vessel_code.selectedindex).text 
	   v_port=form1.inspection_port.value
	   v_del_date= form1.inspection_date.value
	   v_notify=form1.notification_type.value	
	   str2=replace(replace(str(1),chr(13),""),chr(10),"")
	   form1.action="ins_vessel_notification.asp?ACTION=SAVE&vesselCode=" & str(0) & "&notify_type=" & v_notify & "&delivery_port=" & v_port & "&delivery_date=" & v_del_date &  "&vess_name=" & str(2) & "&imr_nr=" & str2
	   form1.submit	    
	   format_date = day(now) & " " & MonthName(Month(now),true) & " " & year(now)
	 if form1.notification_type.value="sale" then
			 sSubject ="Update on Fleet - MT " & str(2) & " ( IMO No. " & str2 & ")"		
			 sBody= " To   : " & vbcrlf
			 sBody=sBody & "Attn  :" & vbcrlf & vbcrlf 
		     sBody=sBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf & vbcrlf
			 sBody=sBody & " Dear Mr " & vbcrlf & vbcrlf
			 sBody=sBody & "  Please be advised that effective  " & v_del_date & " , the captioned vessel is no longer under our management."  & vbcrlf & vbcrlf 
			 sBody=sBody & "  Kindly delete this vessel from our records accordingly."  & vbcrlf & vbcrlf 
			 sBody=sBody & "  For smooth and prompt response," 
			 sBody=sBody & " we would appreciate it if all vetting-related communications are directed to our vessel "  &  vbcrlf 			 
			 sBody=sBody & "  inspections group email : mocinspectionteam@tanker.com.sg."        
	 elseif  form1.notification_type.value="acquisition"  then		 	 
			sSubject = "New Acquisition Update - MT " &  str(2) & " ( IMO No. " & str2 & ")"
			sBody= " To   : " & vbcrlf
			sBody=sBody & "Attn  :"  & vbcrlf & vbcrlf
			sBody=sBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf & vbcrlf
			sBody=sBody & " Dear Mr " & vbcrlf & vbcrlf
			sBody=sBody & " This is to advise you that on behalf of the Owners, Tanker Pacific Management (Singapore) Pte Ltd "
			sBody=sBody & " has taken over the full" & vbcrlf
			sBody=sBody & " management of MT  " & str(2) & "  on  " & v_del_date & "  at  " & v_port & ".  We attach herewith "
			sBody=sBody & " details of the vessel for your records." &vbcrlf &vbcrlf 
	        sBody=sBody & " For smooth and prompt response, we would appreciate it if all vetting-related" 
	        sBody=sBody & " communications are directed to our vessel  " &vbcrlf 
	        sBody=sBody & " inspections group email : mocinspectionteam@tanker.com.sg." 
	  else	        
			sSubject = "Update on Fleet - MT " &  str(2) & " ( IMO No. " & str2 & ")"
			sBody= " To   : " & vbcrlf
			sBody=sBody & "Attn  :"  & vbcrlf & vbcrlf
			sBody=sBody & " Fm  : Tanker Pacific Management (Singapore) Pte Ltd"  & vbcrlf & vbcrlf & vbcrlf
			sBody=sBody & " Dear Mr " & vbcrlf & vbcrlf			
			sBody=sBody & " This is to advise that our managed vessel, MT " &  str(2) & "( IMO No." & str2 & ")"
			sBody=sBody & " has been delivered to the Shipbreakers" & vbcrlf
			sBody=sBody & " at " & v_port & "  on  " & v_del_date & "." & vbcrlf
	        sBody=sBody & " The vessel is scheduled to be beached on/arnd " & format_date & ".  Therefore, we request you "
	        sBody=sBody & " to delete this vessel from your records." & vbcrlf & vbcrlf
	        sBody=sBody & " For smooth and prompt response, we would appreciate it if all vetting-related" 
	        sBody=sBody & " communications are directed to our vessel " & vbcrlf	       
	        sBody=sBody & " inspections group email : mocinspectionteam@tanker.com.sg." 
	end if		    
	   form1.mail.displayMailClient "","","",cstr(sSubject),cstr(sBody),"" 		 		   
end function

function SavePage() 
	Dim str11  
	If form1.vessel_code.value = "" Then
		MsgBox("Please select a vessel")	       
		Exit Function
	end if 
	
	if form1.notification_type.value = "" then
		MsgBox("Please select notification type")	         
		Exit Function
	end if
	
	If form1.notification_id.value <> "" Then
		str11 = form1.notification_id.value
	Else
		str11 = 0
	End If 

	If str11 = 0 Then 
		SendMail()
		form1.vcode.value = Mid(form1.vessel_code.value, 1, instr(form1.vessel_code.value, "~~") -1)
	    form1.action="ins_notification_save.asp"
	    form1.submit      
	Else  
	    form1.action="ins_notification_save.asp"
	    form1.submit      
	End If 
End Function 
</script>
<script language="Javascript" src="js_date.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<script language="JavaScript" src="autocomplete.js"></script>
<script language="JAVASCRIPT">
function port_list()
{
	winStats='toolbar=no,location=no,directories=no,menubar=no,'
	winStats+='scrollbars=yes'
	if (navigator.appName.indexOf("Microsoft")>=0) {
		winStats+=',left=400,top=10,width=270,height=300'
	}
	else{
		winStats+=',screenX=350,screenY=120,width=575,height=500'
	}
	adWindow=window.open("port_list.asp","port_list",winStats);
	adWindow.focus();
}
function cClose(){
	var name= confirm("Are you sure? The changes will be lost!!")
	if (name== true) {
		v_val = "ins_request_notification_maint.asp?";
		self.opener.form1.action=v_val;
		self.opener.form1.submit();
		self.close();
		return false;
	}
	else {
		return false;
	}
}

</script>
</HEAD>
<BODY class=bcolor><br>
<form name=form1  method="post" onsubmit="javascript:return validate_fields();">
<input type=hidden name="notify_type" id="notify_type" value="<%=notify_type%>">
<center><span style="font-size:20px;font-weight:bold;color:maroon;padding:5px;"><%=v_header%></span></center>
<object id="mail" style="LEFT: 0px; TOP: 0px" name="mail"    codebase="MailClient.CAB" classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" width="0" height="0" VIEWASTEXT>
	<param NAME="_ExtentX" VALUE="26">
	<param NAME="_ExtentY" VALUE="26">
</object>
 <br><br><br>
 <table width="100%" border="0" cellspacing="1" cellpadding="0" align=center>
    <tr><td class="tableheader" align="right" width="16%">Vessel</td>
        <%Dim str11
        str1=""
        if Idval<>0 then
              str=SFIELD("vessel_name")
              str11="disabled"
          else             
             str=v_name             
          end if           
          	%>  
          <input type=hidden id="vcode" name="vcode" value="<%=SFIELD("vessel_code")%>">          
         <input type=hidden id="vname" name="vname" value="<%=SFIELD("vessel_name")%>">         
        <td><select name="vessel_code" id="vessel_code" onkeypress="control_onkeypress" <%=str11%>>
			    <option value="">--Select Vessel--</option> 
			<%   
			     Dim str1
			     strSql=" SELECT distinct TRIM(ANSWER_DATA) RET_IMO_NO,wv.vessel_code,initcap(vessel_name) vessel_name"
                 strSql = strSql & " FROM VPD_MASTER_ANSWERS vma,vessels wv"
                 strSql = strSql & " where vma.vessel_code=wv.vessel_code"
                 strSql = strSql & " and MASTER_QUESTION_ID = 10001326"
	             strSql = strSql & " order by vessel_name "		
	             set rsObj_vessels = connObj.Execute(strSql)	             			
				 while not rsObj_vessels.eof
				    %>				
			        <option value="<%=rsObj_vessels("vessel_code")%>~~<%=rsObj_vessels("RET_IMO_NO")%>~~<%=rsObj_vessels("vessel_name")%>" <%if  rsObj_vessels("vessel_name")=str then Response.Write("selected")%>> <%=rsObj_vessels("vessel_name")%></option>
				<%	rsObj_vessels.movenext
				 wend%>	
        </select> </td>          
      <tr><td class="tableheader" align="right">Notification Type</td>      
          <td><select name="notification_type" id="notification_type" onkeypress="control_onkeypress">
			      <option value="<%%>">--Select Notification--</option> 
			      <option value="sale" <%if SFIELD("notification_type")="sale" then Response.Write "selected"%>>SALE</option>
			      <option value="acquisition" <%if SFIELD("notification_type")="acquisition" then Response.Write "selected"%>>ACQUISITION</option>
			      <option value="scrap" <%if SFIELD("notification_type")="scrap" then Response.Write "selected"%>>SCRAP</option>
			      <option value="others" <%if SFIELD("notification_type")="others" then Response.Write "selected"%>>OTHERS</option>
              			
              </select></td></tr>                            
     <tr><td class="tableheader" align="right">
             <div align="right">Delivery Port</div></td>      
         <td >
              <input type="text" name="inspection_port" value="<%=SFIELD("delivery_port")%>" maxlength="50">
              <a href="Javascript:port_list()">Select</a> </td>      
     </tr>
     <tr><td class="tableheader" align="right">
            <div align="right">Delivery Date</div></td>      
         <td>
           <input type="text" name="inspection_date" value="<%=SFIELD("delivery_date")%>" onblur="vbscript:valid_date inspection_date,'Inspection Date','form1'">
           <a HREF="javascript:show_calendar('form1.inspection_date',form1.inspection_date.value);">
           <img SRC="Images/calendar.gif" alt="Pick Date from Calendar" WIDTH="20" HEIGHT="18" BORDER="0">
       </td>
    </tr> 
    <tr><td>&nbsp;</td></tr>  
    <tr><td>&nbsp;</td></tr>   
    <tr><td>&nbsp;</td></tr> 
    <tr><td>&nbsp;</td></tr> 
<%
Dim showHide
showHide = ""
strSql= "Select * from moc_documents where deleted is null and doc_type='NOTIFICATION' and parent_id='" & Idval & "' order by doc_id"
set rsDocs = connObj.execute(strSql)
if Idval = 0 then showHide = "none"
%>  
    <tr style="display:<%=showHide%>;">
      <td colspan=3 style="font-size:12px;font-weight:bold;color:red"><br>NOTE: The documents listed here are for internal reference only.<BR>
      Please do not release any MOC related documents to third parties. 
      Kindly consult Capt Mishra if required.
    <tr style="display:<%=showHide%>;">
      <td style="vertical-align:top;">
		<OBJECT id=TPMDnDCommonCtrl1 style="height:90px;width:90%;LEFT: 0px; TOP: 0px; BACKGROUND-COLOR: midnightblue" 
		data=data:application/x-oleobject;base64,EcFFRl5khkOn3XMSXMq6vAAHAADYEwAA2BMAAA== 
		classid=clsid:4645C111-645E-4386-A7DD-73125CCABABC  codebase="../../../TPMDnDCommon.dll#version=1,0,0,0" VIEWASTEXT></OBJECT>
	  <td colspan=2 style="vertical-align:top;">
	    <table width=100% border=0 cellspacing=1 cellpadding=0 bgcolor=lightgrey>
			<tr class="tableheader">
			  <td width=15px class=clsFile>
			  <td class=clsFile>File
			  <td width=35px class=clsFile>
			</tr>
			<tr bgcolor=white><td colspan=3 class=clsFile>&nbsp;
			<%
			while not rsDocs.eof%>
			<tr bgcolor=white>
			  <td class=clsFile>
			  <td class=clsFile><a href="<%=MOC_NOTIFICATION_PATH%><%=rsDocs("doc_path")%>" target="moc_document"><%=rsDocs("doc_name")%></a>
					<input name=txtDocID type=hidden value="<%=rsDocs("doc_id")%>">
					<input name=txtDocName type=hidden value="<%=rsDocs("doc_name")%>">
					<input name=txtDocPath type=hidden value="">
			  <td class=clsFile>
			  <%if UserIsAdmin then%>
				<a href="javascript:void(0)" onclick="javascript:window.open('ins_doc_delete.asp?doc_id=<%=rsDocs("doc_id")%>','mocdocdelete','width=400,height=100,top=150,left=150,location=no');">Delete</a>
			  <%end if%>
			  <%
			  rsDocs.movenext
			wend%>
			<tbody id=tbFiles class=clsFile>
			<!--rows will be added dynamically-->
			</tbody>
	    </table>
    </tr>
    <tr>
      <td colspan="3" class="tabledata" align="center">
        <input type="button" value="Save" id="submit1" name="submit1" <%=ROB%> <%=v_button_disabled %> onclick="SavePage()">
        &nbsp;        
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="Close without Save" id="submit4" name="submit4" OnClick="javascript:return cClose();">
      </td>
    </tr>
    <tr>
      <td colspan="3">&nbsp;
   <%if v_mode="edit" then %>
    <tr>
      <td colspan="2" class="tabledata" align="left"> <strong> Created By :</strong>
        <%=SFIELD("created_by")%>&nbsp;&nbsp;<strong>Create Date: </strong><%=SFIELD("creation_date")%>
      </td>      
    </tr>
  <%end if %>
  </table>  
    <input type=hidden name="notification_id" id="notification_id" value="<%=SFIELD("notification_id")%>">    
</form>
<script LANGUAGE="javascript">
<!--
var v_mess="<%=request("v_message") %>";
if (v_mess!="") {
	alert(v_mess);
}
//-->
</script>
</BODY>
</HTML>
