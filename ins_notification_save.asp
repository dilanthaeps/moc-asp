<%@ Language=VBScript %>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<html>
<link REL="stylesheet" HREF="moc.css"></link>
<body>
<OBJECT id=TPMDnDCommonCtrl1 style="display:none;height:100px;width:100%;LEFT: 0px; TOP: 0px; BACKGROUND-COLOR: midnightblue" 
		data=data:application/x-oleobject;base64,EcFFRl5khkOn3XMSXMq6vAAHAADYEwAA2BMAAA== 
		classid=clsid:4645C111-645E-4386-A7DD-73125CCABABC VIEWASTEXT></OBJECT>
<div id=divMessage style="width:300px;height:100px;text-align:center;padding:10px;position:absolute;background-color:gold;font-weight:bold;font-size:14px">
<br>
Saving data and uploading files<br><br>
Please wait ...
<br><br>
</div>
</body>
</html>
<script language=vbscript>
divMessage.style.left = document.body.offsetWidth/2 - divMessage.offsetWidth/2
divMessage.style.top = document.body.offsetHeight/2 - divMessage.offsetHeight/2
</script>
<%Response.buffer=true%>
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
   function PathFromID(id,sPath)
		dim strID,strPath,strName,strExt,i
		strID = right("00000000" & id,8)
		For i = 1 To 5 Step 2
		    strPath = strPath & Mid(strID, i, 2) & "\"
		Next
		strName = mid(sPath,InStrRev(sPath,"\")+1)
		if instr(1,strName,".")>0 then
			strExt = mid(strName,InStrRev(strName,"."))
		end if
		PathFromID = strPath & id & strExt
   end function

</SCRIPT>
<%
Dim v_no_error, vessel_code, v_message, notification_id
Dim delivery_date, delivery_port, notification_type	

	v_message = ""
	If request("notification_id")<>"" Then
		notification_id=request("notification_id")
	Else
		notification_id="0"
		v_no_error = "false"
		v_message = v_message&"Notification ID cannot be empty<br>"
	End if	
	v_no_error="true"
	
	vessel_code	= request("vcode")
	delivery_date =	request("inspection_date")
	delivery_port =  request("inspection_port")
	notification_id = request("notification_id")
	notification_type =	request("notification_type")	
	
	If (v_no_error <> "true") Then
		Response.Clear
		Response.Redirect "ins_vessel_notification.asp?v_ins_request_id="&notification_id&"&v_message="&v_message
	End If
	Response.Flush	
	Dim Idval, rs, rsFile
	set rs = Server.CreateObject("ADODB.Recordset")
	set rsFile = Server.CreateObject("ADODB.Recordset")
	with rs
		.CursorLocation = adUseClient
		.CursorType = adOpenDynamic
		.LockType = adLockOptimistic
		set .ActiveConnection = connObj
	end with
	with rsFile
		.CursorLocation = adUseClient
		.CursorType = adOpenDynamic
		.LockType = adLockOptimistic
		set .ActiveConnection = connObj
	end with

	Idval = Request("notification_id")

	rs.open "Select * from MOC_VESSEL_NOTIFICATION where notification_id = '" & Idval & "'"
	if rs.eof then rs.addnew
	rs("vessel_code")= vessel_code
	rs("Created_by") =  USER
	rs("notification_type") = notification_type 
	If delivery_port <> "" Then rs("delivery_port") = delivery_port
	If delivery_date <> "" Then rs("delivery_date") = delivery_date	
	rs.Update
	if Idval = "" then
		rs.Resync	
		Idval = rs("notification_id")
	end if
	
	v_message = "MOC Notification details Updated Successfully"		
	'insert records for uploaded documents
	Dim i, sClientScript, obj
	i = 1
	rsFile.Open "Select * from moc_documents where doc_type='NOTIFICATION' and parent_id = '" & Idval & "'"
	for each obj in Request.Form("txtDocID")
		if obj = "" then
			rsFile.AddNew
			rsFile("doc_type") = "NOTIFICATION"
			rsFile("parent_id") = Idval
			rsFile("doc_name") = Request.Form("txtDocName")(i)
			rsFile("uploaded_by") = USER

			rsFile.Update
			rsFile.Resync
			
			rsFile("doc_path") = PathFromID(rsFile("doc_id"), Request.Form("txtDocPath")(i))
			rsFile.Update
			
			sClientScript = sClientScript & "TPMDnDCommonCtrl1.UploadFile """ & Request.Form("txtDocPath")(i) & """, """ & MOC_NOTIFICATION_PATH & rsFile("doc_path") & """" & vbcrlf
		end if
		i = i + 1
	next	
	connObj.Close
	set connObj=nothing
If Request("v_save_close") = "Save and Close" Then
%>
	<SCRIPT LANGUAGE="vbscript">
		function window_onload
			DoUpload
			on error resume next
			self.parent.opener.document.form1.action = "ins_request_notification_maint.asp"
			self.parent.opener.document.form1.target = ""
			self.parent.opener.document.form1.submit
			self.close()
		end function
		
		sub DoUpload
			<%=sClientScript%>
		end sub
	</SCRIPT>
<%Else%>
	<SCRIPT LANGUAGE="vbscript">
		function window_onload
			DoUpload		
			self.location.href = "ins_vessel_notification.asp?VALUE=SAVE&v_ins_request=<%=IdVal%>&v_message=" & v_message 
		end function
		
		sub DoUpload
			<%=sClientScript%>
		end sub
	</SCRIPT>	
<%End If%>