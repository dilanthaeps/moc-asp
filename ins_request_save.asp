<%@ Language=VBScript %>
<%option explicit
	'===========================================================================
	'	Template Name	:	MOC Inspector  Save Screen
	'	Template Path	:	.../inspector_save.asp
	'	Functionality	:	Save new MOC Inspector details or update the existing details
	'	Called By		:	../ins_request_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:	21st August 2002
	'						1.
	'						2.
	'===========================================================================

%>
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
Function SQE(text_with_single_quote)
    If Not IsNull(text_with_single_quote) Then
    SQE = Trim(Replace(text_with_single_quote, "'", "''"))
    Else
    SQE = text_with_single_quote
    End If
    Exit Function
End Function

Function SETDATA(name_of_the_form_variable,v_comment,v_mandatory,data_type)
	if request(name_of_the_form_variable)<>"" then
			SETDATA=request(name_of_the_form_variable)
		else
			SETDATA=null
				'if data_type="Number" then
				'  SETDATA="null"
				'end if
			if v_mandatory="Yes" then
				v_no_error = "false"
				v_message = v_message & v_comment & " cannot be empty<br>"
			end if
	end if
End Function

Function IIF(expr, trueValue, falseValue)
	If expr Then
		IIF = trueValue
	Else
		IIF = falseValue
	End If
End Function
</SCRIPT>
<%
dim v_no_error,vessel_code,insp_status,inspection_date,MOC_ID,AGENT_ADVISED_DATE, Is_Sire
dim DATE_ACCEPTED,DATE_CONFIRM_REJECT,DATE_DEFS_REPLIED,DATE_DEFS_TO_VESSEL,DATE_REPLIED_TO_SIRE
dim DATE_SIRE_TO_VESSEL,EXPIRY_DATE,OFFICE_ADVISED_DATE,REQUEST_DATE,SIRE_RECD_DATE
dim TECH_ALERTED,TECH_ALERT_DATE,TECH_REPLY_DATE,VESSEL_ADVISED_DATE,EXPENCES,EXPENCES_IN_USD
dim AGENT_ID,BASIS_SIRE,INSPECTOR_ID,COMBINED,REJECTION_REASON,REMARKS,TECH_DECLINED_REASON
dim CONFIRMED_OR_REJECTED,OPERATION,TECH_STATUS,local_CURRENCY,DEFICIENCY_RECD,TECH_DEPT_REPLIED
dim INSPECTION_PORT,SUPTD_ATTENDED,TECH_PIC,status,detention,insp_type,v_message,REQUEST_ID,basis_sire_moc
dim PO_NUMBER,OCIMF_REPORT_NUMBER
'************************
	
	
	v_message=""
	
	If request("REQUEST_ID")<>"" then
		REQUEST_ID=request("REQUEST_ID")
	else
		REQUEST_ID="0"
		v_no_error = "false"
		v_message = v_message&"Request No cannot be empty<br>"
	end if
	
	v_no_error="true"
	
	vessel_code				=	SETDATA("vessel_code","Vessel Code","Yes","Char")
	insp_status				=	SETDATA("insp_status","Inspection Status","Yes","Char")
	inspection_date			=	SETDATA("inspection_date","Inspection Date","Yes","Date")
	MOC_ID					=	SETDATA("MOC_ID","MOC Details","Yes","Number")
	AGENT_ADVISED_DATE		=	SETDATA("AGENT_ADVISED_DATE","AGENT_ADVISED_DATE","No","Date")
	DATE_ACCEPTED			=	SETDATA("DATE_ACCEPTED","DATE_ACCEPTED","No","Date")
	DATE_CONFIRM_REJECT		=	SETDATA("DATE_CONFIRM_REJECT","DATE_CONFIRM_REJECT","No","Date")
	'DATE_DEFS_REPLIED		=	SETDATA("DATE_DEFS_REPLIED","DATE_DEFS_REPLIED","No")
	'DATE_DEFS_TO_VESSEL	=	SETDATA("DATE_DEFS_TO_VESSEL","DATE_DEFS_TO_VESSEL","No")
	DATE_REPLIED_TO_SIRE	=	SETDATA("DATE_REPLIED_TO_SIRE","DATE_REPLIED_TO_SIRE","No","Date")
	'DATE_SIRE_TO_VESSEL	=	SETDATA("DATE_SIRE_TO_VESSEL","DATE_SIRE_TO_VESSEL","No")
	EXPIRY_DATE				=	SETDATA("EXPIRY_DATE","EXPIRY_DATE","No","Date")
	OFFICE_ADVISED_DATE		=	SETDATA("OFFICE_ADVISED_DATE","OFFICE_ADVISED_DATE","No","Date")
	REQUEST_DATE			=	SETDATA("REQUEST_DATE","REQUEST_DATE","No","Date")
	SIRE_RECD_DATE			=	SETDATA("SIRE_RECD_DATE","SIRE_RECD_DATE","No","Date")

	
	If Request("v_request_office") <> "" Then
		TECH_ALERTED		=	"YES"
		TECH_ALERT_DATE		=	now
	Else
		TECH_ALERTED		=	IIF(Trim(Request("TECH_ALERTED")) = "", null, Trim(Request("TECH_ALERTED")))
		TECH_ALERT_DATE		=	IIF(Trim(Request("TECH_ALERT_DATE")) = "", null, Trim(Request("TECH_ALERT_DATE")))
	End If

	TECH_REPLY_DATE			=	SETDATA("TECH_REPLY_DATE","TECH_REPLY_DATE","No","Date")
	VESSEL_ADVISED_DATE		=	SETDATA("VESSEL_ADVISED_DATE","VESSEL_ADVISED_DATE","No","Date")
	
	EXPENCES				=	SETDATA("EXPENCES","EXPENCES","No","Number")
	EXPENCES_IN_USD			=	SETDATA("EXPENCES_IN_USD","EXPENCES_IN_USD","No","Number")
	PO_NUMBER				=	SETDATA("PO_NUMBER","PO_NUMBER","No","Char")
	
	AGENT_ID				=	SETDATA("AGENT_ID","AGENT_ID","No","Number")
	INSPECTOR_ID			=	SETDATA("INSPECTOR_ID","INSPECTOR_ID","No","Number")
	COMBINED				=	SETDATA("COMBINED","COMBINED","No","Char")
	'REJECTION_REASON		=	SETDATA("REJECTION_REASON","REJECTION_REASON","No","Char")
	REMARKS					=	SETDATA("REMARKS","REMARKS","No","Char")
	TECH_DECLINED_REASON	=	SETDATA("TECH_DECLINED_REASON","TECH_DECLINED_REASON","No","Char")
	CONFIRMED_OR_REJECTED	=	SETDATA("CONFIRMED_OR_REJECTED","CONFIRMED_OR_REJECTED","No","Char")
	OPERATION				=	SETDATA("OPERATION","OPERATION","No","Char")
	TECH_STATUS				=	SETDATA("TECH_STATUS","TECH_STATUS","No","Char")
	local_CURRENCY			=	SETDATA("local_CURRENCY","local_CURRENCY","No","Char")
	DEFICIENCY_RECD			=	SETDATA("DEFICIENCY_RECD","DEFICIENCY_RECD","No","Char")
	TECH_DEPT_REPLIED		=	SETDATA("TECH_DEPT_REPLIED","TECH_DEPT_REPLIED","No","Char")
	INSPECTION_PORT			=	SETDATA("INSPECTION_PORT","INSPECTION_PORT","No","Char")
	SUPTD_ATTENDED			=	SETDATA("SUPTD_ATTENDED","SUPTD_ATTENDED","No","Char")
	TECH_PIC				=	SETDATA("TECH_PIC","TECH_PIC","No","Char")

	status					=	SETDATA("status","Status","No","Char")
	detention				=	SETDATA("detention","Detention","No","Char")
	insp_type				=	SETDATA("insp_type","Inspection Type","Yes","Char")
	basis_sire_moc			=	SETDATA("basis_sire_moc_name","Basis SIRE MOC name","No","Char")
	BASIS_SIRE				=	SETDATA("BASIS_SIRE","Basis SIRE MOC inspection id","No","Number")
	OCIMF_REPORT_NUMBER		=	SETDATA("OCIMF_REPORT_NUMBER","OCIMF Report number","No","Char")

	Is_SIre			=	SETDATA("IS_SIRE","Is SIRE","Yes","Char")

	
	if (v_no_error <> "true") then
		Response.Clear
		Response.Redirect "ins_request_entry.asp?v_ins_request_id="&request_id&"&v_message="&v_message
	end if		
	
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
	
	Idval = Request("request_id")
	
	rs.open "Select * from MOC_INSPECTION_REQUESTS where request_id=" & Idval

	if rs.eof then rs.addnew

	rs("VESSEL_CODE")= VESSEL_CODE
	rs("MOC_ID") = MOC_ID
	rs("insp_status") = insp_status
	rs("inspection_date") = inspection_date
	rs("agent_advised_date") = agent_advised_date
	rs("date_accepted")= date_accepted
	rs("date_confirm_reject")= date_confirm_reject
	rs("date_replied_to_sire")= date_replied_to_sire
	rs("expiry_date")= expiry_date
	rs("request_date")= request_date
	rs("sire_recd_date")= sire_recd_date
	rs("tech_reply_date")= tech_reply_date
	rs("vessel_advised_date")= vessel_advised_date
	rs("expences_in_usd")= expences_in_usd
	rs("po_number")= PO_NUMBER
	rs("expences")= expences 
	rs("agent_id")= agent_id 
	rs("inspector_id")= inspector_id
	rs("local_currency")= local_currency
	rs("inspection_remarks") = remarks
	'rs("rejection_reason")= rejection_reason
	rs("tech_declined_reason")= tech_declined_reason
	rs("operation")= operation
	rs("tech_status")= tech_status
	rs("tech_alerted") = tech_alerted
	rs("tech_alert_date") = tech_alert_date
	rs("inspection_port")= inspection_port
	rs("tech_pic")= tech_pic
	rs("last_modified_by")= USER
	rs("last_modified_date")= now
	rs("status")= status
	rs("detention")= detention
	rs("insp_type")= insp_type
	rs("basis_sire")= basis_sire
	rs("basis_sire_moc")= basis_sire_moc
	rs("OCIMF_REPORT_NUMBER")= OCIMF_REPORT_NUMBER
	rs("IS_Sire")= IS_Sire

	
	rs.Update
	
	rs.Resync
	
	Idval = rs("request_id")
	
	v_message = "MOC Inspection details Updated Successfully"


	
	'insert records for uploaded documents
	dim i,sClientScript,obj
	i=1
	rsFile.Open "Select * from moc_documents where doc_type='INSPECTION' and parent_id=" & Idval
	for each obj in Request.Form("txtDocID")
		if obj="" then
			rsFile.AddNew
			rsFile("doc_type") = "INSPECTION"
			rsFile("parent_id") = Idval
			rsFile("doc_name") = Request.Form("txtDocName")(i)
			rsFile("uploaded_by") = USER
			
			rsFile.Update
			rsFile.Resync
			
			rsFile("doc_path") = PathFromID(rsFile("doc_id"),Request.Form("txtDocPath")(i))
			rsFile.Update
			
			sClientScript = sClientScript & "TPMDnDCommonCtrl1.UploadFile """ & Request.Form("txtDocPath")(i) & """,""" & MOC_PATH & rsFile("doc_path") & """" & vbcrlf
		end if
		i = i+1
	next
	
	connObj.Close
	set connObj=nothing

	If Request("v_save_close") = "Save and Close" Then
%>
<SCRIPT LANGUAGE="vbscript">
	function window_onload
		DoUpload
		
		on error resume next
		self.parent.opener.document.form1.action = "ins_request_maint.asp"

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
		
		self.location.href = "ins_request_entry.asp?v_ins_request_id=<%=IdVal%>&v_message=" & v_message
	end function
	
	sub DoUpload
		<%=sClientScript%>
	end sub
</SCRIPT>	
<%		'Response.Redirect "ins_request_entry.asp?v_ins_request_id="&request_id&"&v_message="&v_message
	End If
%>
