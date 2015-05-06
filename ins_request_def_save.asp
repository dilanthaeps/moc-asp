<%
'===========================================================================
'	Template Name	:	MOC inspection  deficiency Save Screen
'	Template Path	:	...0/ins_request_def_save.asp
'	Functionality	:	Save new MOC Inspection deficiency details or update the existing details
'	Called By		:	../ins_request_def_entry.asp
'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
'	Update History	:	12th September 2002
'						1.
'						2.
'===========================================================================
option explicit
%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
Function SQE(text_with_single_quote)
    If Not IsNull(text_with_single_quote) Then
		SQE = text_with_single_quote 'Trim(Replace(text_with_single_quote, "'", "''"))
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
			if data_type="Number" then
				SETDATA=null
			end if
			if v_mandatory="Yes" then
			v_no_error = "false"
			v_message = v_message&v_comment&" cannot be empty<br>"
			end if
	end if
End Function
</SCRIPT>
<%	

	dim v_message,v_no_error,deficiency_id,section,deficiency,reply,status,sort_order
	dim risk_factor,action_code,action_description,viq_dont_count
	dim request_id
	v_message=""
	
	v_no_error="true" 
	deficiency_id			=	SETDATA("deficiency_id","Deficiency ID","No","Number")
	section					=	setdata("section","Section","Yes","Char")
	action_code				=	setdata("action_code","Action Code","No","Char")
	action_description		=	setdata("action_description","Action Description","No","Char")
	request_id				=   SETDATA("request_id","Request ID","Yes","Number")
	deficiency				=	SETDATA("deficiency","Deficiency ","Yes","Char")
	reply					=	SETDATA("reply","Reply ","No","Char")
	status					=	setdata("status","Status","Yes","Char")
	sort_order				=	setdata("sort_order","Sort Order","No","Number")
	risk_factor				=	setdata("risk_factor","Risk","Yes","Char")
	
	'for VIQ sections only -- enter the input for viq count -- else 0
	if section = "8.75" OR  section = "9.27" OR section = "11.3" then
		viq_dont_count =	setdata("viq_dont_count","VIQ Dont count","No","Number")
	else
		viq_dont_count = 0
	end if
	


	
	if (v_no_error <> "true") then
		Response.Redirect "ins_request_def_maint.asp?v_ins_request_id="&request_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	end if		
	
	request_id = Request("request_id")
	deficiency_id	= request("deficiency_id")
	if deficiency_id="" then deficiency_id=-1
	
	dim rs
	set rs=Server.CreateObject("ADODB.Recordset")
	with rs
		.CursorLocation = adUseClient
		.CursorType = adOpenDynamic
		.LockType = adLockOptimistic
		set .ActiveConnection = connObj
	end with
	rs.Open "Select * from moc_deficiencies where deficiency_id=" & deficiency_id

	if rs.EOF then rs.AddNew
	
	rs("REQUEST_ID") = sqe(request_id)
	rs("section") = sqe(section)
	rs("action_code") = sqe(action_code)
	rs("action_description") = sqe(action_description)
	rs("deficiency") = sqe(deficiency)
	rs("reply") = sqe(reply)
	rs("status") = sqe(status)
	rs("sort_order") = sqe(sort_order)
	rs("risk_factor") = risk_factor
	rs("CREATE_DATE") = now
	rs("CREATED_BY") = ""
	rs("LAST_MODIFIED_DATE") = now
	rs("LAST_MODIFIED_BY") = ""
	rs("VIQ_DONT_COUNT") = viq_dont_count
	
	rs.Update
	
	if deficiency_id = -1 then
		rs.Resync
		deficiency_id = rs("deficiency_id")
	end if
	
	rs.Close
	

	'rsObj.Close
	set rsObj=nothing
	connObj.Close
	set connObj=nothing
	'Response.Redirect "ins_request_def_maint.asp?v_ins_request_id="&request_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	
	if request("submit4") = "Save and Close" then
		Response.Redirect "ins_request_def_entry.asp?v_ins_request_id="&request_id&"&v_def_id="&deficiency_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")&"&v_close=true"
	else
		Response.Redirect "ins_request_def_entry.asp?v_ins_request_id="&request_id&"&v_def_id="&deficiency_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	end if
%>   
