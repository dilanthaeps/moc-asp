<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<%
dim v_message,v_no_error,remark_id,request_id,subject,remarks,remark_pic,remark_target_date,remark_status

%>
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
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
		if data_type="Number" then
			SETDATA="null"
		end if
		if v_mandatory="Yes" then
			v_no_error = "false"
			v_message = v_message&v_comment&" cannot be empty<br>"
		end if
	end if
End Function
</SCRIPT>
<%	
	Response.Buffer = false
	'===========================================================================
	'	Template Name	:	MOC inspection  remark Save Screen
	'	Template Path	:	...0/ins_remark_save.asp
	'	Functionality	:	Save new MOC Inspection remark details or update the existing details
	'	Called By		:	../ins_request_remark_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:	11th September 2002
	'						1.
	'						2.
	'===========================================================================
	v_message=""

	
	v_no_error="true" 
	remark_id				=	SETDATA("remark_id","Remark ID","No","Number")
	request_id				=   SETDATA("request_id","Request ID","Yes","Number")
	subject					=	SETDATA("subject","Subject ","Yes","Char")
	remarks					=	SETDATA("remarks","Remarks ","Yes","Char")
	remark_pic				=	SETDATA("remark_pic","Remarks Person Incharge ","Yes","Char")
	remark_target_date		=	setdata("remark_target_date","Remark Target Date","No","Date")
	remark_status			=	setdata("remark_status","Remark Status","Yes","Char")
	
	 
	
	if (v_no_error <> "true") then
		Response.Redirect "ins_request_remark_maint.asp?v_ins_request_id="&request_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	end if		
	
	request_id = Request("request_id")
	remark_id	= request("remark_id")
	
	if (remark_id = "0" or remark_id="" or remark_id=null) then
		'strSeq="select 	SEQ_MOC_REQUEST_remarks.nextval remark_id from dual 	"
		'set rsObj_seq=connObj.execute(strSeq)
		'while not rsObj_seq.eof
		'	remark_id = rsObj_seq("remark_id")
		'	rsObj_seq.movenext
		'wend
		
		'updated by trigger
		remark_id = 0
		
		strSql= "INSERT INTO  MOC_REQUEST_remarkS "
		strSql= strSql & "( REMARK_ID,REQUEST_ID , SUBJECT, REMARKS ,REMARK_PIC,"
		strSql= strSql & "  REMARK_TARGET_DATE, REMARK_STATUS "
		strSql= strSql & ",CREATE_DATE "
		strSql= strSql & ",CREATED_BY "
		strSql= strSql & ",LAST_MODIFIED_DATE "		
		strSql= strSql & ",LAST_MODIFIED_BY "
		strSql= strSql  & ")"
		
		strSql = strSql & "  values "
		
		strSql = strSql & "( " 
		strSql = strSql &   remark_id 
		strSql = strSql & "," & sqe(request_id)											'Number Field
		strSql = strSql & ",'" & sqe(subject) & "'"
		strSql = strSql & ",'" & sqe(remarks) & "'"
		strSql = strSql & ",'" & sqe(remark_pic) & "'"
		strSql= strSql  & ",to_date('"& remark_target_date & "','DD-MON-YYYY')"		'Date Field
		strSql = strSql & ",'" & sqe(remark_status) & "'"
		strSql= strSql  & ",sysdate"		'Date Field
		strSql = strSql & ",'" & USER & "'"
		strSql= strSql  & ",sysdate"		'Date Field
		strSql = strSql & ",'" & USER & "'"
		
		strSql= strSql  & ")"

		v_message = "MOC Inspection Remarks Created Successfully"
	else
		strSql = "UPDATE  MOC_REQUEST_REMARKS  A SET "
		strSql = strSql &"   A.SUBJECT = '"& sqe(SUBJECT) &"'"
		strSql = strSql &" , A.REMARKS = '"& sqe(REMARKS) &"'"									'Varchar Field
		strSql = strSql &" , A.REMARK_status = '"& sqe(REMARK_STATUS) &"'"
		strSql = strSql &" , A.REMARK_pic = '"& sqe(REMARK_pic) &"'"
		strSql = strSql &" , A.REMARK_TARGET_DATE = to_date('"& sqe(REMARK_TARGET_DATE) &"','DD-MON-YYYY')"	' Date Field
		strSql = strSql &" , A.last_modified_by='" & USER &"'"
		strSql = strSql &" , A.last_modified_date=sysdate"		
		strSql = strSql &" where a.REmark_ID=" & request("remark_id")
		v_message = "MOC Inspection Remark details Updated Successfully"
	end if
	'Response.Write strSql
	Set rsObj = connObj.Execute(strSql)
	'rsObj.Close
	set rsObj=nothing
	connObj.Close
	set connObj=nothing
	if request("submit4") = "Save and Close" then
	Response.Redirect "ins_request_remark_entry.asp?v_ins_request_id="&request_id&"&v_remark_id="&remark_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")&"&v_close=true"
	else
	Response.Redirect "ins_request_remark_entry.asp?v_ins_request_id="&request_id&"&v_remark_id="&remark_id&"&v_message="&v_message&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	end if
%>   
