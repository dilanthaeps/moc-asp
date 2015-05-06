<!--#include file="common_dbconn.asp"-->
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
		'Response.Write name_of_the_form_variable &"-"& request(name_of_the_form_variable)& "-"&sqe( request(name_of_the_form_variable))&"<br>"
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
	'	Template Name	:	fbm Save Screen
	'	Template Path	:	...0/fbm_save.asp
	'	Functionality	:	Save new fbm details or update the existing details
	'	Called By		:	../fbm_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:	25th September 2002
	'						1.
	'						2.
	'===========================================================================
	v_message=""

	
	v_no_error="true" 
	from_dept				=	SETDATA("from_dept","From Department","Yes","Char")
	date_sent				=   SETDATA("date_sent","Date Sent","Yes","Date")
	primary_category		=	SETDATA("primary_category","Primary Category ","Yes","Char")
	secondary_category		=	SETDATA("secondary_category","Secondary Category ","Yes","Char")
	msg_subject				=	setdata("msg_subject","Message Subject","Yes","Char")
	msg_body				=	setdata("msg_body","Message Body","Yes","Char")
	status					=	setdata("status","Status","Yes","Char")
	
	
	if (v_no_error <> "true") then
	Response.Redirect "fbm_maint.asp?v_fbm_id="&fbm_id&"&v_message="&v_message
	end if		
	
	v_fbm_id = Request("v_fbm_id")
	
	if (v_fbm_id = "0" or v_fbm_id="" or isnull(v_fbm_id)) then
		strSeq="select 	SEQ_wls_fbm.nextval fbm_id from dual 	"
		set rsObj_seq=connObj.execute(strSeq)
		while not rsObj_seq.eof
			v_fbm_id = rsObj_seq("fbm_id")
			rsObj_seq.movenext
		wend
		strSql= "INSERT INTO  wls_fbm "
		strSql= strSql & "( fbm_id, from_dept, date_sent, "
		strSql= strSql & " primary_category, secondary_category, msg_subject, msg_body, "
		strSql= strSql & " status, attachement,CREATE_DATE "
		strSql= strSql & ",CREATED_BY "
		strSql= strSql & ",LAST_MODIFIED_DATE "		
		strSql= strSql & ",LAST_MODIFIED_BY "
		strSql= strSql  & ")"
		
		strSql = strSql & "  values "
		
		strSql = strSql & "( " 
		strSql = strSql &   v_fbm_id 
		strSql = strSql & ",'" & sqe(from_dept)	& "'"										'Char Field
		strSql= strSql  & ",to_date('"& date_sent & "','DD-MON-YYYY')"		'Date Field
		strSql = strSql & ",'" & sqe(primary_category) & "'"
		strSql = strSql & ",'" & sqe(secondary_category) & "'"
		strSql = strSql & ",'" & sqe(msg_subject) & "'"
		strSql = strSql & ",'" & sqe(msg_body) & "'"
		strSql = strSql & ",'" & sqe(status) & "'"
		strSql = strSql & ",'" & sqe(attachement) & "'"
		strSql= strSql  & ",sysdate"		'Date Field
		strSql = strSql & ",'" & session("moc_user_id") & "'"
		strSql= strSql  & ",sysdate"		'Date Field
		strSql = strSql & ",'" & session("moc_user_id") & "'"
		
		strSql= strSql  & ")"

		v_message = "Fleet Broadcase Message Created Successfully"
	else
		strSql = "UPDATE  wls_fbm  A SET "
		strSql = strSql &"  A.from_dept = '"& sqe(from_dept) &"'"									'Varchar Field
		strSql = strSql &" , A.date_sent = to_date('"& sqe(date_sent) &"','DD-MON-YYYY')"	' Date Field
		strSql = strSql &" , A.primary_category = '"& sqe(primary_category) &"'"
		strSql = strSql &" , A.secondary_category = '"& sqe(secondary_category) &"'"
		strSql = strSql &" , A.msg_subject = '"& sqe(msg_subject) &"'"
		strSql = strSql &" , A.msg_body = '"& sqe(msg_body) &"'"
		strSql = strSql &" , A.status = '"& sqe(status) &"'"
		strSql = strSql &" , A.attachement = '"& sqe(attachement) &"'"
		strSql = strSql &" , A.last_modified_by='"&session("moc_user_id")&"'"
		strSql = strSql &" , A.last_modified_date=sysdate"		
		strSql = strSql &" where a.fbm_id=" & request("v_fbm_id")
		v_message = "Fleet Broadcast Message Updated Successfully"
		
	end if
	'Response.Write strSql
	Set rsObj = connObj.Execute(strSql)
	'rsObj.Close
	set rsObj=nothing
	connObj.Close
	set connObj=nothing
	if request("submit4") = "Save and Close" then
	Response.Redirect "fbm_entry.asp?v_fbm_id="&v_fbm_id&"&v_message="&v_message&"&v_close=true"
	else
	Response.Redirect "fbm_entry.asp?v_fbm_id="&v_fbm_id&"&v_message="&v_message
	end if
	
%>   
