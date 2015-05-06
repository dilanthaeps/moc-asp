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
</SCRIPT>

<%	
	Response.Buffer = false
	'===========================================================================
	'	Template Name	:	MOC Save Screen
	'	Template Path	:	...0/sytem_parameter_save.asp
	'	Functionality	:	Save new MOC details or update the existing details
	'	Called By		:	../MOC_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:	21st August 2002
	'						1.
	'						2.
	'===========================================================================
	
	 
	if request("short_name")<>"" then
		short_name=request("short_name")
	else
		short_name=null
		v_no_error = false
		v_message = v_message&"Short Name cannot be empty<br>"
	end if
	
	if request("full_name")<>"" then
		full_name=request("full_name")
	else
		full_name=null
		v_no_error = false
		v_message = v_message&"Full Name cannot be empty<br>"
	end if
	
	if request("address")<>"" then
		address= request("address")
	else
		address=null
	end if
	
	if request("telephone")<>"" then
		telephone= request("telephone") 
	else
		telephone=null
	end if
	
	if request("fax_no")<>"" then
		fax_no= request("fax_no") 
	else
		fax_no=null
	end if
	
	if request("email")<>"" then
		email= request("email") 
	else
		email=null
	end if
	
	if request("pic")<>"" then
		pic= request("pic") 
	else
		pic=null
	end if
	
	if request("remarks")<>"" then
		remarks= request("remarks") 
	else
		remarks=null
	end if
	
	if request("entry_type")<>"" then
		entry_type = request("entry_type") 
	else
		entry_type = null
	end if
	
	if v_no_error <> "" then
'	Response.Redirect "moc_maint.asp?v_message="&v_message
		If Request("v_child_opener") = "yes" Then
%>   
			<SCRIPT LANGUAGE="JavaScript">
				self.parent.opener.document.form1.action = self.parent.opener.location;
				self.close();
				self.parent.opener.document.form1.target = "";
				self.parent.opener.document.form1.submit();
			</SCRIPT>
<%
		Else
%>
			<SCRIPT LANGUAGE="JavaScript">
				self.parent.opener.document.form1.action = "moc_maint.asp?v_message=<% =v_message %>";
				//alert(self.parent.opener.location);
				self.close();
				self.parent.opener.document.form1.target = "";
				self.parent.opener.document.form1.submit();
			</SCRIPT>
<%
		End If
	end if		
			
	Dim Idval
	Idval = Request("mode")
	if Idval = "" then
		strSql= "INSERT INTO  MOC_MASTER"
		strSql = strSql & " ( MOC_ID , SHORT_NAME , FULL_NAME , ADDRESS , TELEPHONE "
		strSql = strSql & " , FAX_NO , EMAIL , PIC , REMARKS , CREATE_DATE , CREATED_BY ," 
		strSql = strSql & " LAST_MODIFIED_DATE , LAST_MODIFIED_BY, ENTRY_TYPE ) values "
		strSql = strSql & "( 0,'" & sqe(short_name) & "','" & sqe(full_name)
		strSql = strSql & "','" & sqe(address) & "','" & sqe(telephone) & "','" & sqe(fax_no) & "','"
		strSql = strSql & sqe(email) & "','" & sqe(pic) & "','" & sqe(remarks) & "',sysdate,'" & moc_user  
		strSql = strSql & "',sysdate,'" & moc_user & "','" & entry_type & "')" 
		v_message = "MOC details Created Successfully"
	else
		strSql = "UPDATE  MOC_MASTER  A SET "
		strSql = strSql &"  A.SHORT_NAME = '"& sqe(short_name) &"'"
		strSql = strSql &" , A.FULL_NAME = '"& sqe(full_name) &"'"
		strSql = strSql &" , A.ADDRESS = '"& sqe(address) &"'"
		strSql = strSql &" , A.TELEPHONE = '"& sqe(telephone) &"'"
		strSql = strSql &" , A.FAX_NO = '"& sqe(fax_no) &"'"
		strSql = strSql &" , A.EMAIL = '"& sqe(email) &"'"
		strSql = strSql &" , A.PIC = '"& sqe(pic) &"'"
		strSql = strSql &" , A.REMARKS = '"& sqe(remarks) &"'"
		strSql = strSql &" , A.LAST_MODIFIED_DATE = sysdate "
		strSql = strSql &" , A.LAST_MODIFIED_BY = '"& sqe(moc_user) &"'"
		strSql = strSql &" , A.ENTRY_TYPE = '"& sqe(entry_type) &"'"
		strSql = strSql &" where a.moc_id=" & request("moc_id")
		v_message = "MOC details Updated Successfully"
	end if
	'Response.Write strSql
	Set rsObj = connObj.Execute(strSql)
	'rsObj.Close
	set rsObj=nothing
	connObj.Close
	set connObj=nothing
	'Response.Redirect "moc_maint.asp?v_message="&v_message

	If Request("v_child_opener") = "yes" Then
%>   
<SCRIPT LANGUAGE="JavaScript">
	//self.parent.opener.document.form1.action = self.parent.opener.location;
	self.parent.opener.document.form1.action = "";
	self.parent.opener.document.form1.target = "";
	self.parent.opener.history.go(0);
	self.close();
	//self.parent.opener.document.form1.target = "";
	//self.parent.opener.document.form1.submit();
</SCRIPT>
<%
	Else
%>
<SCRIPT LANGUAGE="JavaScript">
	//self.parent.opener.document.form1.action = "moc_maint.asp?v_message=<% =v_message %>";
	//alert(self.parent.opener.location);
	self.parent.opener.document.form1.action = "";
	self.parent.opener.document.form1.target = "";
	self.parent.opener.history.go(0);
	self.close();
	//self.parent.opener.document.form1.submit();
</SCRIPT>
<%
	End If
%>