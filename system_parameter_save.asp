	<!--#include file="common_dbconn.asp"-->
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
Function SQE(text_with_single_quote)
    If Not IsNull(text_with_single_quote) Then
		SQE = Trim(Replace(text_with_single_quote, "'", "''"))
    Else
		SQE = text_with_single_quote
    End If
End Function

</SCRIPT>

<%	
	Response.Buffer = false
	'===========================================================================
	'	Template Name	:	System Parameter Save Screen
	'	Template Path	:	...0/sytem_parameter_save.asp
	'	Functionality	:	Save new system parameter details or update the existing details
	'	Called By		:	../sytem_parameter_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	
	if request("sys_para_id")<>"" then
		sys_para_id=request("sys_para_id")
	else
		sys_para_id=null
		v_no_error = false
		v_message = v_message&"System Parameter ID cannot be empty<br>"
	end if   
	 
	if request("para_desc")<>"" then
		para_desc=request("para_desc")
	else
		para_desc=null
		v_no_error = false
		v_message = v_message&"System Parameter Description cannot be empty<br>"
	end if
	
	if request("parent_id")<>"" then
		parent_id=request("parent_id")
	else
		parent_id=null
		v_no_error = false
		v_message = v_message&"System Parent ID cannot be empty<br>"
	end if
	
	if request("remarks")<>"" then
		remarks= request("remarks")
	else
		remarks=null
	end if
	
	if request("related_asp_pages")<>"" then
		related_asp_pages= request("related_asp_pages") 
	else
		related_asp_pages=null
		v_no_error = false
		v_message = v_message&"System Related ASP cannot be empty<br>"
	end if
	
	if request("sort_order")<>"" then
		sort_order= request("sort_order") 
	else
		sort_order=null
	end if
	if v_no_error <> "" then
	'Response.Redirect "system_parameter_maint.asp?v_message="&v_message
%>
<SCRIPT LANGUAGE="JavaScript">
	self.parent.opener.document.form1.action = "system_parameter_maint.asp?v_message=<% =v_message %>";
	//alert(self.parent.opener.document.v_form.action);
	self.close();
	self.parent.opener.document.form1.target = "";
	self.parent.opener.document.form1.submit();
</SCRIPT>
<%
	end if		
			
	Dim Idval
	Idval = Request("mode")
	if Idval = "" then
		strSql = "INSERT INTO moc_system_parameters( SYS_PARA_ID , PARA_DESC , PARENT_ID ,REMARKS , CREATE_DATE , LAST_MODIFIED_DATE, RELATED_ASP_PAGES ,  SORT_ORDER ) values "
		strSql = strSql & "('" & sqe(sys_para_id) & "', '" & sqe(para_desc) & "','"&sqe(parent_id)&"','"&sqe(remarks)&"', sysdate, sysdate,'"&sqe(related_asp_pages)&"',"&sort_order&")"
		v_message = "System Parameter Created Successfully"
	else
		strSql = "Update moc_system_parameters set para_desc='"&sqe(para_desc)&"',parent_id='"&sqe(parent_id)&"',remarks='"&sqe(remarks)&"',last_modified_date=sysdate,related_asp_pages='"&sqe(related_asp_pages)&"',sort_order="& sort_order&" where sys_para_id="
		strSql = strSql & "'" & sys_para_id & "'"
		v_message = "System Parameter Updated Successfully"
	end if
	'Response.Write strSql
	Set rsObj = connObj.Execute(strSql)
	'rsObj.Close
	set rsObj=nothing
	connObj.Close
	set connObj=nothing
	'Response.Redirect "system_parameter_maint.asp?v_message="&v_message
%>   
<SCRIPT LANGUAGE="JavaScript">
	self.parent.opener.document.form1.action = "system_parameter_maint.asp?v_message=<% =v_message %>";
	//alert(self.parent.opener.document.v_form.action);
	self.close();
	self.parent.opener.document.form1.target = "";
	self.parent.opener.document.form1.submit();
</SCRIPT>
