<%@ Language=VBScript %>
<!--#include file="common_dbconn.asp"-->
<%
	Function IIF(expr, trueValue, falseValue)
		If expr Then
			IIF = trueValue
		Else
			IIF = falseValue
		End If
	End Function

	If Request("v_tech_reply") <> "" Then
		strSql = "update moc_inspection_requests "
		strSql = strSql & "set tech_status = '" & Trim(Request("reply_status")) & "', "
		strSql = strSql & "tech_reply_date = sysdate, "
		strSql = strSql & "tech_declined_reason = '" & Replace(Trim(Request("remarks")), "'", "''") & "' "
		strSql = strSql & "where request_id = " & Trim(Request("v_ins_request_id"))

		'Response.Write strSql & "<BR>"
		'Response.End		

		connObj.Execute(strSql)
		Response.Write "Technical reply has been updated successfully !"
	End If
%>
<SCRIPT LANGUAGE="JavaScript">
	self.close();
</SCRIPT>