<%@ Language=VBScript %>
<%Response.Expires=-1%>
<!--#include file="common_dbconn.asp"-->
<HTML>
<HEAD>
<TITLE>Technical Reply</TITLE>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<H4>Technical Reply</H4>
<H4>Vessel Name: <%=Request("v_vessel_name")%>&nbsp;&nbsp;&nbsp;MOC: <%=Request("v_moc_name")%></H4>
<%
	strReplyBody = "select mir.moc_id, mm.short_name moc_short_name, mm.full_name moc_full_name, v.vessel_name, "
	strReplyBody = strReplyBody & "mir.inspection_port, to_char(inspection_date, 'dd-Mon-yyyy') inspection_date_disp, "
	strReplyBody = strReplyBody & "um.user_name, "
	strReplyBody = strReplyBody & "moc_fn_sys_para_desc('Coordinator Email', 'Coordinator') coordinator_email, "
	strReplyBody = strReplyBody & "moc_fn_sys_para_desc('Coordinator Email1', 'Coordinator') coordinator_email1 "
	strReplyBody = strReplyBody & "from moc_inspection_requests mir, moc_master mm, "
	strReplyBody = strReplyBody & "wls_vw_vessels_new v, wls_user_master um "
	strReplyBody = strReplyBody & "where mir.moc_id = mm.moc_id "
	strReplyBody = strReplyBody & "and mir.vessel_code = v.vessel_code "
	strReplyBody = strReplyBody & "and mir.tech_pic = um.user_id(+) "
	strReplyBody = strReplyBody & "and mir.request_id = " & Trim(Request("v_ins_request_id"))
stop			
	'Response.Write strReplyBody & "<BR>"
	Set rsObjReplyBody = connObj.Execute(strReplyBody)

	If not rsObjReplyBody.EOF Then
		v_vessel_name = rsObjReplyBody("vessel_name")
		v_inspection_date_disp = rsObjReplyBody("inspection_date_disp")
		v_inspection_port = rsObjReplyBody("inspection_port")
		v_moc_short_name = rsObjReplyBody("moc_short_name")
		v_user_name = rsObjReplyBody("user_name")
		v_addressee = rsObjReplyBody("coordinator_email") '& ";" & rsObjReplyBody("coordinator_email1")
	End If
				
	rsObjReplyBody.Close
	Set rsObjReplyBody = Nothing

	ReplyBody = ReplyBody & "Vessel: " & v_vessel_name & "\n"
	ReplyBody = ReplyBody & "MOC: " & v_moc_short_name & "\n"
	ReplyBody = ReplyBody & "Inspection Port: " & v_inspection_port & "\n"
	ReplyBody = ReplyBody & "Inspection Date: " & v_inspection_date_disp & "\n"

	Regds = "\nRegards"

	Regds = Regds & "\n" & v_user_name
				

	ReplyBody = Replace(ReplyBody, Chr(13) & Chr(10), "\n")
	ReplyBody = Replace(ReplyBody, Chr(34), "")
	ReplyBody = Replace(ReplyBody, "'", "")
%>
<FORM NAME="v_form" METHOD="post" ACTION="technical_reply_save.asp" OnSubmit="return validate('<% =v_addressee %>', '<% =ReplyBody %>', '<% =Regds %>');">
<OBJECT id=mail style="LEFT: 0px; TOP: 0px" name=mail  codebase="MailClient.CAB" classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" 
	width=0 height=0 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="26">
	<PARAM NAME="_ExtentY" VALUE="26">
</OBJECT>
<INPUT TYPE="hidden" NAME="v_ins_request_id" VALUE="<% =Trim(Request("v_ins_request_id")) %>">
<TABLE WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="1">
	<TR HEIGHT="30pt">
		<TD WIDTH="15%" CLASS="tableheader">Status</TD>
		<TD WIDTH="85%" CLASS="tabledata">&nbsp;
			<SELECT NAME="reply_status" STYLE="width:75pt">
<%
				strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
				strSql = strSql & "from moc_system_parameters "
				strSql = strSql & "where parent_id = 'Tech_Status' "
				strSql = strSql & "order by sort_order "

				Response.Write strSql
				Set rsObj = connObj.Execute(strSql)

				While Not rsObj.EOF

					Response.Write "<OPTION VALUE='" & rsObj("sys_para_id") & "'>"
					Response.Write rsObj("para_desc")
					Response.Write "</OPTION>"
					rsObj.MoveNext
				Wend

				rsObj.Close
				Set rsObj = Nothing
%>
		</TD>		
	</TR>
	<TR HEIGHT="20pt">
		<TD WIDTH="15%" CLASS="tableheader">Remarks</TD>
		<TD WIDTH="85%" CLASS="tabledata">&nbsp;
			<textarea name="remarks" cols="47" rows="6"></textarea>
		</TD>		
	</TR>
</TABLE>
<BR>
<INPUT TYPE="submit" NAME="v_tech_reply" VALUE="Reply">
<INPUT TYPE="button" NAME="v_close" VALUE="Close without Reply" OnClick="self.close()">
</FORM>
</BODY>
<SCRIPT LANGUAGE="JavaScript">
function validate(addressee, msgBody, regds)
{
	if (document.v_form.remarks.value == "")
	{
		alert("Please enter remarks !");
		document.v_form.remarks.focus();
		return false;
	}

	//alert(addressee);

	subject = "TECHINCAL CONFIRMATION - <%=v_vessel_name%> - <%=v_moc_short_name%> @ <%=v_inspection_port%> on/around <%=v_inspection_date_disp%>"
	
	body = "== TECHINCAL CONFIRMATION ==\n\n";
	body += msgBody + "\n\n";
	body += "Status: " + document.v_form.reply_status(document.v_form.reply_status.selectedIndex).text + "\n\n";
	body += "Remarks: " + document.v_form.remarks.value + "\n\n";
	body += regds;

	//document.applets[0].displayMailClient(addressee,"Technical Superintendents; Operations Superintendents","",subject,body);
	document.applets[0].displayMailClient(addressee,"","",subject,body);
	//return false;

	return true;
}
</SCRIPT>

</HTML>
