<%@ Language=VBScript %>
<!--#include file="common_dbconn.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT LANGUAGE="Javascript" >
function fun(to,subject,body)
{
	alert(document.applets[0].MailFormat);
	//mail.displayMailClient("suresh_kumar_v@yahoo.com","","","Hello",body);
	
	document.applets[0].displayMailClient(to,"","",subject,body);
}
</SCRIPT>
</HEAD>
<BODY>
<OBJECT id=mail style="LEFT: 0px; TOP: 0px" name=mail  codebase="MailClient.CAB" classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" 
	width=0 height=0 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="26">
	<PARAM NAME="_ExtentY" VALUE="26">
</OBJECT>
<%
	'v_ins_request_id = "10000087"
	'v_ins_request_id = "10000042"
	'v_ins_request_id = "10000089"
	'v_ins_request_id = "10000043"
	v_ins_request_id = "10000161"
	
	strSqlVessMail = "select request_id, moc_id, moc_short_name, moc_full_name, inspection_date, "
	strSqlVessMail = strSqlVessMail & "section, deficiency, status, inspection_port, vessel_code, vessel_name, "
	strSqlVessMail = strSqlVessMail & "vessel_short_name, fleet_code, "
	strSqlVessMail = strSqlVessMail & "to_char(inspection_date, 'dd-Mon-yyyy') inspection_date_disp from moc_vwr_list_of_observations "
	strSqlVessMail = strSqlVessMail & "where request_id = " & v_ins_request_id & " "
	strSqlVessMail = strSqlVessMail & "order by section "
	
	Set rsObjVessMail = connObj.Execute(strSqlVessMail)
	
	If rsObjVessMail.EOF = False Then
		VessMailBody = "Dear Capt. ..........,\n\n"
		VessMailBody = VessMailBody & "Further to Vetting Inspection by:\t" & rsObjVessMail("moc_short_name") & "\n"
		VessMailBody = VessMailBody & "At Port:\t\t" & rsObjVessMail("inspection_port") & "\n"
		VessMailBody = VessMailBody & "Inspected on:\t" & rsObjVessMail("inspection_date_disp") & "\n"
		VessMailBody = VessMailBody & "\n"
		VessMailBody = VessMailBody & "Listed below are deficiencies noted by the inspector :-\n"
		VessMailBody = VessMailBody & "\n"

		v_send_to = rsObjVessMail("vessel_name")
		
		srNo = 1
		While Not rsObjVessMail.EOF
			VessMailBody = VessMailBody & srNo & "\t" & rsObjVessMail("deficiency") & "\n"

			srNo = srNo + 1
			rsObjVessMail.MoveNext
		Wend

		VessMailBody = VessMailBody & "\n\n"
		VessMailBody = VessMailBody & "Plese advise in brief, the actions / rectifications taken on the above deficiencies."
		VessMailBody = VessMailBody & "\n\n"
		VessMailBody = VessMailBody & "Tks & Brgds\n"

		VessMailBody = Replace(VessMailBody, Chr(13) & Chr(10), "\n")
		VessMailBody = Replace(VessMailBody, Chr(34), "")
		VessMailBody = Replace(VessMailBody, "'", "")
		
		Response.Write "Testing - <P>" & VessMailBody & "<BR>"

	Else	'no deficiencies exist

	End If
%>
<INPUT TYPE="button" NAME="v_mail" VALUE="Send Mail" OnClick="javascript:fun('','','<% =VessMailBody %>');">
</BODY>
</HTML>
