<%Response.Buffer=true%>
<!-- Journal Details Begin-->
<!-- Journal Details End-->
<HEAD>
	<TITLE>Crystal ActiveX Report Viewer</TITLE>
</HEAD>
<OBJECT ID="CRViewer" CLASSID="CLSID:C4847596-972C-11D0-9567-00A0C9273C2A" WIDTH=100% HEIGHT=95% CODEBASE="http://webserver/viewer/activeXViewer/activexviewer.cab#Version=8,0,0,371">
	<PARAM NAME="EnableDrillDown" VALUE=1>
	<PARAM NAME="EnableExportButton" VALUE=1>
	<PARAM NAME="DisplayGroupTree" VALUE=1>
	<PARAM NAME="EnableGroupTree" VALUE=1>
	<PARAM NAME="EnableAnimationControl" VALUE=1>
	<PARAM NAME="EnablePrintButton" VALUE=1>
	<PARAM NAME="EnableRefreshButton" VALUE=1>
	<PARAM NAME="EnableSearchControl" VALUE=1>
	<PARAM NAME="EnableZoomControl" VALUE=1>
	<PARAM NAME="EnableSearchExpertButton" VALUE=0>
	<PARAM NAME="EnableSelectExpertButton" VALUE=0>
	<PARAM NAME="MorePrintEngineErrorMessages" VALUE=0>
</OBJECT>
<!--<INPUT TYPE="button" VALUE="Click" onclick="Page_Initialize('PieceMeal.rpt', '%3f', '20/10/2001', '01/01/2002', '%3f', '40', '%3f', '%3f', '%3f', '%3f', '%3f')">-->
<!--<INPUT TYPE="button" VALUE="Click" onclick="Page_Initialize('Ready_Spares.rpt', '256', '%3f', '%3f', '%3f', '%3f', '%3f', '%3f', '%3f', '%3f', '%3f')" id=button1 name=button1><br>-->
<!--#include file="common_dbconn.asp"-->
<%

	'Journal Details Begin
	'ipaddr_in = Request.ServerVariables ("REMOTE_ADDR") & "."
	'ipaddr_out = ""
	
	'while instr(ipaddr_in, ".") <> 0 
	'	ipaddr_out = ipaddr_out & Mid(ipaddr_in, 1, instr(ipaddr_in, ".") - 1)
	'	ipaddr_in = Mid(ipaddr_in, instr(ipaddr_in, ".") + 1, 100)
	'wend
	'v_name1 = "name_" & ipaddr_out
    'v_report_name=Request.Form("rr1")	
    'v_ipaddr=Request.ServerVariables ("REMOTE_ADDR")
    'v_user=Trim(Application(v_name1))
    'strSql="insert into web_log(web_log_id,user_name,ip_address,url_path) values(seq_web_log.nextval,'"&v_user&"','"&v_ipaddr&"','"&v_report_name&"')"
    'Response.write strSql
    'connObj.execute(strSql)
    'connObj.close
    'set connObj=nothing
	'Journal Details End

	Dim Qstr,z, firstpart, lastpart
	z=0
	Qstr="Page_Initialize("
	for i=1 to Request.QueryString.Count - 1
	if Request.QueryString.item(i) = "?" then
		if z=0 then
			Qstr=Qstr & "'%3f'" 
		else
			Qstr=Qstr & ",'%3f'" 
		end if
	else
		if z=0  then
			iput = Request.QueryString.item(i)
			iput = Replace(iput, "&", "%26")
			iput = Replace(iput, "'", "%27")
			iput = Replace(iput, "(", "%28")
			iput = Replace(iput, ")", "%28")
			'iput = Replace(iput, "-", "%2a")
			'iput = Replace(iput, "/", "%2F")
			QStr = Qstr & "'" & iput & "'"
		else
			iput = Request.QueryString.item(i)
			iput = Replace(iput, "&", "%26")
			iput = Replace(iput, "'", "%27")
			iput = Replace(iput, "(", "%28")
			iput = Replace(iput, ")", "%28")
			'iput = Replace(iput, "-", "%2a")
			'iput = Replace(iput, "/", "%2F")
			QStr = Qstr & ",'" & iput  & "'"
		end if
	
	end if
	z=z+1
	next
	Qstr=Qstr + ")"
	Response.write Qstr
	
%>
<!--<a href="<%=Qstr%>")>fasdf</a>
OnLoad="Page_Initialize('<%=Request.QueryString("rr1")%>', '<%=Request.QueryString("rp1")%>', '<%=Request.QueryString("rp2")%>', '<%=Request.QueryString("rp3")%>', '<%=Request.QueryString("rp4")%>', '<%=Request.QueryString("rp5")%>', '<%=Request.QueryString("rp6")%>', '<%=Request.QueryString("rp7")%>', '<%=Request.QueryString("rp8")%>', '<%=Request.QueryString("rp9")%>', '<%=Request.QueryString("rp10")%>')"
<input type=button name=c value=click onclick="<%=Qstr%>">-->
<body onload="<%=Qstr%>">
<SCRIPT LANGUAGE="VBScript">
	'Sub window_onLoad()
	'	Page_Initialize()
	'End Sub

	Sub Page_Initialize (r1, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18, p19, p20, s1, s2, s3, s4, s5, S6)
		On Error Resume Next
		Dim webBroker, repname

		reppath = "http://webserver/wls/reproot_new/"
		repname = r1
		repAPSuser = "?apsuser=Administrator&apspassword=&apsauthtype=secEnterprise"
		'repODBCuser = "&user0=appln1&password0=1appln"
		'repparams = "&promptex0=" & p1 & "&promptex1=" & p2 & "&promptex2=" & p3 & "&promptex3=" & p4
		'repparams = repparams & "&promptex4=" & p5 & "&promptex5=" & p6
		'repparams = repparams & "&promptex6=" & p7 & "&promptex7=" & p8
		'repparams = repparams & "&promptex8=" & p9 & "&promptex9=" & p10
		'repparams = repparams & "&promptex10=" & p11 & "&promptex11=" & p12
		'repparams = repparams & "&promptex12=" & p13
		'repsubs = "&user0@" & s1 & "=appln1&password0@" & s1 & "=1appln"
		'repsubs = repsubs & "&user0@" & s2 & "=appln1&password0@" & s2 & "=1appln"
		'repsubs = repsubs & "&user0@" & s3 & "=appln1&password0@" & s3 & "=1appln"		
		'repsubs = repsubs & "&user0@" & s4 & "=appln1&password0@" & s4 & "=1appln"		
		'repsubs = repsubs & "&user0@" & s5 & "=appln1&password0@" & s5 & "=1appln"		

		Set webBroker = CreateObject("WebReportBroker.WebReportBroker")

		if err.number <> 0 then
			window.alert "The Seagate Software Report Viewer for ActiveX is unable to create it's resource. Error Code: " & err.number
			CRViewer.ReportName = reppath & repname & repAPSuser & repODBCuser & repparams & repsubs
		else
			Dim webSource0

			Set webSource0 = CreateObject("WebReportSource.WebReportSource")

			webSource0.ReportSource = webBroker
			'webSource0.URL = reppath & repname & repAPSuser &  repODBCuser & repparams & repsubs
			'msgbox (reppath & repname & repAPSuser)
					
			
			webSource0.URL = reppath & repname & repAPSuser
			
			webSource0.AddParameter "apsuser", "Administrator"
			webSource0.AddParameter "apspassword", ""
			webSource0.AddParameter "apsauthtype", "secEnterprise"
			webSource0.AddParameter "user0", "appln1"
			webSource0.AddParameter "password0", "1appln"
			webSource0.AddParameter "promptex0", p1
			webSource0.AddParameter "promptex1", p2
			webSource0.AddParameter "promptex2", p3
			webSource0.AddParameter "promptex3", p4
			webSource0.AddParameter "promptex4", p5 
			webSource0.AddParameter "promptex5", p6
			webSource0.AddParameter "promptex6", p7
			webSource0.AddParameter "promptex7", p8
			webSource0.AddParameter "promptex8", p9
			webSource0.AddParameter "promptex9", p10
			webSource0.AddParameter "promptex10", p11
			webSource0.AddParameter "promptex11", p12
			webSource0.AddParameter "promptex12", p13
			webSource0.AddParameter "promptex13", p14
			webSource0.AddParameter "promptex14", p15
			webSource0.AddParameter "promptex15", p16
			webSource0.AddParameter "promptex16", p17
			webSource0.AddParameter "promptex17", p18
			webSource0.AddParameter "promptex18", p19
			webSource0.AddParameter "promptex19", p20
			webSource0.AddParameter "user0@" & s1, "appln1"
			webSource0.AddParameter "password0@" & s1, "1appln"
			webSource0.AddParameter "user0@" & s2, "appln1"
			webSource0.AddParameter "password0@" & s2, "1appln"
			webSource0.AddParameter "user0@" & s3, "appln1"
			webSource0.AddParameter "password0@" & s3, "1appln"		
			webSource0.AddParameter "user0@" & s4, "appln1"
			webSource0.AddParameter "password0@" & s4, "1appln"		
			webSource0.AddParameter "user0@" & s5, "appln1"
			webSource0.AddParameter "password0@" & s5, "1appln"
			webSource0.AddParameter "user0@" & s6, "appln1"
			webSource0.AddParameter "password0@" & s6, "1appln"

			webSource0.PromptOnRefresh = False
			CRViewer.ReportSource = webSource0
		end if

		'alert repparams
		
		CRViewer.ViewReport
		
		'prompt "hai testing", reppath & repname & repAPSuser & repODBCuser & repparams & repsubs
	End Sub
</SCRIPT>
</body>