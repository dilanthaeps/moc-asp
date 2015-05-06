<!--#include file="moc.css"-->
<%
report_name = request("report_name")
'Response.Write request("report_name")
report_name = replace(report_name,"~","\")
v_url_of_report_name = mid(report_name,4,200)
v_no_of_times = 0
v_url = ""
v_merge_errors = ""
'Response.Write v_url_of_report_name

'E:\mail_merge\rpt_docs\q10000118-v902-37290-7080092593.doc

%>
  <!--#include file="common_dbconn.asp"-->
  
<%
  strSql="SELECT  OUTPUT_FILE_NAME FROM VPD_TERM_REPORTS WHERE OUTPUT_FILE_NAME='" & report_name & "' and OUT_FILE_CREATION_END_TIME is not null"
  'Response.Write strSql
  set rsObj= connObj.execute(strSql)
  while not rsObj.eof
  v_url= "<a href=http://webserve2/vid/" & v_url_of_report_name & " >Click Here to View the report</a>"
  rsObj.Movenext
  wend
  strSql="update VPD_TERM_REPORTS  set no_of_times = no_of_times+1 WHERE OUTPUT_FILE_NAME='" & report_name & "'"
  set rsObj= connObj.execute(strSql)
  strSql="SELECT  no_of_times,merge_errors FROM VPD_TERM_REPORTS WHERE OUTPUT_FILE_NAME='" & report_name & "'"
  'Response.Write strSql
  set rsObj= connObj.execute(strSql)
  while not rsObj.eof
  v_no_of_times= rsobj("no_of_times")
  v_merge_errors = rsobj("merge_errors")
  rsObj.Movenext
  wend
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Terminal Reports Result Page</title>
</HEAD>
<BODY class=bcolor>
<center>
<h1>Generating Questionnaire Template</h1>
<p><p><p><p><p><p><p><p><p><p><p>
<table width=100%>
<%
if v_merge_errors <> "" then
Response.write "<br><font color=red> <h2>Error, please contact Administrator</h2> </font><br> The following error was reported in mail merge: <br>1. " &  v_merge_errors  & " <br>"
%>
<b>Other Possible Reasons :</b><p>

01.  The word file template may be corrupted<br>
02.  The network connection of the Deamon machine may be off<br>
03.  MS word application  may be open at the Deamon machine<br>
04.  Some unkilled word_save process may running and interupting the current process. You may switch off and restart the mail merge deamon computer.<br>

<%
elseif cint(v_no_of_times) > 7 then
Response.write "<br><font color=red> <h2>Error, please contact Administrator</h2> </font><br> The possible reasons could be : <br>1.The mail merge deamon may not be running. <br> 2.The network connection may be faulty at the mail merge deamon machine. <br> "
%>
<b>Other Possible Reasons :</b><p>

01.  The word file template may be corrupted<br>
02.  The network connection of the Deamon machine may be off<br>
03.  MS word application  may be open at the Deamon machine<br>
04.  Some unkilled word_save process may running and interupting the current process. You may switch off and restart the mail merge deamon computer.<br>

<%
elseif v_url = "" then
%>
<tr><td>
 

<!--<strong><font color=red> Processing the Report!! </font></strong>-->
<p>
<b><center>Please Wait !! </b>
</center>
<p>
 
</td></tr>
<%
v_current_url="show_report.asp?report_name="&Request.QueryString("report_name")
v_current_url=replace(v_current_url,"\","~")
'Response.Write v_current_url &"<br>"
%>

<script language="Javascript">
<!--
var URL   = "<%=v_current_url %>"
var speed = 7000
function reload() {
location = URL
}
setTimeout("reload()", speed);
//-->
</script>

<%
end if
%>
<tr><td>
<h1>
<%
Response.Write v_url & "&nbsp;</h1> </td></tr>"
%>
</table>
<hr>
<%
Response.Write "<br><center> Time :" & now & "</center>"
%>
<hr>
 <!--#include file="common_footer.asp"-->  
</BODY>
</HTML>
