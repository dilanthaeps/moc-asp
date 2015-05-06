  <!--#include file="common_dbconn.asp"-->
  <%
  'Response.Buffer=false
  Function RT(word_with_tilde)
	If word_with_tilde = Null Or IsNull(word_with_tilde) Or word_with_tilde = "" Or word_with_tilde = "NULL" Then
		RT = ""
	Else
		RT=replace(replace(word_with_tilde,"~","<tilde>"),chr(13),"<enter>")
	End If
  End function
 
   rndno= Replace(CStr(CDbl(Now)), ".", "-")
  'Create by Sethu Subramanian Rengarajan, Tecsol Pte Ltd.

  questionnaire_id = request("questionnaire_id")
  moc_id= request("moc_id")
  vessel_code=moc_id ' Just for the sake of clear reference.

'testing by senthil.
  questionnaire_id = "90000017"
  moc_id= "88888888"
  vessel_code=moc_id ' Just for the sake of clear reference.
'testing by senthil.
 
  '-------------- Word Merge text file creation --------------
if questionnaire_id = "90000006" then
   strSql="SELECT   user_id fax_to, 'Mannaer & co' fax_company, '998989' fax_phone from wls_user_master where user_id='senthil' "
   v_question="fax_to~fax_company~fax_phone"
end if

if questionnaire_id = "90000017" then
   strSql="SELECT   user_id person_name, 'new com & co' person_job, '998989' person_dob from wls_user_master where user_id='sethu' "
      v_question="person_name~person_job~person_dob"
end if

   Set rsObj=connObj.Execute(strSql)
   'v_bugs = "" ' Answers which has ~ (delimeter) as content
%>
<%
         set fs=CreateObject("Scripting.FileSystemObject")
		 filepath= "E:\mail_merge\data"  'Drive of the local machine which running the deamon - Where VID path of the webserve2 is mapped in the local machine
		 'filepath=filepath&"\q"&questionnaire_id&"_v"&vessel_code&".txt"
		 filepath=filepath&"\q"&cstr(questionnaire_id) & "-v" & vessel_code &"-" & rndno & ".doc" ' currently moc_id is given
		 output_file_name= "E:\mail_merge\rpt_docs" & "\q"&cstr(questionnaire_id) & "-v" & vessel_code &"-r-" & rndno & ".doc"
		 'Response.Write "<p><b><center> The file Path "&filepath& "</center></b><p>"
		 v_rpt_file_name="D:\vpd\mail_merge\data"&"\q"&cstr(questionnaire_id) & "-v" & vessel_code &"-" & rndno & ".doc"
		 'Response.Write v_rpt_file_name 
		 set file=fs.CreateTextFile(v_rpt_file_name,true,false)

         file.writeLine(v_question)
    
         v_answer=""
         
 if questionnaire_id = "90000006" then
         while not rsObj.EOF 
			'Response.Write rsObj("ANSWER_DATA")&"~"
			v_answer=v_answer & RT(rsObj("fax_to")) & "~" & RT(rsObj("fax_company")) & "~" & RT(rsObj("fax_phone"))
			rsObj.MoveNext
         wend
 end if
 
 If questionnaire_id = "90000017" then
         while not rsObj.EOF 
			'Response.Write rsObj("ANSWER_DATA")&"~"
			v_answer = v_answer & RT(rsObj("person_name")) & "~" & RT(rsObj("person_job")) & "~" & RT(rsObj("person_dob"))
			rsObj.MoveNext
         wend
 end if

 file.WriteLine(v_answer)
 file.close
 set file = nothing

  '-------------- Word Merge text file creation end ---------
 
strSql = "INSERT INTO  VPD_TERM_REPORTS ( TERM_REPORT_ID,QUESTIONNAIRE_ID , VESSEL_CODE , REPORT_NAME ,REPORT_TYPE, DATA_FILE_NAME , OUTPUT_FILE_NAME      , CREATE_DATE ) VALUES (seq_vpd_term_reports.nextval,"& questionnaire_id &",'" & vessel_code & "','" & "MOC" & "','" & "WORD" & "','" & filepath &"','" &output_file_name & "',sysdate)"
'Response.Write "<br>" & strSql
Set rsObj = connObj.Execute(strSql)

  connObj.Close
  set connObj=nothing
  %>
  <%
  v_url="show_report.asp?report_name=" & output_file_name

  Response.Write "Data File Name - " & v_rpt_file_name & "<BR>"
  Response.Write "Output File Name - " & output_file_name & "<P>"
  Response.Write v_url & "<BR>"
  'senthil commneted for testing, Response.Redirect(v_url)
  %>