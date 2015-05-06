	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	FBM  Delete Screen
	'	Template Path	:	.../fbm_delete.asp
	'	Functionality	:	To Fleet Broadcase Message Delete information
	'	Called By		:	../fbm_maint.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	'Response.Buffer = false
	v_message=""
	dim i
	For each i in Request("v_deleteditems")
		strSql="DELETE FROM wls_fbm where fbm_id="
		strSql=strSql & "'"&i&"'" 
		'Response.Write strSql
		connObj.Execute(strSql)
		v_message = v_message&" Fleet broadcase message <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	'http://webserve2/wls/new/moc/ins_request_remark_maint.asp?v_ins_request_id=10000087&vessel_name=AMOY&moc_name=CEPSA
	v_string="fbm_entry.asp?v_fbm_id=0&v_message="
	v_string = v_string & message& "&v_close=true"
	Response.Redirect v_string
%>   


