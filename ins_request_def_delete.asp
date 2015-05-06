	<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Inspection Deficiency  Delete Screen
	'	Template Path	:	.../ins_request_def_delete.asp
	'	Functionality	:	To MOC Inspection Deficiency Delete information
	'	Called By		:	../ins_request_def_entry.asp
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	'Response.Buffer = false
	v_message=""
	dim i
	For each i in Request("v_deleteditems")
		strSql="DELETE FROM moc_deficiencies where deficiency_id="
		strSql=strSql & "'"&i&"'" 
		'Response.Write strSql
		connObj.Execute(strSql)
		v_message = v_message&" MOC Deficiency Detail <i>: " &i&"</i> deleted Successfully !<br>"
	next      
	message = Server.URLEncode(v_message)
	connObj.close
	set connObj=nothing
	'http://webserve2/wls/new/moc/ins_request_remark_maint.asp?v_ins_request_id=10000087&vessel_name=AMOY&moc_name=CEPSA
	v_string="ins_request_def_maint.asp?v_message="
	v_string = v_string & message& "&v_ins_request_id="&request("v_ins_request_id")&"&vessel_name="&request("vessel_name")&"&moc_name="&request("moc_name")
	Response.Redirect v_string
%>   


