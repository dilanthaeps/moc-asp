<!--#include file="common_dbconn.asp"-->
<%	'===========================================================================
	'	Template Name	:	Menu Group Assignment Save
	'	Template Path	:	.../menu_grp_asgn_save.asp
	'	Functionality	:	To save the Menu and User Group Assignment. 
	'	stored_proc		:	N/A
	'	Called By		:	../menu_grp_asgn_entry.asp  
	'	Created By		:	Sethu Subramanian Rengarajan, Tecsol Pte Ltd, Singapore
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<%
'Response.Buffer = false
v_tem = Now()&session.SessionID 
dim v_array
for  each varitem in Request.Form
if left(varitem,2)="v_" then
	v_array = split(varitem,"_",-1,1)
	ctr = 0
		while ctr <= ubound(v_array)
		 'Response.Write "Within while loop <br>"&v_array(ctr)&" -- "&cstr(ctr)&"<br>"
		 strSql = "SELECT 1   FROM moc_tc_moc_asgn  where time_charterer_id="&v_array(2)&" and moc_id="&v_array(1)&" and rownum=1"
		 'Response.Write "<br>"&strSql&"<br>"
		 Set rsObj = connObj.Execute(strSql) 
		 v_record_found = "No"
			while not rsObj.EOF
				v_record_found = "Yes"
				strSql = "Update moc_tc_moc_asgn SET CONTROL='"&v_tem&"' where time_charterer_id="&v_array(2)&" and moc_id="&v_array(1)
				Set rsObj1 = connObj.Execute(strSql)
				rsObj.MoveNext
			wend
			if v_record_found="No" then
				strSql = "Insert into  moc_tc_moc_asgn(tc_moc_asgn_id,moc_id,time_charterer_id,control) values (0,"&v_array(1)&","&v_array(2)&",'"&v_tem&"')"
				'Response.Write strSql
				Set rsObj2 = connObj.Execute(strSql)
			end if
		 ctr=ctr+1
		wend
	v_grp_id = varitem
	'Response.Write varitem &"= "&request(varitem) &"-"&v_tem&"<br>"
	
	'strSql = "SELECT    UGA.USER_GRP_ASGN_ID , UGA.USER_ID , UGA.GRP_ID , UGA.CREATE_DATE FROM USER_GRP_ASGN UGA where UGA.USER_ID= and UGA.GRP_ID="
	'Set rsObj = connObj.Execute(strSql) 
end if
next
strSql = "delete  moc_tc_moc_asgn where control <> '"&v_tem&"'"
Set rsObj3 = connObj.Execute(strSql)
%>
<!--include file="common_footer.asp"-->
<%
Response.Redirect "moc_tc_asgn.asp?v_message=Changes+effected"
%>
