<%option explicit%>
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
dim SQL,arr,i

if Request.Form("cmbVessel")="" then
	Response.Write "Parameters missing"
	Response.End
end if

'delete existing time charterer assignment for vessel
SQL = "Delete from MOC_TC_VESSEL_ASGN where vessel_code='" & Request.Form("cmbVessel") & "'"
connObj.execute SQL

'insert record for time charterer assignment to vessel
if Request.Form("cmbTC")<>"" then
	SQL = "Insert into MOC_TC_VESSEL_ASGN(vessel_code,time_charterer_id,created_by,remarks)"
	SQL = SQL & " values("
	SQL = SQL & " '" & Request.Form("cmbVessel") & "',"
	SQL = SQL & " " & Request.Form("cmbTC") & ","
	SQL = SQL & " '" & USER & "',"
	SQL = SQL & " '" & left(Replace(Request.Form("txtRemarks"),"'","''"),500) & "'"
	SQL = SQL & " )"
	connObj.execute SQL
end if

'delete existing moc records for time charterer and vessel
SQL = "Delete from MOC_TC_MOC_ASGN where vessel_code='" & Request.Form("cmbVessel") & "'"' and time_charterer_id=" & Request.Form("cmbTC")
connObj.execute SQL

'insert records for mandatory moc to time charterer and vessel
if Request.Form("txtSelMOCMand")<>"" and Request.Form("cmbTC")<>"" then
	arr = split(Request.Form("txtSelMOCMand"),",")
	for i=0 to ubound(arr)
		if arr(i)<>"" then
			SQL = "Insert into MOC_TC_MOC_ASGN(vessel_code,time_charterer_id,moc_id,created_by,mandatory)"
			SQL = SQL & " values("
			SQL = SQL & " '" & Request.Form("cmbVessel") & "',"
			SQL = SQL & " " & Request.Form("cmbTC") & ","
			SQL = SQL & " " & trim(arr(i)) & ","
			SQL = SQL & " '" & USER & "',"
			SQL = SQL & " 1"
			SQL = SQL & " )"
			connObj.execute SQL
		end if
	next
end if

'insert records for optional moc to time charterer and vessel
if Request.Form("txtSelMOCOpt")<>"" and Request.Form("cmbTC")<>"" then
	arr = split(Request.Form("txtSelMOCOpt"),",")
	for i=0 to ubound(arr)
		if arr(i)<>"" then
			SQL = "Insert into MOC_TC_MOC_ASGN(vessel_code,time_charterer_id,moc_id,created_by,mandatory)"
			SQL = SQL & " values("
			SQL = SQL & " '" & Request.Form("cmbVessel") & "',"
			SQL = SQL & " " & Request.Form("cmbTC") & ","
			SQL = SQL & " " & trim(arr(i)) & ","
			SQL = SQL & " '" & USER & "',"
			SQL = SQL & " 0"
			SQL = SQL & " )"
			connObj.execute SQL
		end if
	next
end if
%>
<!--include file="common_footer.asp"-->
<%
Response.Redirect "ves_tc_asgn.asp?VID=" & Request.Form("cmbVessel") & "&v_message=Changes+effected"
%>
