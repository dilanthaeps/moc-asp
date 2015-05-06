<!--#include file="common_dbconn.asp"-->
<%		    
	Response.Write "Testing<BR>"
	Response.Write Request("cons_pass_number") & "<BR>"
	
	if request("cons_id")<>"" then
		cons_id=request("cons_id")
	else
		cons_id=null
	end if
	if request("cons_fin_number")<>"" then
		cons_fin_number=request("cons_fin_number")
	else
		cons_fin_number=null
	end if
	if request("cons_pass_number")<>"" then
		cons_pass_number=request("cons_pass_number")
	else
		cons_pass_number=null
	end if
	if request("cons_acc_number")<>"" then
		cons_acc_number=request("cons_acc_number")
	else
		cons_acc_number=null
	end if

	if request("cons_local_address")<>"" then
		cons_local_address=request("cons_local_address")
	else
		cons_local_address=null
	end if
	if request("cons_overseas_address")<>"" then
		cons_overseas_address=request("cons_overseas_address")
	else
		cons_overseas_address=null
	end if

	if request("cons_status")<>"" then
		cons_status=request("cons_status")
	else
		cons_status=null
	end if

	if request("cons_date_of_join")<>"" then
		cons_date_of_join="format('" & Trim(Request("cons_date_of_join")) & "','mm/dd/yyyy')"
	else
		cons_date_of_join="null"
	end if
	if request("date_of_birth")<>"" then
		date_of_birth="format('" & Trim(Request("date_of_birth")) & "','mm/dd/yyyy')"
	else
		date_of_birth="null"
	end if
	if request("ep_expiry_date")<>"" then
		ep_expiry_date="format('" & Trim(Request("ep_expiry_date")) & "','mm/dd/yyyy')"
	else
		ep_expiry_date="null"
	end if
	if request("skill_set")<>"" then
		skill_set=request("skill_set")
	else
		skill_set=null
	end if

	Dim Idval
	Idval = Request("v_history_id")
	if Idval = "" then
		strSqlcount="select max(history_id)+1 as r_count from tec_cons_details"
		set rsObjcount=connObj.execute(strSqlcount)
		if not(rsObjcount.eof or rsObjcount.bof) then
			rsObjcount.movefirst
			set rec_count=rsObjcount("r_count")
		end if
		strSql = "INSERT INTO tec_cons_details(history_id, cons_id, cons_date_of_join, date_of_birth, cons_fin_number,cons_pass_number,cons_acc_number,cons_local_address,cons_overseas_address,ep_expiry_date,skill_set,cons_status) Values("&rec_count&", '"&cons_id&"',"&cons_date_of_join&","&date_of_birth&",'"&cons_fin_number&"','"&cons_pass_number&"','"&cons_acc_number&"','"&cons_local_address&"','"&cons_overseas_address&"',"&ep_expiry_date&",'"&skill_set&"','"&cons_status&"')"
		v_message = "Consultant Details Created Successfully"
	else
		strSql = "Update tec_cons_details set cons_id='"&id_name&"',cons_date_of_join="&cons_date_of_join&",date_of_birth="&date_of_birth&",cons_fin_number='"&cons_fin_number&"',cons_pass_number='"&cons_pass_number&"',cons_acc_number='"&cons_acc_number&"',cons_local_address='"&cons_local_address&"',cons_overseas_address='"&cons_overseas_address&"',ep_expiry_date="&ep_expiry_date&",skill_set='"&skill_set&"',cons_status='"&cons_status&"' where history_id=" & Idval & ""
		v_message = "Consultant Details Updated Successfully"
	end if	
	   
	'connObj.Execute(strSql)
	connObj.Close
	set connObj=nothing

	Response.Redirect "cons_maint.asp?v_message="&v_message
%>   
