<%@language=vbscript%>
<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Inspection Request Deficiency Entry
	'	Template Path	:	ins_request_def_entry.asp
	'	Functionality	:	To allow the entry/edit of the MOC Inspection Request Deficiency details
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	12th September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	dim v_ins_request_id,v_button_disabled,v_mode,v_header
	dim rsObj_Deficiency_Status,rsCode,SQL
	dim v_selected,v_sort_order,rsObjNextSortOrder
	dim sGeneral,sLow,sHigh
	dim VID,VNAME,INSP_TYPE
	
	dim viq_count_check
	dim viq_dontcount_check
	
	VID = Request.QueryString("VID")
	VNAME = request("VNAME")
	
	dim DISPLAY_LINK_TO_WORKLIST
	'to hide link to Worklist, make DISPLAY_LINK_TO_WORKLIST = false
	DISPLAY_LINK_TO_WORKLIST = true
	'DISPLAY_LINK_TO_WORKLIST = false
	v_ins_request_id=request("v_ins_request_id")
	Response.Buffer = false

	v_button_disabled = "DISABLED"
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCSUP" Then
	if UserIsAdmin or UserIsSuper then
		v_button_disabled = ""
	End If

	strSql = "Select insp_type from moc_inspection_requests where request_id=" & v_ins_request_id
	set rsObj = connObj.Execute(strSql)
	INSP_TYPE = rsObj("insp_type")
	rsObj.close

	Dim Idval
	Idval = Request.QueryString("v_def_id")
	if Idval <> "0" then
		v_mode="edit"
		v_header="Observations"
				
		strSql = "Select"
		strSql = strSql & " deficiency_id, request_id, deficiency,reply,wls_job_list_id,"
		strSql = strSql & " section, status, sort_order, risk_factor,"
		strSql = strSql & " action_code,action_description,"
		strSql = strSql & " to_char(create_date, 'DD-Mon-YYYY') create_date1, created_by,"
		strSql = strSql & " to_char(last_modified_date,'DD-Mon-YYYY') last_modified_date1, last_modified_by, nvl(VIQ_DONT_COUNT,0) VIQ_DONT_COUNT"
		strSql = strSql & " from moc_deficiencies "
		strSql = strSql & " WHERE deficiency_id =" & Idval
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
		
		select case cstr(rsObj("risk_factor").value)
			case "GENERAL":sGeneral=" checked"
			case "LOW":sLow=" checked"
			case "HIGH":sHigh=" checked"
			case else sGeneral=" checked"
		end select
		
		if cint(rsObj("VIQ_DONT_COUNT")) = 1 then
			viq_count_check = ""
			viq_dontcount_check = " checked "
		else
			viq_count_check = " checked"
			viq_dontcount_check = ""
		end if
		
		
	else
		v_mode="Add"
		v_header="Create Observation "
	end if
	
	' Select List -      Remark_Status   
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Deficiency_status' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Deficiency_Status = connObj.Execute(strSql)
	
	SQL = "Select * from MOC_DEFICIENCY_ACTION_CODES where code_type='Deficiency Action Code' order by code"
	set rsCode = connObj.Execute(SQL)
	
	if sGeneral="" and sLow="" and sHigh="" then sGeneral= " checked"
%>	
<html>
<head>
<TITLE> MOC Inspection Observation Entry/Edit Screen</TITLE>
<META name=VI60_defaultClientScript content=VBScript>
<META HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 1996 14:25:27 GMT"> 
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<style>
h4
{
	margin:2px;
}
.link1
{
	font-size:9px;
}
.link2
{
	font-size:9px;
	color:gray;
}
A:hover
{
	color:red;
}
</style>
<script language="VBScript" runat=server>
function SFIELD(fname)
	if v_mode="edit" then
		SFIELD = rsObj(cstr(fname))
	else
		SFIELD = ""
	end if
End function
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	form1.action_code.value = "<%=SFIELD("action_code")%>"
	SetActionDesc
End Sub

sub risk_factorClick
	set obj = window.event.srcElement

	if obj.checked then
		form1.risk_factor(0).checked = false
		form1.risk_factor(1).checked = false
		form1.risk_factor(2).checked = false
		obj.checked=true
	end if
	if not form1.risk_factor(0).checked and not form1.risk_factor(1).checked and not form1.risk_factor(2).checked then
		form1.risk_factor(0).checked = true
	end if
end sub
sub section_onblur
	GetSectionText(window.event.srcElement.value)
end sub
sub GetSectionText(s)
	
end sub
function GetJoblistID(obj)
	dim selText,sRisk
	dim sURL,win,NewJobID,defid,joblistid
	defid = obj.GetAttribute("defid")
	joblistid = obj.GetAttribute("joblistid")
	
	sURL = "NCR_list.asp?VID=<%=VID%>&VNAME=<%=VNAME%>&DEFID=" & defid & "&JOBLISTID=" & joblistid
	for i=0 to 2
		if form1.risk_factor(i).checked then sRisk = form1.risk_factor(i).value
	next
	selText = "Risk: " & sRisk & vbCrLf & form1.deficiency.value
	NewJobID = window.showModalDialog(sURL,selText,"dialogHeight:500px;dialogWidth:800px;resizable:yes;center:yes;status:no;scroll:yes;")
	if NewJobID<>joblistid then
		form1.action=""
		form1.submit
	end if
end function
sub SetActionDesc
	lblActionDesc.innerText=form1.action_code.options(form1.action_code.selectedIndex).getAttribute("desc")
end sub
</script>
<script language="Javascript" src="js_date.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<SCRIPT LANGUAGE="JAVASCRIPT">
function sClose()
{
	var name = true
	if (name== true)
	{
		v_val = "ins_request_def_save.asp?";
		document.form1.action=v_val;
		//document.form1.submit();
	}
	else
	{
		return false;
	}
}
function cClose()
{
	//var name= confirm("Are you sure? The changes will be saved and !!")
	var name = true
	if (name== true)
	{
		//v_val = "ins_request_def_maint.asp";
		//self.opener.document.form1.action=v_val;
		//self.opener.document.form1.submit();
		self.opener.history.go(0);
		self.close();
	}
	else
	{
		return false;
	}
}
</script>
<%
if request("v_close") = "true" then
%>
<SCRIPT LANGUAGE=javascript>
<!--
cClose();
//-->
</SCRIPT>

<%
end if
%>
<script language="Javascript">
function fndelete(v_delete_item)
{
var name=confirm("Are you sure? This Observation record will be deleted!")
if (name== true)
		{
			v_val = "ins_request_def_delete.asp?v_deleteditems="+v_delete_item;
			//v_val = "list_all_server_variables.asp?v_deleteditems="+v_delete_item;
			document.form1.action=v_val;
			//alert(document.form1.action);
			document.form1.submit();
			//self.close() ;
		}
		else
		{
		return false;
		}
}
function validate_fields()
{
	if (document.form1.section.value == "")
	{
		alert ("Enter Section");
		document.form1.section.focus();
		return false;
	}
	if (document.form1.action_description.value.length>500)
	{
		alert ("Action description too long. Please restrict to 500 chars");
		document.form1.action_description.select();
		document.form1.action_description.focus();
		return false;
	}
	if (document.form1.deficiency.value == "")
	{
		alert ("Enter Deficiency");
		document.form1.deficiency.focus();
		return false;
	}
	if(document.form1.deficiency.value.length>4000)
	{
		alert("Deficiency Maximum 4000 Chars only");
		document.form1.deficiency.focus();
		return false;
	}
	if(document.form1.reply.value.length>4000)
	{
		alert("Reply Maximum 4000 Chars only");
		document.form1.reply.focus();
		return false;
	}
	if(document.form1.status.value == "")
	{
		alert ("Enter  Status");
		document.form1.status.focus();
		return false;
	}

	if(document.form1.sort_order.value == "")
	{
		alert ("Enter  Sort Order");
		document.form1.sort_order.focus();
		return false;
	}
	if(isNaN(document.form1.sort_order.value))
	{
		alert ("Enter  Sort Order - Number only");
		document.form1.sort_order.focus();
		return false;
	}
} 
function openJoblistPopup(joblist_id)
{
	//alert(joblist_id)
	winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	winStats += 'scrollbars=yes,status=yes'
	if (navigator.appName.indexOf("Microsoft") >= 0) 
	{
		winStats += ',left=20,top=10,width=' + (screen.width - 50) + ',height=' + (screen.height - 90)
	}
	else
	{
		winStats += ',screenX=350,screenY=200,width=300,height=180'
	}
	adWindow = window.open("http://webserve2/wls/joblist_popup_frame.asp?v_joblist_id=" + joblist_id + "&v_opener_name=moc", "JobDetails", winStats);
	adWindow.focus();
	return false;
}

</SCRIPT>

<script id="GetAsyncResponse" src="GetAsyncResponse.js" type="text/javascript"></script>
<script type="text/javascript">
        function CheckHighRisk(txtId){
            var sec = document.getElementById(txtId).value;
            try{
                
                if (sec == "8.75" || sec == "9.27" || sec == "11.3")
                {
                    //document.getElementById("viq").style.display="";                    
                    var res = confirm("This is a VIQ section.\n\n Click OK if you want it to count");
                    if (res == true){
                        document.getElementById("viq_count").checked = true;
                    }
                    else{
                        document.getElementById("viq_dontcount").checked = true;
                    }
                    return;
                }
                else
                {
                    //document.getElementById("viq").style.display="none";
                }
                document.getElementById("checking").innerHTML = "<<<";
                var d = new Date();
                var ms = d.valueOf();    
                
                var qs = "noCache=" + ms
                var url;
                
                qs += "&func=CHECK_HIGHRISK"
                qs += "&sec=" + sec
                
                url = "approval.asp?" + qs; 
                
                var c = new  AsyncResponse(url,asyncCallBack,0)
                c.getResponse();        
           }
           catch(ex)
           {
           //nothing
           }         
        } 
        function asyncCallBack(retval)
        {
            if (retval == "YES"){
                document.getElementById("chkHigh").checked = true;}
                document.getElementById("checking").innerHTML = "";
        }
</script>


</HEAD>
<BODY class=bcolor>

<h3><%= v_header %></h3>
<h4>Vessel : <%=VNAME%>  &nbsp; &nbsp; &nbsp; &nbsp;MOC &nbsp; &nbsp;: <%=request("moc_name")%></h4>
<form name=form1 action=ins_request_def_save.asp method=post onsubmit="javascript:return validate_fields();" >

<INPUT type="hidden" id="v_ins_request_id" name="v_ins_request_id" value="<%=request("v_ins_request_id")%>">
<INPUT type="hidden" id="deficiency_id" name="deficiency_id" value="<%=SFIELD("deficiency_id")%>">
<INPUT type="hidden" id="request_id" name="request_id" value="<%=request("v_ins_request_id")%>">
<INPUT type="hidden" id="vessel_name" name="vessel_name" value="<%=VNAME%>">
<INPUT type="hidden" id="moc_name" name="moc_name" value="<%=request("moc_name")%>">
<table width="100%" border=1>
<%
if v_mode = "edit" then
%>
<div align="right">
<%if DISPLAY_LINK_TO_WORKLIST then%>
<%if rsObj("wls_job_list_id")<>"" then%>
<a class=link1 href="#" onclick="openJoblistPopup(<%=SFIELD("wls_job_list_id")%>)" onmouseover="javascript:window.event.cancelBubble=true" onmousemove="javascript:window.event.cancelBubble=true">View job in Worklist</a>
&nbsp;&nbsp;
<%end if%>
<a class=link2 href="#" defid="<%=SFIELD("deficiency_id")%>" joblistid="<%=SFIELD("wls_job_list_id")%>" onclick="javascript:GetJoblistID(this);return false;" onmouseover="javascript:window.event.cancelBubble=true" onmousemove="javascript:window.event.cancelBubble=true">Create link to Worklist</a>
<%end if%>

<INPUT type="button" id="delete" name="delete" value="Delete" <% =v_button_disabled %>
	onclick="Javascript:return fndelete(<%=SFIELD("deficiency_id")%>);">
</div>
<%
end if
%>
<TR>
	<TD class="tableheader">Section</TD>
	<TD class="tabledata" width=200><div nowrap>
		<!--if v_mode <> "edit" then -->
		<INPUT type="text" id="section" name="section" value="<%=SFIELD("section")%>" onblur="CheckHighRisk('section')">
		<font color=red>*</font></div>
		<div id=viq style="display:block;">		
		VIQ: <input type=radio id=viq_count name=viq_dont_count value="0" <%=viq_count_check%> />Count&nbsp;&nbsp;<input type=radio id=viq_dontcount name=viq_dont_count value="1" <%=viq_dontcount_check%>/>Don't Count<br>
		(for VIQ Sections: 8.75, 9.27, 11.3 only) </div>
	<td class="tableheader">Risk&nbsp;&nbsp;&nbsp;&nbsp;
	<td nowrap class="tabledata" width=200 style="font-size:10px;font-weight:bold">&nbsp;&nbsp;
		General<input type=radio id="chkGen" name=risk_factor value="GENERAL" <%=sGeneral%> onclick="risk_factorClick">&nbsp;&nbsp;
		Low<input type=radio id="chkLow" name=risk_factor value="LOW" <%=sLow%> onclick="risk_factorClick">&nbsp;&nbsp;
		High<input type=radio id="chkHigh" name=risk_factor value="HIGH" <%=sHigh%> onclick="risk_factorClick">
		<span id="checking"></span>
	</TD>
</TR>
<tr <%if INSP_TYPE<>"PSC" then Response.Write " style='display:none'"%>>
	<td class="tableheader">Action code/<br>Remarks</td>
	<TD class="tabledata" colspan=3>
	<table width=100% cellspacing=0 cellpadding=0>
	  <tr>
	    <TD class="tabledata" style="vertical-align:top">
	    <select name=action_code onchange="SetActionDesc">
	      <option desc="" value="">Select code</option>
	      <%while not rsCode.eof%>
	      <option desc="<%=rsCode("description")%>" value="<%=rsCode("code")%>"><%=rsCode("code")%>
	      <%rscode.movenext
	      wend%>
	    </select>
	    <td width=100% style="text-align:left;"><div id=lblActionDesc style="background-color:lightgrey;font-style:italic;width:100%"></div>
	  <tr>
	    <td colspan=2><textarea name=action_description style="width:100%"><%=SFIELD("action_description")%></textarea>
	</table>
<TR>
	<TD class="tableheader">Observation</TD>
	<TD class="tabledata" colspan=3>
		<TEXTAREA rows=8 style="width:97%" id="deficiency" name="deficiency"><%=SFIELD("deficiency")%></TEXTAREA>
		<font color=red>*</font>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		Reply
	</TD>
	<TD class="tabledata" colspan=3>
		<TEXTAREA rows=8 style="width:97%" id="reply" name="reply" class="textlightbluecolor"><%=IIF(SFIELD("reply") = "", "Noted.", SFIELD("reply")) %></TEXTAREA>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		STATUS 
	</TD>
	<TD class="tabledata" colspan=3>
	 	<select name="status" id="status">
				<%
					if not(rsObj_Deficiency_Status.eof or rsObj_Deficiency_Status.bof) then
					
						while not rsObj_Deficiency_Status.eof

							v_selected = ""
							If SFIELD("STATUS") <> "" Then
								If CStr(rsObj_Deficiency_Status("sys_para_id")) = CStr(SFIELD("STATUS")) Then
									v_selected = "SELECTED"
								End If
							Else
								If UCase(rsObj_Deficiency_Status("sys_para_id")) = "ACTIVE" Then
									v_selected = "SELECTED"
								End If
							End If	
				%>
						<OPTION VALUE="<%=rsObj_Deficiency_Status("sys_para_id")%>" <% =v_selected %>><%=rsObj_Deficiency_Status("para_desc")%></OPTION>
				<%		rsObj_Deficiency_Status.movenext
						wend
					end if
				%>
		</select> 
		<font color=red>*</font>
		
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		Sort Order
	</TD>
	<TD class="tabledata" colspan=3>
<%
	v_sort_order = SFIELD("sort_order")
	If v_sort_order = "" Then

		strSql = "select nvl(max(sort_order), 0) + 1 next_sort_order "
		strSql = strSql & "from moc_deficiencies "
		strSql = strSql & "where request_id = " & v_ins_request_id
		
		'Response.Write strSql & "<BR>"
		Set rsObjNextSortOrder = connObj.Execute(strSql)
		
		If rsObjNextSortOrder.EOF = False Then
			v_sort_order = rsObjNextSortOrder("next_sort_order")
		End If
		
		rsObjNextSortOrder.Close
		Set rsObjNextSortOrder = Nothing

	End If
%>

		<INPUT type="text" id="sort_order" name="sort_order"  value="<% =v_sort_order %>" maxlength=4>
		<font color=red>*</font>
	</TD>
</TR>
<TR>
	<td colspan="4" class="tabledata" align=center> 
         <INPUT type="submit" value="Save" id=submit1 name="submit1" <% =v_button_disabled %>> &nbsp;
         <INPUT type="submit" value="Save and Close" id=submit4 name="submit4" <% =v_button_disabled %> OnClick="javascript:return sClose();">&nbsp;
         <INPUT type="button" value="Close without Save" id=back name="back" onclick="javascript:return cClose();"> &nbsp;
      </td>
</TR>
</table>
</form>
<font color=red>*</font> Denotes Mandatory Field.<br>
<span style="color:maroon;font-size:smaller;font-weight:bold;">* Only members of OPS SUPER and TECH SUPER groups can edit/update</span><br>
<%'close record set and connection object
if v_mode="edit" then
  rsObj.Close
  Set rsObj=nothing
end if

connObj.Close
set connObj=nothing

if request("v_close")<>"true" then
%>

<SCRIPT LANGUAGE=javascript>
<!--
var v_mess="<%=request("v_message") %>"
if (v_mess!="") {
alert(v_mess);
}
//-->
</SCRIPT>
<%
end if
%>

</BODY>
</HTML>