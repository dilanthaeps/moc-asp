<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	'===========================================================================
	'	Template Name	:	MOC Inspection Request Remark Entry
	'	Template Path	:	ins_request_remark_entry.asp
	'	Functionality	:	To allow the entry/edit of the MOC Inspection Request Remarks details
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	10th September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	v_ins_request_id=request("v_ins_request_id")
	Response.Buffer = false

	Function IIF(expr, trueVal, falseVal)
		If expr Then
			IIF = trueVal
		Else
			IIF = falseVal
		End If
	End function

	v_button_disabled = "DISABLED"
	'If getAppVar("ACCESS_LEVEL") = "USRADM" Or getAppVar("ACCESS_LEVEL") = "USRMOCADM" Then
	if UserIsAdmin then
		v_button_disabled = ""
	End If

	Dim Idval
	Idval = Request.QueryString("v_remark_id")
	if Idval <> "0" then
		v_mode="edit"
		v_header="Follow-up Remarks"
		strSql = "Select "
		strSql = strSql & "remark_id, request_id, remarks, remark_pic, "
		strSql = strSql & "to_char(remark_target_date,'DD-Mon-YYYY') remark_target_date1, remark_status, "
		strSql = strSql & "to_char(create_date, 'DD-Mon-YYYY') create_date1, created_by, "
		strSql = strSql & "to_char(last_modified_date,'DD-Mon-YYYY') last_modified_date1, last_modified_by, subject"
		strSql = strSql & " from moc_request_remarks "
		strSql = strSql & " WHERE remark_id ="&Idval
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Create Remark "
	end if

	' Select List -      Person Incharge
	SQL = " SELECT distinct oum.user_id, oum.user_name, oum.user_name email"
	SQL = SQL &  "   FROM oca_user_master oum, oca_user_user_grp_asgn oug, oca_user_group_master ogm"
	SQL = SQL &  "  WHERE (    (oum.user_id = oug.user_id)"
	SQL = SQL &  "         AND (oug.user_grp_id = ogm.user_grp_id)"
	SQL = SQL &  "         AND (oum.status = 'ACTIVE')"
	SQL = SQL &  "         AND (upper(ogm.user_grp_name) in ('MARINE SUPER','OPS SUPER','TECH SUPER'))"
	SQL = SQL &  "        )"
	SQL = SQL &  "  order by upper(oum.user_name)"

	set rsObj_pic = connObj.Execute(SQL)

	' Select List -      Remark_Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & "from moc_system_parameters "
	strSql = strSql & "where parent_id = 'Remark_Status' "
	strSql = strSql & "order by sort_order "
	'Response.Write strSql
	set rsObj_Remark_Status = connObj.Execute(strSql)
%>
<html>
<head>
<META HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 1996 14:25:27 GMT">
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="VBScript" runat=server>
	   function SFIELD(fname)
	      if v_mode="edit" then
			'if not (rsObj.bof or rsObj.eof) then
				rsObj.MoveFirst
	      		Do Until rsObj.EOF
					v_tem = rsObj(cstr(fname))
					rsObj.MoveNext
				Loop
			'end if
	         SFIELD=v_tem
	      else
	         SFIELD = ""
	      end if
	   End function
</script>
<script language="Javascript" src="js_date.js">
</script>
<script language="VBScript" src="vb_date.vs">
</script>
<SCRIPT LANGUAGE="JAVASCRIPT">
  	function sClose()
  		{
  			//var name= confirm("Are you sure? The changes will be saved and !!")
  			var name = true
  			if (name== true)
  			{
				v_val = "ins_request_remark_save.asp?";
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
				v_val = "ins_request_remark_maint.asp";
				self.opener.document.form1.action=v_val;
				self.opener.document.form1.submit();
  				self.close() ;
  			}
  			else
  			{
  			return false;
  			}
  		}
	function fndelete(v_delete_item)
	{
	var name=confirm("Are you sure? This Remark record will be deleted!")
	if (name== true)
  			{
				v_val = "ins_request_remark_delete.asp?v_deleteditems="+v_delete_item;
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

			if (document.form1.remarks.value == "")
			{
				alert ("Enter Remarks");
				document.form1.remarks.focus();
				return false;
			}
			if(document.form1.remarks.value.length>4000)
			{
				alert("Remarks Maximum 4000 Chars only");
				document.form1.remarks.focus();
				return false;
			}
			if (document.form1.remark_pic.value == "")
			{
				alert ("Enter Person Incharge");
				document.form1.remark_pic.focus();
				return false;
			}
			if (document.form1.REMARK_TARGET_DATE.value == "")
			{
				alert ("Enter Remark Target Date");
				document.form1.REMARK_TARGET_DATE.focus();
				return false;
			}
			if (document.form1.REMARK_STATUS.value == "")
			{
				alert ("Enter Remark Status");
				document.form1.REMARK_STATUS.focus();
				return false;
			}
  	}
</SCRIPT>
<%
if request("v_close") = "true" then
%>
<SCRIPT LANGUAGE=javascript>
<!--
cClose()
//-->
</SCRIPT>

<%
end if
%>
<TITLE> MOC Inspection Follow-up Remarks Entry/Edit Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<h4>Vessel : <%=request("vessel_name")%>  &nbsp; &nbsp; &nbsp; &nbsp;MOC &nbsp; &nbsp;: <%=request("moc_name")%></h4>
<form name="form1" action=ins_request_remark_save.asp method=post onsubmit="javascript:return validate_fields();">
<INPUT type="hidden" id="remark_id" name="remark_id" value="<%=SFIELD("remark_id")%>">
<INPUT type="hidden" id="request_id" name="request_id" value="<%=v_ins_request_id%>">
<INPUT type="hidden" id="v_ins_request_id" name="v_ins_request_id" value="<%=v_ins_request_id%>">
<INPUT type="hidden" id="vessel_name" name="vessel_name" value="<%=request("vessel_name")%>">
<INPUT type="hidden" id="moc_name" name="moc_name" value="<%=request("moc_name")%>">
<%
if v_mode = "edit" then
%>
<div align="right">
<INPUT type="button" id="delete" name="delete" value="Delete" <% =v_button_disabled %>
	onclick="Javascript:return fndelete(<%=SFIELD("remark_id")%>);">
</div>
<%
end if
%>
<table WIDTH="100%">
<tr>
	<TD class="tableheader">
		SUBJECT
	</TD>
	<TD class="tabledata">
		<input type=text name="subject" maxlength=500 size="120" value="<%=server.HTMLEncode(SFIELD("SUBJECT"))%>">
		<font color=red>*</font>
<TR>
	<TD class="tableheader">
		REMARKS
	</TD>
	<TD class="tabledata" nowrap>
		<TEXTAREA rows=16 cols=100 id="remarks" name="REMARKS"><%=SFIELD("REMARKS")%></TEXTAREA>
		<font color=red>*</font>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		PERSON IN CHARGE
	</TD>
	<TD class="tabledata">


	 	<select name="remark_pic" >
 			<option value="">Select PIC</option>
				<%
					if not(rsObj_pic.eof or rsObj_pic.bof) then
						while not rsObj_pic.eof
				%>
						<option value="<%=rsObj_pic("user_id")%>" <%if not(isnull(SFIELD("REMARK_PIC"))) then%><%if cstr(rsObj_pic("user_id"))=cstr(SFIELD("REMARK_PIC")) then%>selected<%end if%><%end if%>><%=rsObj_pic("user_name")%></option>
				<%		rsObj_pic.movenext
						wend
					end if
				%>
			</select>
			<font color=red>*</font>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		REMARK TARGET DATE
	</TD>
	<TD class="tabledata">
	 <input type="text" name="REMARK_TARGET_DATE" value="<%=SFIELD("REMARK_TARGET_DATE1")%>" onblur="vbscript:valid_date REMARK_TARGET_DATE,'Remark Target Date','form1'">
        <A HREF="javascript:show_calendar('form1.REMARK_TARGET_DATE',form1.REMARK_TARGET_DATE.value);">
				<IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0">
		</A>
		<font color=red>*</font>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		STATUS
	</TD>
	<TD class="tabledata">
	 	<select name="REMARK_STATUS" >
 			<option value="">Select Status</option>
				<%
					if not(rsObj_Remark_Status.eof or rsObj_Remark_Status.bof) then
					
						while not rsObj_Remark_Status.eof

							v_selected = ""
							If SFIELD("REMARK_STATUS") <> "" Then
								If CStr(rsObj_Remark_Status("sys_para_id")) = CStr(SFIELD("REMARK_STATUS")) Then
									v_selected = "SELECTED"
								End If
							Else
								If UCase(rsObj_Remark_Status("sys_para_id")) = "ACTIVE" Then
									v_selected = "SELECTED"
								End If
							End If	
				%>
						<OPTION VALUE="<%=rsObj_Remark_Status("sys_para_id")%>" <% =v_selected %>><%=rsObj_Remark_Status("para_desc")%></OPTION>
				<%		rsObj_Remark_Status.movenext
						wend
					end if
				%>
		</select>
		<font color=red>*</font>
	</TD>
</TR>
<%
if v_mode = "edit" then
%>
<TR>
	<TD class="tableheader">
		CREATE DATE
	</TD>
	<TD class="tabledata">
	<%=SFIELD("CREATE_DATE1")%>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		CREATED BY
	</TD>
	<TD class="tabledata">
	<%=SFIELD("CREATED_BY")%>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		LAST MODIFICATION DATE
	</TD>
	<TD class="tabledata">
	<%=SFIELD("LAST_MODIFIED_DATE1")%>
	</TD>
</TR>
<TR>
	<TD class="tableheader">
		LAST MODIFIED BY
	</TD>
	<TD class="tabledata">
	<%=SFIELD("LAST_MODIFIED_BY")%>
	</TD>
</TR>
<%
end if
%>
<TR>
	<td colspan="2" class="tabledata" align=center>
         <INPUT type="submit" value="Save" id=submit1 <% =v_button_disabled %> name="submit1"> &nbsp;
         <INPUT type="submit" value="Save and Close" id=submit4 name="submit4" <% =v_button_disabled %> OnClick="javascript:return sClose();"> &nbsp;
         <INPUT type="button" value="Close without Save" id=back name="back" onclick="javascript:return cClose();"> &nbsp;
      </td>
</TR>
</table>
</form>
<font color=red>*</font> Denotes Mandatory Field.<br>
<%'close record set and connection object
if v_mode="edit" then
  rsObj.Close
  Set rsObj=nothing
end if

connObj.Close
set connObj=nothing
if request("v_close") <> "true" then
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
