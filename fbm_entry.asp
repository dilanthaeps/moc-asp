	<!--#include file="common_dbconn.asp"-->
<%
	'===========================================================================
	'	Template Name	:	Fleet Broadcast Message Entry
	'	Template Path	:	fbm_entry.asp
	'	Functionality	:	To allow the entry/edit of the Fleet Broadcase message details
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23rd September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	v_fbm_id=request("v_fbm_id")
	Response.Buffer = false
	Dim Idval
	if v_fbm_id <> "0" then
		v_mode="edit"
		v_header="Fleet Broadcast Message Details"

		strSql = "SELECT "
		strSql = strSql & "   A.FBM_ID , A.FROM_DEPT , to_char(A.DATE_SENT,'DD-Mon-YYYY') date_sent1 , A.PRIMARY_CATEGORY , A.SECONDARY_CATEGORY "
		strSql = strSql & "  , A.MSG_SUBJECT , A.MSG_BODY , A.STATUS , A.ATTACHEMENT , to_char(A.CREATE_DATE,'DD-Mon-YYYY') create_date1 "
		strSql = strSql & "  , A.CREATED_BY , to_char(A.LAST_MODIFIED_DATE ,'DD-Mon-YYYY') last_modified_date1, A.LAST_MODIFIED_BY "
		strSql = strSql & " FROM "
		strSql = strSql & " WLS_FBM A "
		strSql = strSql & " WHERE fbm_id ="&v_fbm_id
		'Response.Write strSql
		Set rsObj = connObj.Execute(strSql)
		strSql = "Select "
		strSql = strSql & "   A.DOCUMENT_ID , A.DESCRIPTION , A.DOC_PATH , A.UPLOADED_BY ,"
		strSql = strSql & " A.UPLOADED_DATE , A.FBM_ID "
		strSql = strSql & " FROM WLS_FBM_ATTACHEMENTS A where a.fbm_id='" &v_fbm_id&"'"
		set rsObj_docs= connObj.Execute(strSql)
	else
		v_mode="Add"
		v_header="Create Fleet Broadcast Message "
	end if

	strSql = "select para_code, para_value, upper(para_value) sort_column"
	strSql = strSql & " from wls_system_parameters"
	strSql = strSql & " where parent_id in ('DEPT', 'OFFDEPT')"
	strSql = strSql & " union"
	strSql = strSql & " select NULL para_code, ' ' para_value, 'AAAAAAA' sort_column"
	strSql = strSql & " from dual"
	strSql = strSql & " order by 3"

	'Response.Write strSql
	Set rsObj_dept = connObj.Execute(strSql)


	' Select List -      Remark_Status
	strSql = "SELECT sys_para_id, para_desc,parent_id,sort_order "
	strSql = strSql & " from moc_system_parameters "
	strSql = strSql & " where parent_id = 'Remark_Status' "
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
				v_val = "fbm_save.asp?";
				document.form1.action=v_val;
				document.form1.submit();
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
				v_val = "fbm_maint.asp";
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
	var name=confirm("Are you sure? This Fleet Broadcase message will be deleted!")
	if (name== true)
  			{
				v_val = "fbm_delete.asp?v_deleteditems="+v_delete_item;
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

			if (document.form1.from_dept.value == "")
			{
				alert ("Enter From department");
				document.form1.from_dept.focus();
				return false;
			}
			if (document.form1.date_sent.value == "")
			{
				alert ("Enter Date to send");
				document.form1.date_sent.focus();
				return false;
			}
			if (document.form1.primary_category.value == "")
			{
				alert ("Enter Primary Category");
				document.form1.primary_category.focus();
				return false;
			}
			if (document.form1.secondary_category.value == "")
			{
				alert ("Enter Secondary Category");
				document.form1.secondary_category.focus();
				return false;
			}
			if (document.form1.msg_subject.value == "")
			{
				alert ("Enter Message Subject");
				document.form1.msg_subject.focus();
				return false;
			}
			if (document.form1.msg_body.value == "")
			{
				alert ("Enter Message Body");
				document.form1.msg_body.focus();
				return false;
			}
			if(document.form1.msg_body.value.length>4000)
			{
				alert("Msg Body Maximum 4000 Chars only");
				document.form1.msg_body.focus();
				return false;
			}
			if (document.form1.status.value == "")
			{
				alert ("Enter Message Status");
				document.form1.status.focus();
				return false;
			}

  	}
  	function docsmaint(v_fbm_id)
  		{
		var w = document.body.clientWidth;
		winStats='toolbar=no,location=no,directories=no,menubar=no,'
		winStats+='scrollbars=yes'
		if(w<=800){
			winStats+=',left=30,top=50,width=650,height=215'
		} else {
		winStats+=',left=120,top=50,width=650,height=210'
		}
		adWindow=window.open("docs_maint.asp?v_fbm_id="+v_fbm_id,"docs_entry",winStats);
		adWindow.focus();
	}
	function fncall(v_fbm_id)
	{
		document.form1.action="docs_delete.asp?v_proj_id="+v_proj_id;
		document.form1.submit();
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
<TITLE>Fleet Broadcase Message Entry/Edit Screen</TITLE>
</HEAD>
<BODY class=bcolor>
<h3><%= v_header %></h3>
<form name="form1" action=fbm_save.asp method=post onsubmit="javascript:return validate_fields();">

<%
if v_mode = "edit" then
%>
<div align="right">
<INPUT type="button" id="delete" name="delete" value="Delete" onclick="Javascript:return fndelete(<%=SFIELD("fbm_id")%>);">
<INPUT type="hidden" id=v_fbm_id name=v_fbm_id value="<%=SFIELD("fbm_id")%>">
</div>
<%
end if
%>
<table >
<TR>
	<TD class="tableheader">
		From Dept.
	</TD>
	<TD>
		<SELECT id=from_dept name=from_dept>
		<%
			While Not rsObj_dept.EOF

					v_selected = ""
					If Trim(SFIELD("from_dept")) = rsObj_dept("para_code") Then
						v_selected = "SELECTED"
					End If

					Response.Write "<OPTION VALUE='" & rsObj_dept("para_code") & "' " & v_selected & ">"
					Response.Write rsObj_dept("para_value")
					Response.Write "</OPTION>"
					rsObj_dept.MoveNext
				Wend
		%>
		</SELECT>
		<font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		Date Sent
	</TD>
	<TD>
		 <input type="text" name="date_sent" value="<%=SFIELD("date_sent1")%>" onblur="vbscript:valid_date date_sent,'Valid Date','form1'">
              <A HREF="javascript:show_calendar('form1.date_sent',form1.date_sent.value);">
              <IMG SRC="Images/calendar.gif" alt="Pick Date from Calendar"  WIDTH="20" HEIGHT="18" BORDER="0">
              </A>
		<font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		Primary Category
	</TD class="tabledata">
	<TD>
		<SELECT id=primary_category name=primary_category>
			<OPTION value=primary selected>primary</OPTION>

		</SELECT>
		<font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		Secondary Category
	</TD class="tabledata">
	<TD>
		<SELECT id=secondary_category name=secondary_category>
			<OPTION value=secondary selected>secondary</OPTION>

		</SELECT>
		<font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		Message Subject
	</TD>
	<TD>
		 <INPUT type="text" id=msg_subject name=msg_subject size=60 value="<%=SFIELD("msg_subject")%>">
		 <font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		Message Body
	</TD>
	<TD>
		<TEXTAREA rows=8 cols=50 id="msg_body" name="msg_body"><%=SFIELD("msg_body")%></TEXTAREA>
		<font color=red>*</font>
	</TD>
</TR>

<TR>
	<TD class="tableheader">
		STATUS
	</TD>
	<TD class="tabledata">
	 	<select name="status" >
 			<option value="">Select Status</option>
				<%
					if not(rsObj_Remark_Status.eof or rsObj_Remark_Status.bof) then
						while not rsObj_Remark_Status.eof
				%>
						<option value="<%=rsObj_Remark_Status("sys_para_id")%>" <%if not(isnull(SFIELD("status"))) then%><%if cstr(rsObj_Remark_Status("sys_para_id"))=cstr(SFIELD("status")) then%>selected<%end if%><%end if%>><%=rsObj_Remark_Status("para_desc")%></option>
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
		Attachement
	</TD>
	<TD>
		 <a href="javascript:docsmaint(<%=SFIELD("fbm_id")%>);">Attach file </a><br>
		 <%
		 if not (rsObj_docs.eof or rsObj_docs.bof) then
		 Response.Write "<STRONG>List of Attachements</STRONG> <br>"
			while not rsObj_docs.eof
			Response.Write rsObj_docs("description") & "<br>"
			rsObj_docs.movenext
			wend
		end if
		%>
	</TD>
</TR>

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
	</TD class="tabledata">
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
         <INPUT type="submit" value="Save" id=submit1 name="submit1"> &nbsp;
         <INPUT type="submit" value="Save and Close" id=submit4 name="submit4"  OnClick="javascript:return sClose();"> &nbsp;
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
  rsObj_docs.close
  set rsObj_docs = nothing
end if
rsObj_dept.close
set rsObj_dept = nothing
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
