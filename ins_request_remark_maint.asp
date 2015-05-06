<%	'===========================================================================
	'	Template Name	:	Inspection Request Remark Maintenance
	'	Template Path	:	ins_request_remark_maint.asp
	'	Functionality	:	To show the list of requests
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	10 September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	option explicit
	
	dim v_ins_request_id,VID,v_mess,v_filter,v_ctr,class_color,v_diff_days
	dim SQL,rs2
	
	v_ins_request_id=Request("v_ins_request_id")
	VID=Request("VID")
%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<script LANGUAGE="vbscript" RUNAT="Server">
function b13(field_with_chr_13)
	if isnull(field_with_chr_13) then
		b13=field_with_chr_13
	else
		b13=replace(replace(field_with_chr_13,chr(13),"<br>"),"  ","&nbsp;&nbsp;")
	end if
end function
</script>
<html>
<head>
<title>List of Follow-up Remarks</title>
<meta HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 2000 14:25:27 GMT"> 
<link REL="stylesheet" HREF="moc.css"></link>
<style>
h4
{
	margin:2px;
}
</style>
<script language="Javascript" src="js_date.js"></script>
<script language="VBScript" src="vb_date.vs"></script>
<script language="Javascript">
function cClose()
  		{
  			//var name= confirm("Are you sure?")
  			var name = true
  			if (name== true)
  			{
				//v_val = "ins_request_maint.asp?";
				//self.opener.document.form1.action=v_val;
				self.opener.document.form1.submit();
  				self.close() ;
  			}
  			else
  			{
  			return false;
  			}
  		}

function fncall(v_ins_request_id,v_remark_id,vessel_name,moc_name)
		{
			winStats='toolbar=no,location=no,directories=no,menubar=no,'
			winStats+='scrollbars=yes,resizable=yes'
			if (navigator.appName.indexOf("Microsoft")>=0) {				
				winStats+=',top=0, left=0, width='+screen.availWidth+', height='+screen.availHeight								
			}else{
				
				winStats+=',top=0, left=0, width='+screen.availWidth+', height='+screen.availHeight				
			}

			adWindow=window.open("ins_request_remark_entry.asp?v_ins_request_id="+v_ins_request_id+"&v_remark_id="+v_remark_id+"&vessel_name="+vessel_name+"&moc_name="+moc_name,"moc_request_remark_rem_entry",winStats);     
			adWindow.focus();
		}
		
function v_sort(v_sort_field,v_sort_order)
		{
		var v_ins_request_id='<%=v_ins_request_id%>'
		var vessel_name = '<%=request("vessel_name")%>'
		var moc_name = '<%=replace(request("moc_name"),"'","\'")%>'
		//alert('v_sort_field:'+v_sort_field)
		//alert('v_sort_order:'+v_sort_order)
		//alert(v_ins_request_id)
		document.form1.action="ins_request_remark_maint.asp?item="+v_sort_field+"&order="+v_sort_order+"&v_ins_request_id="+v_ins_request_id+"&vessel_name="+vessel_name+"&moc_name="+moc_name
		//alert(document.form1.action)
		document.form1.submit();
		}
</script>
</head>
<body class="bgcolorlogin">
<center>
<br>
<h4>List of Follow-up Remarks</h4>
<table width=90% border=0>
  <tr>
	<td align=right nowrap>
		<h4>Vessel:</h4>
	</td>
	<td nowrap>
		<h4><%=request("vessel_name")%> </h4>
	</td>
	<td align=right nowrap>
		<h4>MOC:</h4>
	</td>
	<td nowrap>
		<h4><%=request("moc_name")%></h4>
	</td>
  </tr>
  <tr>
	<td align=right nowrap>
		<h4>Insp. Date:</h4>
	</td>
	<td nowrap>
		<h4><%=request("v_insp_date")%></h4>
	</td>
	<td align=right nowrap>
		<h4>Insp. Port:</h4>
	</td>
	<td nowrap>
		<h4><%=request("v_insp_port")%></h4>
	</td>
  </tr>
  <tr>
	<td COLSPAN="4">
		<%  v_mess=Request.QueryString("v_message")
		if v_mess <> "" then
		%>
		<font color="red" size="+2"><%=v_mess%></font>
		<% end if%>
	</td>
  </tr>
  <tr>
	<td colspan=4 align="center">
<%
	if UserIsAdmin then
%>
		Click <a name="1" href="javascript:fncall('<%=v_ins_request_id%>','0','<%=request("vessel_name")%>','<%=replace(request("moc_name"),"'","\'")%>');">Here</a> to Create a New Remark
<%
	End If
%>
	</td>
</table>
<table width=100%>
  <tr>
	<td colspan=4 align="right">
	  <a href='http://webserve2/vid/create_data_file.asp?vessel_code=<%=VID%>&questionnaire_id=10000380' target='vessel_particulars'>
	  <span style='text-align:center;font-size:9;font-weight:bold;color:blue'>Vessel Particulars</span></a>

	  <a href="javascript:window.print()"><img src="Images/print.gif" border="0" alt="Print this Page" WIDTH="22" HEIGHT="20"></a>
	</td>
  </tr>
</table>
<%
v_filter =""
SQL = "Select remark_id, request_id,"
SQL = SQL & " subject || ' (Inspection dated " & request("v_insp_date") & ")' subject,"
SQL = SQL & " REMARK_TARGET_DATE, remarks,remark_pic,"
SQL = SQL & " to_char( REMARK_TARGET_DATE,'DD-Mon-YYYY')  REMARK_TARGET_DATE1,"
SQL = SQL & " nvl(remark_status,'Active')remark_status, to_char(create_date,'DD-Mon-YYYY') create_date1, created_by,"
SQL = SQL & " LAST_MODIFIED_DATE,LAST_MODIFIED_by ,  nvl(remark_target_date,sysdate)-sysdate diff_days"
SQL = SQL & " from MOC_REQUEST_REMARKS"
SQL = SQL & " where request_id=" & v_ins_request_id

SQL = SQL & " union"

SQL = " SELECT remark_id, mrr.request_id,"
SQL = SQL &  "        subject,"
SQL = SQL &  "        remark_target_date, remarks, remark_pic,"
SQL = SQL &  "        TO_CHAR (remark_target_date, 'DD-Mon-YYYY') remark_target_date1,"
SQL = SQL &  "        NVL (remark_status, 'Active') remark_status,"
SQL = SQL &  "        TO_CHAR (mrr.create_date, 'DD-Mon-YYYY') create_date1, mrr.created_by,"
SQL = SQL &  "        mrr.last_modified_date, mrr.last_modified_by,"
SQL = SQL &  "        NVL (remark_target_date, SYSDATE) - SYSDATE diff_days"
SQL = SQL &  "  FROM moc_request_remarks mrr, moc_inspection_requests mir"
SQL = SQL &  "  WHERE mrr.request_id = mir.request_id"
'SQL = SQL &  "  AND (mrr.remark_status='Active' or mrr.remark_status is null)"
SQL = SQL &  "  AND mir.request_id IN ("
SQL = SQL &  "           SELECT request_id"
SQL = SQL &  "            FROM moc_inspection_requests"
SQL = SQL &  "            WHERE vessel_code || '~' || moc_id ="
SQL = SQL &  "                                         (SELECT vessel_code || '~' || moc_id"
SQL = SQL &  "                                            FROM moc_inspection_requests"
SQL = SQL &  "                                           WHERE request_id = " & v_ins_request_id & "))"

if request("item")<>"" then
	SQL=SQL & " order by " & request("item") & " " & request("order")
else
	SQL=SQL & " order by remark_status, remark_target_date "
end if

Set rsObj=connObj.execute(SQL)

%> 
<form name="form1" action="ins_request_remark_delete.asp" method="post">

<input type="hidden" id="v_ins_request_id" name="v_ins_request_id" value="<%=Request("v_ins_request_id")%>">
<input type="hidden" id="vessel_name" name="vessel_name" value="<%=Request("vessel_name")%>">
<input type="hidden" id="moc_name" name="moc_name" value="<%=Request("moc_name")%>">
<table width="100%" ALIGN="LEFT">
  <tr> 
    <td class="tableheader" align="left" colspan="1">Remarks<br><br></td>
    <td class="tableheader" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;PIC&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
		<table width="100%" align="left">
		<tr>
			<td align="left">
			<a href="javascript:v_sort('remark_pic','asc');">
			<img SRC="Images/up.gif" ALT="Sort Ascending by Vessel" border="0" width="15" align="left" hspace="0">
			</a>
			</td>
			<td align="right">
			<a href="javascript:v_sort('remark_pic','desc');">
			<img SRC="Images/down.gif" ALT="Sort Descending by Vessel" border="0" width="15" align="right" hspace="0">
			</a>
			</td>
		</tr>
		</table>
    </td>
    <td class="tableheader" align="center">Target&nbsp;Date&nbsp;&nbsp;&nbsp;<br> 
		<table width="100%"><tr>
			<td align="left">
			<a href="javascript:v_sort('remark_target_date','asc');">
			<img SRC="Images/up.gif" ALT="Sort Ascending by MOC" width="15" border="0">
			</a>
			</td>
			<td align="right">
			<a href="javascript:v_sort('remark_target_date','desc');">
			<img SRC="Images/down.gif" ALT="Sort Descending by MOC" width="15" border="0">
			</a>
			</td>
		</tr></table>
    </td>
    <td class="tableheader" width="5%">Status<br>
		<table width="100%"><tr>
			<td align="left">
			<a href="javascript:v_sort('remark_status','asc');">
			<img SRC="Images/up.gif" ALT="Sort Ascending by Status" width="15" border="0">
			</a>
			</td>
			<td>
			<a href="javascript:v_sort('remark_status','desc');">
			<img SRC="Images/down.gif" ALT="Sort Descending by Status" width="15" border="0">
			</a>
			</td>
		</tr></table>
    </td>
  </tr>
<%
v_ctr=0
if not (rsObj.bof or rsObj.eof) then
while not rsObj.eof
v_ctr=v_ctr+1
	if (v_ctr mod 2) = 0 then
		class_color="columncolor2"
		else
		class_color="columncolor3"
	end if
%>
  <tr> 

    <td class="<%=class_color%>" align="left" valign="top">
     <%if rsObj("REMARK_STATUS")<>"Completed" then%>
     <a name="1" title="Click to Edit" href="javascript:fncall('<%=rsObj("request_id") %>','<%=rsObj("remark_id") %>','<%=request("vessel_name")%>','<%=replace(request("moc_name"),"'","\'")%>');">
	<font color="midnightblue"><b><u><%=rsObj("SUBJECT")%></u></b></font><br>
	<%=b13(rsObj("REMARKS"))%>
	</a>
	<%else%>
	<a style="color:black" name="1" title="Click to Edit" href="javascript:fncall('<%=rsObj("request_id") %>','<%=rsObj("remark_id") %>','<%=request("vessel_name")%>','<%=replace(request("moc_name"),"'","\'")%>');">
	<font color="midnightblue"><b><u><%=rsObj("SUBJECT")%></u></b></font><br>
	<%=b13(rsObj("REMARKS"))%>
	</a>
	<%end if%>
    </td>
    <td class="<%=class_color%>" valign="top"> 
    <%=rsObj("REMARK_PIC") %> 
    </td>
    <td class="<%=class_color%>" valign="top">
    <%=mid(rsObj("REMARK_TARGET_DATE1"),1,7)&mid(rsObj("REMARK_TARGET_DATE1"),10,11) %>&nbsp;
    
    <%
		v_diff_days=clng(rsObj("diff_days"))
		If rsObj("REMARK_STATUS")<>"Completed" and v_diff_days < 0 then
			Response.Write ("<IMG border=0 src='Images/red_triangle.gif' alt="  & v_diff_days & "> ")
		End If	
		%>
		&nbsp;
    </td>
    <td class="<%=class_color%>" valign="top"><%=rsObj("REMARK_STATUS") %>&nbsp;</td>
    
  </tr>

<%
rsObj.movenext
wend
else
Response.Write "<tr><td colspan=7 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if ' if not (rsObj.bof or rsObj.eof) then
%>
<tr>
	<td colspan="7" class="tabledata" align="center"> 
         <!--INPUT type="submit" value="Delete the Selected" id=submit1 name="submit1"--> 
         <input type="submit" value="Close" id="submit4" name="submit4" OnClick="javascript:return cClose();">
      </td>
<tr>
</table>
</form>
</body>
</html>
