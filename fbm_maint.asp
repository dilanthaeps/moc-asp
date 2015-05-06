<%	'===========================================================================
	'	Template Name	:	Fleet Broadcast Message Maintenance
	'	Template Path	:	fbm_maint.asp
	'	Functionality	:	To show the list of messages
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	23 September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	Response.Buffer = false
%>
<!--#include file="common_dbconn.asp"-->
<SCRIPT LANGUAGE=vbscript RUNAT=Server>
function b13(field_with_chr_13)
	if isnull(field_with_chr_13) then
		b13=field_with_chr_13
		else
		b13=replace(replace(field_with_chr_13,chr(13),"<br>"),"  ","&nbsp;&nbsp;")
	end if
end function
</SCRIPT>
<html>
<head>
<Title>Fleet Broadcast Messages - Tanker Pacific </Title>
<META HTTP-EQUIV="expires" CONTENT="Tue, 20 Aug 2000 14:25:27 GMT"> 
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="Javascript" src="js_date.js">
</script>
<script language="VBScript" src="vb_date.vs">
</script>
<script language="Javascript" >
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

function fncall(v_fbm_id)
		{
			winStats='toolbar=no,location=no,directories=no,menubar=no,'
			winStats+='scrollbars=yes,resizable=yes'
			if (navigator.appName.indexOf("Microsoft")>=0) {
				winStats+=',left=120,top=10,width=720,height=650'
			}else{
				winStats+=',screenX=350,screenY=200,width=400,height=280'
			}
			adWindow=window.open("fbm_entry.asp?v_fbm_id="+v_fbm_id,"fbm_entry",winStats);     
			adWindow.focus();
		}
		
function v_sort(v_sort_field,v_sort_order)
		{
		//alert('v_sort_field:'+v_sort_field)
		//alert('v_sort_order:'+v_sort_order)
		//alert(v_ins_request_id)
		document.form1.action="fbm_maint.asp?item="+v_sort_field+"&order="+v_sort_order
		//alert(document.form1.action)
		document.form1.submit();
		}
</script>
</head>
<body class=bgcolorlogin>
<center>
<table >
<tr>
		<td width=95% align=center colspan=2>
			<p>
			<h4>Fleet BroadCast Messages</h4>
			<p></p>
			<%  v_mess=Request.QueryString("v_message")
			if v_mess <> "" then
			%>
			<font color=red size=+2><%=v_mess%></font>
			<% end if%>
			<p>
		<td>
		<td width=5% align=right>
		&nbsp;
		</td>
</tr>
	<tr>
		<td width=90% align=center>
		Click <a name=1 href="javascript:fncall('0');">Here</a> to Create a New Fleet Broadcase Message
		</td>
		<td width=5%>
		&nbsp;
		<td>
		<td width=5% align=right>
		<a href="javascript:window.print()"><img src="Images/print.gif" border=0 alt="Print this Page"></a>
		</td>
	</tr>
</table>
<%
v_filter =""
strSql = "Select fbm_id, from_dept, to_char( date_sent,'DD-Mon-YYYY')  date_sent1 "
strSql = strSql & ", primary_category,secondary_category,msg_subject, msg_body, status,  "
strSql = strSql & "attachement, LAST_MODIFIED_by , to_char( create_date,'DD-Mon-YYYY')  create_date1 "
strSql = strSql & " from wls_fbm "
strSql = strSql & " where 1=1 "


if request("item")<>"" then
strSql=strSql & " order by "&request("item")& " "&request("order") 
else
strSql=strSql & " order by date_sent"
end if

'Response.Write strSql
Set rsObj=connObj.execute(strSql)



%> 
<!--h3 align="center">Selection Filter </h3--> 

<form id=form2 name=form2 method=post action=fbm_maint.asp>
<table width="60%" border="0" cellspacing="1" cellpadding="1" align="center">
  <tr> 
    <td class="tableheader">
      <div align="center">Department</div>
    </td>
    <td class="tableheader">
      <div align="center">Year</div>
    </td>
    <td class="tableheader">
      <div align="center">Subject Having</div>
    </td>
    <td class="tableheader">
      <div align="center">Primary Category</div>
    </td>
    <td class="tableheader">
      <div align="center">Seconday Category</div>
    </td>
    <td class="tableheader">
      <div align="center">Status</div>
    </td>
  </tr>
  <tr> 
	  
      <td class="tabledata"> 
        <div align="center">

             <select name="vessel_code" onchange="javascript:v_select('vessel_code');">
             <option value="">Select Department</option>
				
			</select>

        </div>
      </td>
      <td class="tabledata"> 
        <div align="center">
        
          <select name="moc_id" onchange="javascript:v_select('moc_id');">
             <option value="">Select Year</option>
	
			</select>
			
        </div>
      </td>
      <td class="tabledata"> 
        <div align="center">
           <INPUT type="text" id=search_msg_subject name=search_msg_subject>
        </div>
      </td>
      <td class="tabledata"> 
        <div align="center">
          <select name="inspection_port" onchange="javascript:v_select('inspection_port');">
             <option value="">Select Primary Category</option>
	
			</select>
        </div>
      </td>
      <td class="tabledata"> 
        <div align="center">
          <select name="inspection_port" onchange="javascript:v_select('inspection_port');">
             <option value="">Select Secondary Category</option>
	
			</select>
        </div>
      </td>
      <td class="tabledata"> 
        <div align="center">
          <select name="inspection_port" onchange="javascript:v_select('inspection_port');">
             <option value="">Select Status</option>
	
			</select>
        </div>
      </td>
      
  </tr>
  <tr>
	<td colspan="5" align="center">
	<INPUT type="submit" value="Apply Filter" id=submit1 name=submit1 class=cmdButton>
	<INPUT type="button" value="Clear All Filters"   onclick="Javascript:v_clear_all_filters();" class=cmdButton id=button1 name=button1>
	</td>
  </tr>
</table>
</form>


<form name="form1" action="" method="post">

<table width=100%>
<tr>
<td align=left>&nbsp; </td><td align="right"><STRONG>Date/Time:</STRONG> <% Response.write day(Now())&"-"&mid(monthname(month(now())),1,3)&"-"&year(now())&" "&hour(now())&":"&minute(now())&":"&second(now()) %> </td>
</tr>
</table>
<table  ALIGN="center">
  <tr> 
    <td class="tableheader"   align="left" colspan="1" >
		 
		<a href="javascript:v_sort('from_dept','<% if request("order")="asc" then Response.write "desc" else Response.write "asc" %>');">
			<font color=white>Department</font>
		</a>
    </td>
    <td class="tableheader" align="center">
		<a href="javascript:v_sort('fbm_id','<% if request("order")="asc" then Response.write "desc" else Response.write "asc" %>');">
		<font color=white>FBM No.</font>
		</a> 
    </td>
    <td class="tableheader" align="center"> 
    <a href="javascript:v_sort('date_sent','<% if request("order")="asc" then Response.write "desc" else Response.write "asc" %>');">
		<font color=white>Year</font>
		</a> 
    </td>
    <td class="tableheader" align="center"> 
    <a href="javascript:v_sort('msg_subject','<% if request("order")="asc" then Response.write "desc" else Response.write "asc" %>');">
		<font color=white>Subject</font>
		</a> 
    </td>
    <td class="tableheader" align="center">Quick View 
    </td>
    <td class="tableheader" align="center"> 
    <a href="javascript:v_sort('date_sent','<% if request("order")="asc" then Response.write "desc" else Response.write "asc" %>');">
		<font color=white>Date Sent</font>
		</a> 
    </td>
    <td class="tableheader" align="center">Date Vessel Ack. 
    </td>
    <td class="tableheader" align="center">Attachement 
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

    <td class="<%=class_color%>" align="left"  valign=top>
     <%=trim(rsObj("from_dept")) %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("fbm_id") %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("create_date1") %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
    <a href="javascript:fncall('<%=rsObj("fbm_id")%>');">
     <%=rsObj("msg_subject") %>
    </a>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("fbm_id") %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("date_sent1") %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("fbm_id") %>
    </td>
    <td class="<%=class_color%>" align="left"  valign=top>
     <%=rsObj("attachement") %>
    </td>
    
  </tr>

<%
rsObj.movenext
wend
else
Response.Write "<tr><td colspan=8 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if ' if not (rsObj.bof or rsObj.eof) then
%>
<TR>
	<td colspan="8" class="" align=center> 
         <!--INPUT type="submit" value="Delete the Selected" id=submit1 name="submit1"--> 
         <INPUT type="submit" value="Close" id=submit4 name="submit4"  OnClick="javascript:self.close();">
      </td>
<TR>
</table>
</form>
</body>
</html>
