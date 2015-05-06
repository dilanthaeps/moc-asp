<%option explicit%>
<%
'===========================================================================
'	Template Name	:	PO List
'	Template Path	:	po_list.asp
'	Functionality	:	To choose the po from the list of DAs in DANAOS
'	Called By		:	.
'	Created By		:	Prashant Kumar
'   Create Date		:	18th January, 2006
'	Update History	:
'						1.
'						2.
'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<html>
<head>
<title>PO List - Tanker Pacific</title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="Javascript">
	function fncall(po)
	{
		window.opener.document.form1.PO_NUMBER.value=po;
		self.close();
	}
</script>
</head>
<body>
<center>
<h4> PO List</h4>
<p></p>

<%
dim SQL,REQUESTID
REQUESTID = Request.QueryString("REQUESTID")

SQL = " SELECT internal_ref, port_name, TO_CHAR (eta, 'dd-Mon-yyyy') eta, action"
SQL = SQL &  " FROM expected_call_da eda, moc_inspection_requests mir"
SQL = SQL &  " where eda.vessel_code = mir.vessel_code"
SQL = SQL &  " and eda.port_name like '%' || mir.inspection_port || '%'"
SQL = SQL &  " and trunc(eda.eta) <= trunc(mir.inspection_date)"
SQL = SQL &  " and mir.request_id=" & REQUESTID
SQL = SQL &  " ORDER BY eda.eta DESC"

Set rsObj=connObj.Execute(sql)
%>

<form name=form1 method=post>
	<table>
		<tr>
			<td class=tableheader>PO Number</td>
			<td class=tableheader>Port</td>
			<td class=tableheader>Date</td>
			<td class=tableheader>Activity</td>
		</tr>	
		<%
		if not rsObj.eof then
			dim c,cclass,r_count
			c=0
			r_count=0
			while not rsObj.EOF 
				if c=0 then
					cclass="tabledata"
					c=1
				elseif c=1 then
					cclass="tabledataalt"
					c=0
				end if
		%>
		   
		<tr>
			<td class="<%=cclass%>">
				<a href="Javascript:fncall('<%=rsObj("internal_ref")%>')"><%=rsObj("internal_ref")%>&nbsp;</a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("port_name")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("eta")%>&nbsp;</td>
			<td class="<%=cclass%>"><%=rsObj("action")%>&nbsp;</td>
	    </tr>
		<%		rsObj.MoveNext
				r_count=r_count+1
			wend
			Response.Write "<tr><td colspan=3 align=left><b>Record Count :</b>"&r_count&"</td></tr>"
         else
			Response.Write "<tr><td colspan=3 align=center><b> No Data Found </b></td></tr>"
         end if
       %>  

  </table>
  
  
  <input type=BUTTON name=ebut value=Exit OnClick="Javascript:parent.close();">
  <!--#include file="common_footer.asp"-->  
</form>
</center>
</table>    
</body>
</html>