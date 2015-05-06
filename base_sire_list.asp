   <!--#include file="common_dbconn.asp"-->
 <%	
 '===========================================================================
	'	Template Name	:	List of Inspection based on SIRE
	'	Template Path	:	base_sire_list.asp
	'	Functionality	:	To choose the base inspection  from the list of Inspection
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	21st September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
	%>
<html>
<head>
<title>Base SIRE List - Tanker Pacific</title>
<LINK REL="stylesheet" HREF="moc.css"></LINK>
<script language="Javascript">
	function fncall(request_id,moc_name, insp_date, insp_port)
	{
		if(request_id==0)
		{
			request_id='';
		}
		opener.document.form1.basis_sire.value = request_id;
		opener.document.form1.basis_sire_moc_name.value = moc_name;
		opener.document.form1.inspection_date.value = insp_date;
		opener.document.form1.inspection_port.value = insp_port;
		self.close();
	}
</script>
</head>
<body>
<center>
<h4> Requests List</h4>
<p></p>

<% 
strSql = "SELECT  a.request_id , A.MOC_ID , to_char(A.INSPECTION_DATE,'DD-Mon-YY') INSPECTION_DATE1 , A.INSPECTION_PORT, wls_fn_moc_name(a.moc_id) moc_name"
strSql = strSql & " FROM  MOC_INSPECTION_REQUESTS A"
strSql = strSql & " where a.status = 'ACTIVE'"

strSql = strSql & " and a.insp_status in ('ACCEPTED','SIRE REPORT REPLIED')"
strSql = strSql & " and a.insp_type='MOC' and is_sire='Y'"
if request("vessel_code") <> "" then
	strSql=strSql & " and vessel_code= '" & request("vessel_code") & "'" & " and request_id<>"&request("request_id")
end if
strSql=strSql & " order by a.INSPECTION_DATE desc"
'Response.write strSql
Set rsObj=connObj.Execute(strSql)
%>

<form name=form1 action= method=post>
	<table>
		<tr>
			<td class=tableheader>MOC Name</td>
			<td class=tableheader>Inspection Date</td>
			<td Class=tableheader>Port</td>
		</tr>	
		<% if not(rsObj.eof or rsObj.bof) then
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
				<a href="Javascript:fncall('<%=rsObj("request_id")%>','<%=replace(rsObj("moc_name"),"'","\'")%>', '<% =rsObj("inspection_date1") %>', '<% =rsObj("inspection_port") %>')"><%=rsObj("moc_name")%>&nbsp;</a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("INSPECTION_DATE1")%>&nbsp;</td>
			<td class="<%=cclass%>">
				 <%=rsObj("INSPECTION_PORT")%>  &nbsp;  
			</td>

	    </tr>
       <%rsObj.MoveNext
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