<!--#include file="common_dbconn.asp"-->
<html>
<head>
<link rel="stylesheet" href="moc.css"></link>
   <title>TPM - Document Maintenance</title>
   <script language="javascript">
  	function docsentry(v_fbm_id,v_document_id)
	{
		var w = document.body.clientWidth; 
		winStats='toolbar=no,location=no,directories=no,menubar=no,'
		winStats+='scrollbars=yes'
		if(w<=800){
			winStats+=',left=30,top=50,width=650,height=215'
		} else {
		winStats+=',left=120,top=50,width=650,height=210'
		}
		adWindow=window.open("docs_entry.asp?v_fbm_id="+v_fbm_id+"&v_document_id="+v_document_id,"docs_entry1",winStats);     
		adWindow.focus();
	}
	function fncall(v_proj_id)
	{
		document.form1.action="docs_delete.asp?v_proj_id="+v_proj_id;
		document.form1.submit();
	}
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

   </script>
</head>
<title>Fleet Broadcast Message - Document Maintenance</title>
<body class=bgcolorlogin>
<center>
<h4>Documents</h1>
<p></p>
<% 
	if request("v_fbm_id")<>"" then
		v_fbm_id=request("v_fbm_id")
	end if
    v_mess=Request.QueryString("v_message")
	if v_mess <> "" then
%>	
   <font color=red size=+2><%=v_mess%></font>
<% end if%>
<p>   
Click <a href="Javascript:docsentry('<%=v_fbm_id%>','');">Here</a> to Create a New Document
<%      
   strSql="Select document_id,Description,doc_path from wls_fbm_attachements where fbm_id='"&v_fbm_id&"'"
   'Response.Write strSql
   'Response.end
   Set rsObj=connObj.Execute(strSql)
%>
<form name=form1 action=cons_delete.asp method=post>
<INPUT type="hidden" id=v_fbm_id name=v_fbm_id value='<%=v_fbm_id %>'>
	<table>
		<tr>
			<td class=tableheader>Description</td>
			<td class=tableheader>Path</td>
			<td class=tableheader>View</td>
		</tr>	
		<% if not(rsObj.eof or rsObj.bof) then
		dim c,cclass,r_count
		c=0
		r_count=0
		while not rsObj.EOF 
		if c=0 then
		cclass="columncolor"
		c=1
		elseif c=1 then
		cclass="columncolor1"
		c=0
		end if
		%>
		   
		<tr>
			<td class="<%=cclass%>">
				<input type=checkbox name=v_deleteditems value='<%=rsObj("document_id")%>'>&nbsp;&nbsp;
				<a href="Javascript:docsentry('<%=v_fbm_id%>','<%=rsObj("document_id")%>');"><%=rsObj("description")%></a>
			</td>
			<td class="<%=cclass%>"><%=rsObj("doc_path")%>&nbsp;</td>
			<%
				if rsObj("doc_path")<>"" then
				v_path="//"&Request.ServerVariables("server_name") & rsObj("doc_path")
				'Response.Write v_path
				end if
				
			%>
			<td class="<%=cclass%>"><a href="<%=v_path%>" target="aa">View</a>&nbsp;</td>
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
  <p></p>
  <input type=button  value="Delete"  class=cmdbutton onClick="Javascript:fncall('<%=v_fbm_id%>');">&nbsp;
  <input type=button  value="Close"  class=cmdbutton onClick="Javascript:cClose();" id=button1 name=button1>
       <!--#include file="common_footer.asp"-->  
</form>
</table>    
