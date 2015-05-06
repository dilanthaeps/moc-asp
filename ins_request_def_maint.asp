<%option explicit%>
<%	'===========================================================================
	'	Template Name	:	Inspection Request Deficiency Maintenance
	'	Template Path	:	ins_request_def_maint.asp
	'	Functionality	:	To show the list of requests deficiency
	'	Called By		:	.
	'	Created By		:	Sethu Subramanian R, Tecsol Pte Ltd, Singapore
	'   Create Date		:	12 September, 2002
	'	Update History	:
	'						1.
	'						2.
	'===========================================================================
%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="common_procs.asp"-->
<%	Response.Buffer = true
    dim rsFleetName
	dim DISPLAY_LINK_TO_WORKLIST
	'to hide link to Worklist, make DISPLAY_LINK_TO_WORKLIST = false
	DISPLAY_LINK_TO_WORKLIST = true
	'DISPLAY_LINK_TO_WORKLIST = false
	
	dim v_ins_request_id,sRisk,v_button_disabled
	
	v_button_disabled = "DISABLED"
	if UserIsAdmin then
		v_button_disabled = ""
	End If
	v_ins_request_id=Request("v_ins_request_id")
	
	dim strSqlVessMail,rsObjVessMail,VessMailBody,VessMailBody_SIRE,v_send_to
	dim srNo,strSqlCdrInfo,rsObjCdrInfo,v_coordinator_name,v_coordinator_design
	dim v_mess,v_filter,v_ctr,class_color
	dim ENTRY_TYPE,v_vessel_code,v_vessel_name,MOC_NAME,INSP_DATE,INSP_PORT,FLEET_NAME,MOC_CCList,V_SHORT_NAME
	dim SQL,rs,subject
	dim sTemp,nLen,nPos
	
	
	SQL = "Select ir.vessel_code,v.vessel_name,v.vessel_short_name,entry_type,mm.short_name,inspection_port,inspection_grade,"
	SQL = SQL & " to_char(inspection_date,'DD-Mon-YYYY')inspection_date_disp"
	SQL = SQL & " from moc_master mm,moc_inspection_requests ir, vessels v"
	SQL = SQL & " where ir.vessel_code=v.vessel_code and mm.moc_id=ir.moc_id"
	SQL = SQL & " and ir.request_id=" & v_ins_request_id
	set rs = connObj.execute(SQL)
	ENTRY_TYPE = rs("entry_type")
	v_vessel_code = rs("vessel_code")
	v_vessel_name = rs("vessel_name")
	V_SHORT_NAME = rs("vessel_short_name")
	MOC_NAME = rs("short_name")
	INSP_DATE = rs("inspection_date_disp")
	INSP_PORT = rs("inspection_port")

    'MOC Incentive: CC list
    '----------------------
	if v_vessel_code <> "" then
	    SQL = " SELECT TECH_MANAGER FROM VESSELS WHERE VESSEL_CODE = '" & v_vessel_code & "'"
	    set rsFleetName = connObj.execute(SQL)
	    if not rsFleetName.eof then
	        FLEET_NAME = rsFleetName(0)
	        MOC_CCList = "Inspection Team; " & FLEET_NAME & " Superintendent; Crew_Accounts_Group;Anil;Rajdeep Singh;"
	    end if
    end if
	
	
	strSqlVessMail = "select request_id, moc_id, moc_short_name, moc_full_name, inspection_date, "
	strSqlVessMail = strSqlVessMail & "section,length(deficiency),deficiency, status, inspection_port, vessel_code, vessel_name, "
	strSqlVessMail = strSqlVessMail & "vessel_short_name, fleet_code, "
	strSqlVessMail = strSqlVessMail & "to_char(inspection_date, 'DD-Mon-YYYY') inspection_date_disp "
	strSqlVessMail = strSqlVessMail & "from moc_vwr_list_of_observations "
	strSqlVessMail = strSqlVessMail & "where request_id = " & v_ins_request_id & " "
	strSqlVessMail = strSqlVessMail & "order by sort_order asc"	
	Set rsObjVessMail = connObj.Execute(strSqlVessMail)
	
	If rsObjVessMail.EOF = False Then
		VessMailBody = "Dear Capt. <Surname>," & vbcrlf & vbcrlf
		VessMailBody = VessMailBody & "Further to Vetting Inspection by " & rsObjVessMail("moc_short_name") & " at Port " & rsObjVessMail("inspection_port") & " on " & rsObjVessMail("inspection_date_disp") & ","
		VessMailBody = VessMailBody & " please find below deficiencies noted by the inspector in the final report:-" & vbcrlf & vbcrlf & vbCrLf
	
		v_send_to = rsObjVessMail("vessel_name")		
		srNo = 1		
		While Not rsObjVessMail.EOF
			sTemp = vbTab & rsObjVessMail("deficiency")
			sTemp = replace(sTemp,vbCrLf,vbCrLf & vbTab)
		 	VessMailBody = VessMailBody & srNo & "." & sTemp &  vbcrlf & vbcrlf	
			srNo = srNo + 1
			rsObjVessMail.MoveNext
		Wend
		VessMailBody = VessMailBody & "In order to correctly document the disposition of each item listed above, it is important that ALL noted "		
		VessMailBody = VessMailBody & "deficiencies are recorded in the ""Worklist Program""." & vbcrlf & vbcrlf
		VessMailBody = VessMailBody & "Plese review the above items and include them in the worklist under assignor ""MOC"" and advise your corrective actions." & vbcrlf & vbcrlf
		VessMailBody = VessMailBody & "Please direct all replies to ""Inspection Team"" group email address." & vbcrlf & vbcrlf
		
		VessMailBody = VessMailBody & "Brgds / XXX"	& vbcrlf
		VessMailBody = VessMailBody & "MSQV Dept."

        subject =rs("vessel_name")& " - " &rs("short_name") & " Inspection at " & rs("inspection_port") & " on " & rs("inspection_date_disp")
	Else	'no deficiencies exist

	End If

	rsObjVessMail.filter=0
%>
<html>
<head>
    <title>List of Observations</title>
    <meta http-equiv="expires" content="Tue, 20 Aug 2000 14:25:27 GMT">
    <link rel="stylesheet" href="moc.css">
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
    
    <script id="GetAsyncResponse" src="GetAsyncResponse.js" type="text/javascript"></script>

    <script type="text/javascript">
        function getContent(fn)
        {
            var ar;
            var crews="";
            var i;
            try{        
                var d = new Date();
                var ms = d.valueOf();    
                
                var qs = "noCache=" + ms
                var url 

                if (fn=="") fn="LOAD"
                
                qs += "&func=" + fn            
                qs += "&req_id=<%=v_ins_request_id%>"
                    
                if (fn == "SAVECREW_APPROVE"){
                    
                    ar = document.getElementsByName("chk")
                    for (i=0;i<ar.length;i++){
                        if(ar(i).checked==true)crews += ar(i).id + "~"
                    }
                    if(crews.length>0) crews=crews.substr(0,crews.length-1);
                
                    qs += "&crews=" + crews
                    
                    url = "approval.asp?" + qs; 
                    
                    var c = new  AsyncResponse(url,asyncCallBackApproved,0)
                    c.getResponse();        
                    return;
                }
                
                if (fn == "DISAPPROVE"){
                    url = "approval.asp?" + qs; 
                    var c = new  AsyncResponse(url,asyncCallBackDisapproved,0)
                    c.getResponse();        
                    return;
                }
                
                url = "approval.asp?" + qs; 
                
                var c = new  AsyncResponse(url,asyncCallBack,0)
                c.getResponse();        
           }
           catch(ex)
           {
           //nothing
           }
        }
        function asyncCallBackApproved(retval)
        {
            var mailBody = processCallBack(retval);
            SendMail2("<%=V_SHORT_NAME%> e","<%=MOC_CCList%>", "MOC Incentive Approved",mailBody)
        }
        function asyncCallBackDisapproved(retval)
        {
            var mailBody = processCallBack(retval);
            SendMail2("<%=V_SHORT_NAME%> e","<%=MOC_CCList%>", "MOC Incentive Disapproved",mailBody)
        }
        function asyncCallBack(retval)
        {
            processCallBack(retval)
        }

        function processCallBack(retval){
           var arText
           var arMail
           
           try{
                var textXML;
                if(typeof(retval)=="object")
                    textXML = retval.xml;
                else
                    textXML = retval;

                if (textXML == "") return; 
                
                arText = textXML.split("^^MAIL^^")
                
                document.getElementById("ApprovalMain").style.display='block';           
                document.getElementById("Approval").innerHTML=arText[0]
                
                return arText[1];
            }
           catch(ex)
           {
           //nothing
           }            
        
        }
                
        function finish_click(qid,vcode){
            try{
                var res            
                res = confirm("You have opted to finish adding observations.\n\n Proceed ?")
                if(res == false) return;
                    
                getContent("FINISH")
           }
           catch(ex)
           {
           //nothing
           }  
        }

        function Approve_click(approve,stage){
            try{
                var res            
                
                if (approve == 1){
                    
                    if (stage ==0){
                        res = confirm("You have opted to APPROVE the incentive.\n\nProceed?")
                        if(res == false) return;
                        
                        getContent("CREWLIST");
                    }
                    else if (stage ==1){
                        getContent("SAVECREW_APPROVE");
                    }
                }
                else if (approve == 0){
                    res = confirm("You have opted to DISAPPROVE the incentive. \n\n Proceed?")
                    if(res == false) return;
                    getContent("DISAPPROVE")
                }
           }
           catch(ex)
           {
           //nothing
           }         
        }    

        function showRules(){
            var dv
            dv = document.getElementById("rules")
            if (dv.style.display == "none") dv.style.display = ''
            else dv.style.display = 'none'
        }
        
    </script>
    
    <script language="Javascript" src="js_date.js"></script>

    <script language="VBScript" src="vb_date.vs"></script>

    <script language="vbscript" runat="Server">
        function b13(field_with_chr_13)
	        if isnull(field_with_chr_13) then
		        b13=field_with_chr_13
		        else
		        b13=replace(replace(field_with_chr_13,chr(13),"<br>"),"  ","&nbsp;&nbsp;")
	        end if
        end function
    </script>

    <script id="clientEventHandlersJS" language="javascript">
        <!--
        function cClose()
  		        {
  			        //var name= confirm("Are you sure? ")
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

        function fncall(v_ins_request_id,v_def_id,vessel_name,moc_name)
		        {
			        winStats='toolbar=no,location=no,directories=no,menubar=no,'
			        winStats+='scrollbars=yes,resizable=yes'
			        if (navigator.appName.indexOf("Microsoft")>=0) {
				        winStats+=',left=120,top=10,width=600,height=640'
			        }else{
				        winStats+=',screenX=350,screenY=200,width=400,height=280'
			        }
			        adWindow=window.open("ins_request_def_entry.asp?v_ins_request_id="+v_ins_request_id+"&v_def_id="+v_def_id+"&v_vessel_code=<%=v_vessel_code%>&v_vessel_name="+vessel_name+"&moc_name="+moc_name,"moc_request_def_rem_entry",winStats);     
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
		        document.form1.action="ins_request_def_maint.asp?item="+v_sort_field+"&order="+v_sort_order+"&v_ins_request_id="+v_ins_request_id+"&vessel_name="+vessel_name+"&moc_name="+moc_name
		        //alert(document.form1.action)
		        document.form1.submit();
		        }

        function substStar(inVal)
        {
	        if (inVal == "") return "*"; else return inVal;
        }

        function callLOBReport()
        {
	        var params;

	        params = "repcallnew20.asp?rr1=moc_list_of_obs.rpt";
	        params += "&rp1=" + substStar("<% =v_ins_request_id %>");
	        params += "&rp2=" + substStar("");
	        params += "&rp3=" + substStar("");
	        params += "&rp4=" + substStar("");
	        params += "&rp5=" + substStar("");
	        params += "&rp6=" + substStar("");
	        params += "&rp7=" + substStar("");
	        params += "&rp8=" + substStar("");
	        params += "&rp9=" + substStar("");
	        params += "&rp10=" + substStar("");
	        params += "&rp11=" + substStar("");
	        params += "&rp12=" + substStar("");
	        params += "&rp13=" + substStar("");
	        params += "&rp14=" + substStar("");
	        params += "&rp15=" + substStar("");
	        params += "&rp16=" + substStar("");
	        params += "&rp17=" + substStar("");
	        params += "&rp18=" + substStar("");
	        params += "&rp19=" + substStar("");
	        params += "&rp20=" + substStar("");
	        params += "&rs1="
	        params += "&rs2="
	        params += "&rs3="
	        params += "&rs4="
	        params += "&rs5="
	        params += "&rs6="
	        params += "&rs7="

	        winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	        winStats += 'scrollbars=yes,resizable=yes'

	        if (navigator.appName.indexOf("Microsoft") >= 0) 
	        {
		        winStats += ',left=0,top=0,width=' + (screen.width - 10) + ',height=' + (screen.height - 30)
	        }
	        else
	        {
		        winStats += ',screenX=350,screenY=200,width=350,height=180'
	        }

	        windName = String(parseInt(String(Math.random() * 1000)));

	        repWind = window.open(params, windName, winStats);
	        repWind.focus();

	        return false;
        }

        function callReplyReport(repName)
        {
	        var params;
	        var Fax_Email = GetFaxEmail()
        	
	        params = "repcallnew20.asp?rr1=" + repName;
	        params += "&rp1=" + substStar("<% =v_ins_request_id %>");
	        params += "&rp2=" + substStar(Fax_Email);
	        params += "&rp3=" + substStar("");
	        params += "&rp4=" + substStar("");
	        params += "&rp5=" + substStar("");
	        params += "&rp6=" + substStar("");
	        params += "&rp7=" + substStar("");
	        params += "&rp8=" + substStar("");
	        params += "&rp9=" + substStar("");
	        params += "&rp10=" + substStar("");
	        params += "&rp11=" + substStar("");
	        params += "&rp12=" + substStar("");
	        params += "&rp13=" + substStar("");
	        params += "&rp14=" + substStar("");
	        params += "&rp15=" + substStar("");
	        params += "&rp16=" + substStar("");
	        params += "&rp17=" + substStar("");
	        params += "&rp18=" + substStar("");
	        params += "&rp19=" + substStar("");
	        params += "&rp20=" + substStar("");
	        params += "&rs1="
	        params += "&rs2="
	        params += "&rs3="
	        params += "&rs4="
	        params += "&rs5="
	        params += "&rs6="
	        params += "&rs7="

	        winStats = 'toolbar=no,location=no,directories=no,menubar=no,'
	        winStats += 'scrollbars=yes,resizable=yes'

	        if (navigator.appName.indexOf("Microsoft") >= 0) 
	        {
		        winStats += ',left=0,top=0,width=' + (screen.width - 10) + ',height=' + (screen.height - 30)
	        }
	        else
	        {
		        winStats += ',screenX=350,screenY=200,width=350,height=180'
	        }

	        windName = String(parseInt(String(Math.random() * 1000)));

	        repWind = window.open(params, windName, winStats);
	        repWind.focus();

	        return false;
        }

        function inspection_grade_onpropertychange() {
	        var obj = form1.inspection_grade
	        switch(obj.value)
	        {
		        case "":
		        case "5":
		        case "4":
		        case "3":obj.style.backgroundColor = "";obj.style.color = "black";break;
		        case "2":
		        case "1":obj.style.backgroundColor = "red";obj.style.color = "yellow";break;
	        }
        }
<%if v_button_disabled="" then%>
    function SaveGrade()
    {
	    var obj = form1.inspection_grade
	    if(obj.value=="")
		    return;
	    window.open("SaveData.asp?ID=GRADE&KEYFIELD1=<%=v_ins_request_id%>&VALUE1=" + obj.value,"savegrade","location=no,toolbars=no,menubar=no,left=2000")
    }
<%end if%>
        function window_onload() {
	        inspection_grade_onpropertychange()
        }
        function setTooltip(obj)
        {
	        obj.title = obj.children(0).value
        }
        function openJoblistPopup(jobcode)
        {
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
	        adWindow = window.open("http://webserve2/wls/joblist_popup_frame.asp?v_job_code=" + jobcode + "&v_vessel_code=<%=v_vessel_code%>&v_opener_name=moc", "JobDetails", winStats);
	        adWindow.focus();
	        return false;
        }
//-->
    </script>

    <script language="vbscript">
        function GetJobCode(obj)
	        dim selText
	        dim sURL,win,NewJobCode,defid,jobcode
	        defid = obj.GetAttribute("defid")
	        jobcode = obj.GetAttribute("jobcode")

			'NCR_list.asp returns a job_code
	        sURL = "NCR_list.asp?v_vessel_code=<%=v_vessel_code%>&v_vessel_name=<%=v_vessel_name%>&v_def_id=" & defid & "&v_job_code=" & jobcode
	        selText = obj.parentElement.parentElement.cells(1).innerText
	        NewJobCode = window.showModalDialog(sURL,selText,"dialogHeight:600px;dialogWidth:800px;resizable:yes;center:yes;status:no;scroll:yes;")

	        if NewJobCode<>jobcode then
		        form1.action=""
		        form1.submit
	        end if
        end function
        function GetFaxEmail()
	        dim response
	        response = MsgBox("Do you want to send this report by email?", vbYesNo+vbInformation, "MOC Report")
	        if response = vbYes then
		        GetFaxEmail = "Email"
	        else
		        GetFaxEmail = "Fax"
	        end if
        end function

        dim objTip,mx,my
        Sub ShowTooltip(obj)
	        set objTip = obj
	        mx = window.event.clientX
	        my = window.event.clientY
	        if divTooltip.style.display="" then
		        DisplayTooltip
	        else
		        setTimeout "DisplayTooltip",400,"vbscript"
	        end if
        End Sub

        Sub HideTooltip
	        divTooltip.style.display="none"
	        divTooltip.innerHTML=""
	        objTip = empty
        End Sub
        sub DisplayTooltip
	        dim x,y

	        if IsEmpty(objTip) then exit sub
	        x=mx + 20
	        y=my + document.body.scrollTop
	        'y=40 + objPortCallTip.offsetTop + objPortCallTip.parentElement.offsetTop
        	

	        with divTooltip
		        .innerHTML = objTip.children(0).innerHTML
		        .style.display=""
        				
		        if x+.offsetWidth+20 > document.body.offsetWidth then x=x-.offsetWidth
		        if y+.offsetHeight+5 > document.body.offsetHeight + document.body.scrollTop then y=y-.offsetHeight
		        .style.left = x
		        .style.top = y
	        end with
        end sub
        sub SendMail()
	        if frm1.mailbody.value<>"" then   
		        Set objOutlook = CreateObject("Outlook.Application") 
		        Set objMail = objOutlook.createitem(olMailItem) 
		        objMail.To = "<Vessel>"	
		        objMail.cc="Inspection Team;<Ops Supdt>"	
		        objMail.subject =frm1.mailsubject.value
		        strBody = frm1.mailbody.value	
		        objMail.Body =  strBody  
		        objMail.display
		        Set objMail = Nothing 
		        Set objOutlook = Nothing 
	        else
	          msgbox "No deficiencies to send",vbinformation,"MOC"	
	        end if	
        end sub
    </script>

    <script type="text/vbscript" language="vbscript">
        sub SendMail2(sTo,sCC, sSub,msgBody)

            dim objOutlk
            dim objMail
            dim SafeMail
            dim lPositionOffset
            dim strBody
            
            Set objOutlk = CreateObject("Outlook.Application") 
            Set objMail = objOutlk.CreateItem(olMailItem)
            Set SafeMail = CreateObject("redemption.SafeMailitem")

            SafeMail.item = ObjMail

            with SafeMail
                .Recipients.Add sTo
                
                .cc = sCC
                .Recipients.ResolveAll 
                .subject = sSub
                .body = msgBody

                '.Save
                .display
            end with
            
            'Clean up
            Set SafeMail = Nothing 
            Set objOutlk = Nothing 
        end sub
    
    
    </script>
</head>
<body class="bgcolorlogin" language="javascript" onload="return window_onload()">
    <object id="mail" style="left: 0px; top: 0px" name="mail" codebase="../../MailClient.CAB"
        classid="CLSID:115D7155-2186-4AEC-A57E-A1777087AE01" width="0" height="0" viewastext>
        <param name="_ExtentX" value="26">
        <param name="_ExtentY" value="26">
    </object>
    <center>
        <br>
        <h4>
            List of Observations</h4>
        
        <table width="98%" border="0">
            <tr>
                <td>
                    <h4>
                        Vessel:</h4>
                </td>
                <td>
                    <h4>
                        <%=v_vessel_name%>
                    </h4>
                </td>
                <td rowspan="6" valign="top" align="right">
                    <div id="ApprovalMain" style="width: 300px; border: 2px solid outset; text-align: center;
                        background-color: DarkBlue; color: White; float: right;">
                        <div style="float:left;"><a href=# onmouseover='showRules()' onmouseout='showRules()'><font color=yellow><u>Rules</u></font></a></div>
                        <div>MOC Inspection</div>
                        
                        <div id="Approval" style="background-color: LightBlue; color: Black; text-align: center;
                            padding: 4px;">
                           
                        </div>
                    </div>
                    <div id="rules" style="text-align: left; display: none; margin-top: 3px;">
                        <div>
                            <div style="background-color: #cccccc; color: Blue; text-align: center;">
                                Rules
                            </div>
                            <div style="background-color: #d8d8d8; border: 1px solid outset; font-size: 10px;">
                                <ul style="font-size: 9px;margin-top:0px;margin-bottom:0px;">
                                    <li>Age upto 10 years, Observations &lt;= 4</li>
                                    <li>Age 10 to 15 years, Observations &lt;= 5</li>
                                    <li>Age 15+ years, Observations &lt;= 6</li>
                                    <li>No high risk deficiencies</li>
                                    <li>Incentive: USD 800 + USD 30 increment<br /> for each observation less than 
                                    the maximum allowed observations.</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                
                    <%="<script type=text/javascript>getContent('')</script>" %>
                </td>
            </tr>
            <tr>
                <td>
                    <h4>
                        MOC:</h4>
                </td>
                <td>
                    <h4>
                        <%=MOC_NAME%>
                    </h4>
                </td>
            </tr>
            <tr>
                <td>
                    <h4>
                        Insp. Date:</h4>
                </td>
                <td>
                    <h4>
                        <%=INSP_DATE%>
                    </h4>
                </td>
            </tr>
            <tr>
                <td>
                    <h4>
                        Insp. Port:</h4>
                </td>
                <td>
                    <h4>
                        <%=INSP_PORT%>
                    </h4>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <%  v_mess=Request.QueryString("v_message")
			if v_mess <> "" then
                    %>
                    <font color="red" size="+2">
                        <%=v_mess%>
                    </font>
                    <% end if%>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td align="center" colspan="3">
                    <a href='http://webserve2/vid/create_data_file.asp?vessel_code=<%=v_vessel_code%>&questionnaire_id=10000380'
                        target='vessel_particulars'><span style='text-align: center; font-size: 9; font-weight: bold;
                            color: blue'>Vessel Particulars</span></a> <a href="javascript:window.print()">
                                <img src="Images/print.gif" border="0" alt="Print this Page" width="22" height="20"></a>
                </td>
            </tr>
        </table>
    </center>
    <form name="frm1" method="post">
        <textarea name="mailsubject" id="mailsubject" style="display: none"><%=subject%></textarea>
        <textarea name="mailbody" id="mailbody" style="display: none"><%=VessMailBody%></textarea>
        <table>
            <tr valign="bottom">
                <td colspan="3" align="center">
                    <input type="button" name="v_send_to_vessel" style="width: 95pt" value="Send Observations<%=Chr(13)%>to Vessel"
                        onclick="SendMail()">&nbsp;
                    <!--<input TYPE="button" NAME="v_send_to_vessel" STYLE="width:95pt" VALUE="Send SIRE Obs.<%=Chr(13)%>to Vessel" OnClick="javascript:fun('','','<%=VessMailBody_SIRE%>');">&nbsp;-->
                    <%'if ENTRY_TYPE="MOC" then%>
                    <input type="button" name="v_reply_to_moc" style="width: 95pt" value="Reply to<%=Chr(13)%>SIRE Report"
                        onclick="return callReplyReport('reply_to_sire.rpt');">&nbsp;
                    <%'else%>
                    <input type="button" name="v_reply_to_sire" style="width: 95pt" value="Reply to<%=Chr(13)%>Report"
                        onclick="return callReplyReport('reply_to_moc.rpt');">&nbsp;
                    <%'end if%>
                    <input type="button" name="v_list_of_obs" style="width: 95pt" value="List of<%=Chr(13)%>Observations"
                        onclick="return callLOBReport();">
                </td>
            </tr>
        </table>
    </form>
    <%
        v_filter =""
        strSql = "Select deficiency_id, request_id, section,deficiency, reply,status, risk_factor,"
        strSql = strSql & "md.sort_order, to_char(create_date,'DD-MON-YYYY') create_date1, created_by, "
        strSql = strSql & "LAST_MODIFIED_DATE,LAST_MODIFIED_by,"
        strSql = strSql & " vq.question_text,md.wls_job_code"
        strSql = strSql & " from MOC_deficiencies md, moc_viq_questions vq"
        strSql = strSql & " where md.section = vq.question_number(+)"
        strSql = strSql & " and request_id=" & v_ins_request_id

        if request("item")<>"" then
	        strSql=strSql & " order by "& request("item") & " " & request("order") 
        else
	        strSql=strSql & " order by sort_order "	
        end if

        Set rsObj=connObj.execute(strSql)
    %>
    <!--h3 align="center">Selection Filter</h3-->
    <form name="form1" action="ins_request_def_delete.asp" method="post">
        <input type="hidden" id="v_ins_request_id" name="v_ins_request_id" value="<%=Request("v_ins_request_id")%>">
        <input type="hidden" id="vessel_name" name="vessel_name" value="<%=Request("vessel_name")%>">
        <input type="hidden" id="moc_name" name="moc_name" value="<%=Request("moc_name")%>">
        <table width="100%">
            <tr>
                <td align="left">
                    <strong>Inspection Grade</strong>
                    <select style="font-weight: bold" name="inspection_grade" language="javascript" onpropertychange="return inspection_grade_onpropertychange()">
                        <option value="<%%>">Not graded</option>
                        <option value="1" <%if rs("inspection_grade")="1" then response.write " selected"%>>
                            1 - Unacceptable</option>
                        <option value="2" <%if rs("inspection_grade")="2" then response.write " selected"%>>
                            2 - Poor</option>
                        <option value="3" <%if rs("inspection_grade")="3" then response.write " selected"%>>
                            3 - Average</option>
                        <option value="4" <%if rs("inspection_grade")="4" then response.write " selected"%>>
                            4 - Good</option>
                        <option value="5" <%if rs("inspection_grade")="5" then response.write " selected"%>>
                            5 - Excellent</option>
                    </select>
                    &nbsp;
                    <button <%=v_button_disabled%> onclick="javascript:SaveGrade()">
                        Save</button>
                <td align="right">
                    <strong>Date/Time:</strong>
                    <% Response.write day(Now()) & "-" & monthname(month(now()),true) & "-"&year(now()) & " " & hour(now()) & ":" & minute(now()) & ":" & second(now())%>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td width="10%" class="tableheader" align="center">
                    Section<br>
                    <br>
                    <table width="100%" align="left">
                        <tr>
                            <td align="left">
                                <a href="javascript:v_sort('section','asc');">
                                    <img src="Images/up.gif" alt="Sort Ascending by section" border="0" width="15" align="left"
                                        hspace="0">
                                </a>
                            </td>
                            <td align="right">
                                <a href="javascript:v_sort('section','desc');">
                                    <img src="Images/down.gif" alt="Sort Descending by section" border="0" width="15"
                                        align="right" hspace="0">
                                </a>
                            </td>
                        </tr>
                    </table>
                </td>
                <td width="40%" class="tableheader" align="center">
                    Observation<br>
                    <br>
                    <br>
                </td>
                <td width="40%" class="tableheader" align="center">
                    Reply<br>
                    <br>
                    <br>
                </td>
                <td width="10%" class="tableheader" align="center">
                    Status<br>
                    <br>
                    <table width="100%" align="left">
                        <tr>
                            <td align="left">
                                <a href="javascript:v_sort('status','asc');">
                                    <img src="Images/up.gif" alt="Sort Ascending by status" border="0" width="15" align="left"
                                        hspace="0">
                                </a>
                            </td>
                            <td align="right">
                                <a href="javascript:v_sort('status','desc');">
                                    <img src="Images/down.gif" alt="Sort Descending by status" border="0" width="15"
                                        align="right" hspace="0">
                                </a>
                            </td>
                        </tr>
                    </table>
                </td>
                <td class="tableheader" align="center">
                    Sort<br>
                    Order
                    <table width="100%" align="left">
                        <tr>
                            <td align="left">
                                <a href="javascript:v_sort('sort_order','asc');">
                                    <img src="Images/up.gif" alt="Sort Ascending by sort order" border="0" width="15"
                                        align="left" hspace="0">
                                </a>
                            </td>
                            <td align="right">
                                <a href="javascript:v_sort('sort_order','desc');">
                                    <img src="Images/down.gif" alt="Sort Descending by sort order" border="0" width="15"
                                        align="right" hspace="0">
                                </a>
                            </td>
                        </tr>
                    </table>
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
	select case rsObj("risk_factor")
		case "GENERAL":sRisk="<span style='color:black;font-weight:bold'>Risk: GENERAL</span>"
		case "LOW":sRisk="<span style='color:green;font-weight:bold'>Risk: LOW</span>"
		case "HIGH":sRisk="<span style='color:red;font-weight:bold'>Risk: HIGH</span>"
	end select
            %>
            <tr>
                <td valign="top" class="<%=class_color%>" onmousemove="ShowTooltip(this)" onmouseover="ShowTooltip(this)"
                    onmouseout="HideTooltip()">
                    <textarea id="qtext" style="display: none"><%=rsObj("question_text")%></textarea>
                    <%=rsObj("section") %>
                    &nbsp;
                    <%if DISPLAY_LINK_TO_WORKLIST then%>
                    <%if rsObj("wls_job_code")<>"" then%>
                    <br>
                    <br>
                    <a class="link1" href="#" onclick="javascript:return openJoblistPopup(<%=rsObj("wls_job_code")%>);"
                        onmouseover="javascript:window.event.cancelBubble=true" onmousemove="javascript:window.event.cancelBubble=true">
                        View job in Worklist</a>
                    <%end if%>
                    <br>
                    <br>
                    <a class="link2" href="#" defid="<%=rsObj("deficiency_id")%>" jobcode="<%=rsObj("wls_job_code")%>"
                        onclick="javascript:GetJobCode(this);return false;" onmouseover="javascript:window.event.cancelBubble=true"
                        onmousemove="javascript:window.event.cancelBubble=true">Create link to Worklist</a>
                    <%end if%>
                </td>
                <td valign="top" class="<%=class_color%>">
                    <%=sRisk%>
                    <br>
                    <a name="1" title="Click to Edit" href="javascript:fncall('<%=rsObj("request_id") %>','<%=rsObj("deficiency_id") %>','<%=v_vessel_name%>','<%=replace(MOC_NAME,"'","\'")%>');">
                        <%=b13(rsObj("deficiency")) %>
                        &nbsp; </a>
                </td>
                <td valign="top" class="<%=class_color%>">
                    <%=b13(rsObj("reply")) %>
                    &nbsp;
                </td>
                <td valign="top" class="<%=class_color%>">
                    <%=rsObj("status") %>
                    &nbsp;</td>
                <td valign="top" class="<%=class_color%>">
                    <%=rsObj("sort_order") %>
                    &nbsp;</td>
            </tr>
            <%
rsObj.movenext
wend
            %>
            <%
else
Response.Write "<tr><td colspan=7 class=tabledata align=center><STRONG>No Data Found!!</STRONG> </td></tr>"
end if ' if not (rsObj.bof or rsObj.eof) then
            %>
            <tr>
                <td colspan="5" class="tabledata" align="center">
                    <!--INPUT type="submit" value="Delete the Selected" id=submit1 name="submit1"-->
                    &nbsp;
                    <input type="submit" value="Close" id="submit4" name="submit4" onclick="javascript:return cClose();">
                </td>
                <tr>
        </table>
    </form>
    <div id="divTooltip" style="display: none; position: absolute; background-color: lightblue;
        font-size: 10px; font-weight: bold; height: 50px; width: 250px; overflow: visible;
        text-align: left; padding: 5px; border: 1px solid midnightblue">
    </div>
</body>
</html>
