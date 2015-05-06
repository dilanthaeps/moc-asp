<%option explicit%>
<!--#include file="common_dbconn.asp"-->
<!--#include file="ado.inc"-->
<%
    dim  SQL, RS, Age, Max, Count, output, HighRisks ,crews, arCrews   
    Dim UserName, status, approve, AmtCalculated, func
    Dim msgbody
    
    DIM INCENTIVE_AMOUNT, INCREMENT
    DIM REC_STATUS,REC_STAGE
    dim REQ_ID, VCODE, ENTRY_TYPE,VNAME,MOC_NAME,INSP_DATE_DISP,INSP_DATE,INSP_PORT
    
    
    '******************************************
    'CAN BE CHANGED AS PER THE COMPANY POLICY
    INCENTIVE_AMOUNT = 800
    INCREMENT = 30
    '******************************************

    func = request.QueryString("func")
    status = request.QueryString("status")
    approve = request.QueryString("approve")
    crews = request.QueryString("crews")
    
    REQ_ID = Request("req_id")
    
    set RS=Server.CreateObject("ADODB.Recordset")
        
    '******************************************
    'This part of the code is to check the section 
    'if it is a high risk one or not
    
    if func = "CHECK_HIGHRISK" then
        dim sRet
        dim sec_id
        sec_id = request.QueryString("sec")
        if sec_id <> "" then
            sRet = "NO"
            SQL = " SELECT * from MOC_HIGH_RISK_DEFICIENCIES "
            SQL = SQL &  "  WHERE section = '"  & sec_id & "'"

            RS.Open  SQL,connObj
            
            if not RS.EOF then
                sRet = "YES"
            end if
        end if
        response.Write sRet
        response.End
    end if    
    '******************************************

    
    SQL = "Select ir.vessel_code,v.vessel_name,entry_type,mm.short_name,inspection_port,inspection_grade,"
	SQL = SQL & " to_char(inspection_date,'DD-Mon-YYYY')inspection_date_disp,inspection_date"
	SQL = SQL & " from moc_master mm,moc_inspection_requests ir, vessels v"
	SQL = SQL & " where ir.vessel_code=v.vessel_code and mm.moc_id=ir.moc_id"
	SQL = SQL & " and ir.request_id=" & REQ_ID
	
	
	set rs = connObj.execute(SQL)
	if not rs.EOF then
	    ENTRY_TYPE = rs("entry_type")
	    VCODE = rs("vessel_code")
	    VNAME = rs("vessel_name")
	    MOC_NAME = rs("short_name")
	    INSP_DATE_DISP = rs("inspection_date_disp")
	    INSP_DATE = rs("inspection_date")
	    INSP_PORT = rs("inspection_port")
        RS.Close
    end if

    'REC_STAGE =1 -> FINISHED
    'REC_STAGE =2 -> VERIFY CREW
    'REC_STAGE =3 -> APPROVED/DISAPPROVED
       
    
    UserName = Request.ServerVariables("LOGON_USER")
    if trim(UserName)<>"" then
        UserName = right(UserName, len(UserName) - instr(1,UserName,"\"))
        UserName = ucase(UserName)
    end if
    
    if crews <> "" then    
        arCrews = split(crews,"~")
    end if
    
    with RS
        .CursorLocation=adUseClient
        .CursorType=adOpenStatic
        .LockType=adLockReadonly
    end with
    
    if UserIsAdmin or UserIsSuper then
        select case func
            case "LOAD"

            case "FINISH"
                SQL = "SELECT REQUEST_ID FROM MOC_APPROVALS WHERE REQUEST_ID=" & req_id
                RS.Open  SQL,connObj
                
                IF RS.EOF THEN
	                SQL = "INSERT INTO MOC_APPROVALS(REQUEST_ID,STAGE,APPROVER,APPROVAL_DT,STATUS) VALUES(" & req_id & ",1,'" & UserName & "',SYSDATE,'PENDING APPROVAL')" 
                else
	                SQL = "UPDATE MOC_APPROVALS SET STAGE=1, APPROVER='" & UserName & "', APPROVAL_DT=SYSDATE, STATUS='PENDING APPROVAL' WHERE REQUEST_ID=" & req_id
                end if       
                connObj.execute(SQL)
                RS.Close

            case "CREWLIST"
                SQL = "UPDATE MOC_APPROVALS SET STAGE=2, APPROVER='" & UserName & "', APPROVAL_DT=SYSDATE,STATUS='VERIFY CREW' WHERE REQUEST_ID=" & req_id
                connObj.execute(SQL)                

            case "SAVECREW_APPROVE"
                'Save Crew
                
                dim iCtr 
                ON ERROR RESUME NEXT
                for ictr = 0 to ubound(arCrews)
                    SQL = "INSERT INTO MOC_APPROVED_CREWS VALUES(" & req_id & "," & arCrews(ICTR) & ")"    
                    connObj.execute(SQL)
                next
                ON ERROR GOTO 0
                
                SQL = "UPDATE MOC_APPROVALS SET STAGE=4, APPROVER='" & UserName & "', APPROVAL_DT=SYSDATE,STATUS='APPROVED' WHERE REQUEST_ID=" & req_id
                connObj.execute(SQL)                
            
                msgbody =  "^^MAIL^^" & getApproveMsg                
                msgbody = updateMsgForCrew(msgbody,INSP_DATE_DISP)
                
            case "DISAPPROVE"
                SQL = "UPDATE MOC_APPROVALS SET STAGE=4, APPROVER='" & UserName & "', APPROVAL_DT=SYSDATE,STATUS='DISAPPROVED' WHERE REQUEST_ID=" & req_id            
                connObj.execute(SQL)                

                msgbody = "^^MAIL^^" & getRejectMsg
                msgbody = updateMsgForCrew(msgbody,INSP_DATE_DISP)

        end select

        response.Write getContent
        response.End

    else
        response.Write "You are not authorised to view this section"
        response.End
    end if

'--------------------------Helper Functions --------------------------------    
    function getContent()
        Dim retStr
        
        
        'Check approvals for status>=1 
        'If not found then show the buttons "Add Record", "Finish"
        '---------------------------------------------------------
        set RS = connObj.execute("SELECT STAGE, STATUS FROM MOC_APPROVALS WHERE STAGE >=1 AND REQUEST_ID = " & req_id )
        
        if  RS.eof then
            retStr = retStr & "<br><input type=button name='2' onclick=""javascript:finish_click()"" value='Finish adding Observations'><br>"
            retStr = retStr & "<br><input type=button name='1' onclick=""javascript:fncall('" & req_id & "','0','" & vname & "','" & replace(MOC_NAME,"'","\'") & "')"" value='Add New Observation'><br>"
            getContent = retStr
            exit function
        else
            REC_STAGE = cint(RS(0))
            REC_STATUS = RS(1)
        end If
        RS.close
        

        'Check for Observations 
        'If not found then show the buttons "Add Record"
        '-----------------------------------------------
        Count = getCount
        
        if Count = 0 and REC_STATUS="" then
            retStr = "No Observations found"
            retStr = retStr & "<br><input type=button name='2' onclick=""javascript:finish_click()"" value='Finish adding Observations'><br>"
            retStr = retStr & "<input type=button name='1' onclick=""javascript:fncall('" & req_id & "','0','" & vname & "','" & replace(MOC_NAME,"'","\'") & "')"" value='Add New Observation'>"
            Response.Write retStr
            response.end
        end if
        
        
        'The approval table is having record with status >=1 for the req_id
        'It may be STAGE =1, STAGE= 3, STAGE= 4
        '------------------------------------------------------------------
        
        'Create panel content based on the stage
        '   -   Calculate age of the ship
        '---------------------------------------        
        if  REC_STAGE <> 2 then
            Age = CalculateAge(VCODE)
            if Age >= 15 then
                Max = 6 ' Revised from 7 to 6 on 12mar08
                retStr = retStr &   "Age of vessel: 15+ Years<br>"
            elseif Age < 15 and Age >= 10 then
                Max = 5 ' Revised from 6 to 5 on 12mar08
                retStr = retStr &   "Age of vessel: 10 to 15 Years<br>"
            elseif Age < 10 then
                retStr = retStr &   "Age of vessel: Upto 10 Years<br>"
                Max = 4 ' Revised from 5 to 4 on 12mar08
            end if            
              
            retStr = retStr &  "No of Observations:" & Count & " , Maximum allowed:" & Max & "<br>"
            
            if Count >= 0 then            
                if Count > Max then
                    AmtCalculated = INCENTIVE_AMOUNT 
                    retStr = retStr &  "<div style='background-color:red;color:white;text-align:center;'>This vessel does not qualify for the incentive as the no of observations are more than the maximum allowed observations.</div>"
                else            
                    AmtCalculated = INCENTIVE_AMOUNT +  (INCREMENT * (Max-Count)) * 4
                    
                    HighRisks = getHighRiskDeficiencies
                    
                    if HighRisks <> "" then
                        retStr = retStr &  "High risk deficiencies: " & HighRisks 
	                    retStr = retStr &  "<div style='background-color:red;color:white;text-align:center;'>This vessel does not qualify for the incentive due to observations having HIGH RISK deficiency.</div>"
                    else
	                    retStr = retStr &  "Amount calculated(USD):" & AmtCalculated 
	                    retStr = retStr &  "<div style='background-color:green;color:white;text-align:center;height:40px;'>This vessel qualifies for the incentive.</div>"
	                    
                    end if
                end if
            else
                retStr = "No Observations found"
                retStr = retStr & "<br><input type=button name='2' onclick=""javascript:finish_click()"" value='Finish adding Observations'><br>"
                retStr = retStr & "<input type=button name='1' onclick=""javascript:fncall('" & req_id & "','0','" & vname & "','" & replace(MOC_NAME,"'","\'") & "')"" value='Add New Observation'>"
                Response.Write retStr
                response.end
            end if
            
	        
		end if
		    
        'Create controls based on the stage
        '----------------------------------        
        retStr = retStr &  "<div style='text-align:center'>"
		
        if  REC_STAGE = 1 then
            retStr = retStr &   "<input type=button value=Approve onclick='Approve_click(1,0)'>"
            retStr = retStr &   "<input type=button value=Disapprove onclick='Approve_click(0,0)'"
            retStr = retStr &   "</div>"        
            
        elseif  REC_STAGE = 2 then
            retStr = getCrewList(INSP_DATE_DISP)
        
            retStr = retStr &   "<input type=button value=""Save and Proceed"" onclick='Approve_click(1,1)'>"
            retStr = retStr &   "</div>"        
            
        elseif  REC_STAGE = 3 then
            'retStr = retStr &   "<input type=button value=Finish onclick='Approve_click(1,2)'>"
            retStr = retStr &   "</div>"        
        
        elseif  REC_STAGE = 4 then
            IF msgbody <> "" then
                SQL = "UPDATE MOC_APPROVALS SET INCENTIVE=" & INCENTIVE_AMOUNT & ",UNIT_INCREMENT=" & INCREMENT & ",TOTAL_AMOUNT=" & AmtCalculated & ",BONUS_POINTS=" & Max-Count & " WHERE REQUEST_ID=" & req_id
                connObj.execute(SQL)                
                msgbody = replace(msgbody,"TOTAL_AMT",AmtCalculated)
            end if
            retStr = retStr &   "<div style='border:1px solid outset;margin-top:5px;font-weight:bold;'>The incentive is " & REC_STATUS & "</div></div>" & msgbody
        end if                
        getContent = retStr
    end function

    function CalculateAge (VC)
        'SQL = " SELECT TRUNC ((SYSDATE - TO_DATE (answer_data, 'dd/mm/yyyy')) / 365) years"
        'SQL = SQL &  "   FROM vpd_master_answers"
        'SQL = SQL &  "  WHERE master_question_id = 10001371 AND vessel_code = '" & vcode & "'"

        SQL = "select NVL (moc_fn_vessel_age_years ('" & VC & "'), 0) age from dual"

        RS.Open  SQL,connObj

        if not RS.EOF then
            CalculateAge = CDbl(rs(0))
        else
            CalculateAge = 0
        end if
        RS.close
    end function
    
    function getCrewList(inspectionDate)
        DIM sCrews
        dim conDanaos
        Set conDanaos=server.CreateObject("ADODB.Connection")
	    conDanaos.Open "Provider=MSDAORA.1;User ID=wfadmin; Password=wfa0m1n;Data Source=AIX;"
        
        SQL = " SELECT   pv.persons_code, p.persons_name || ' ' || p.PERSONS_SURNAME FULLNAME,"
        SQL = SQL &  "  P.PERSONS_SURNAME, pv.joining_rank,to_char(pv.JOINING_DATE,'dd Mon yy') doj,"
        SQL = SQL &  "          DECODE (pv.joining_rank,"
        SQL = SQL &  "                  'MST', '1',"
        SQL = SQL &  "                  'C/O', '2',"
        SQL = SQL &  "                  'EIC', '3',"
        SQL = SQL &  "                  'C/E', '4',"
        SQL = SQL &  "                  '1/E', '5'"
        SQL = SQL &  "                 ) ID"
        SQL = SQL &  "     FROM tankpac.persons_voyages pv, tankpac.persons p"
        SQL = SQL &  "    WHERE pv.persons_code = p.persons_code"
        SQL = SQL &  "      AND pv.vessel_code = '" & vcode & "'"
        SQL = SQL &  "      AND (    pv.joining_date < TO_DATE ('" & inspectionDate & "', 'DD-Mon-YYYY')"
        SQL = SQL &  "      AND NVL (pv.sign_off_date, SYSDATE) >= TO_DATE ('" & inspectionDate & "', 'DD-Mon-YYYY')"
        SQL = SQL &  "      )"        
        SQL = SQL &  "      AND UPPER (pv.joining_rank) IN ('MST', 'C/O', 'C/E', 'EIC', '1/E')"
        SQL = SQL &  " ORDER BY ID, joining_rank"
        
        RS.Open  SQL,conDanaos
        sCrews = sCrews & "Crew list as of : " & INSP_DATE_DISP
        sCrews = sCrews & "<table class=crewTable cellspacing=0 cellpadding=0><tr class=crewHeader><td>Name<td>Rank&nbsp;&nbsp;<td>D.O.J</td></tr>"
        while not rs.EOF         
            sCrews = sCrews & "<tr class=crewData><td><input type=checkbox name=chk id=" & rs("persons_code") & " checked>" & rs("FULLNAME") & "<td>(" & rs("joining_rank") & ")<td>" & rs("DOJ") & "</td>"
            rs.MoveNext
        wend
        sCrews = sCrews & "</table>"
        getCrewList = sCrews    
        RS.Close
    end function

    function getCount()
        SQL = " SELECT count(*)"
        SQL = SQL &  "    FROM moc_deficiencies WHERE  request_id = "  & req_id & " AND viq_dont_count <> 1"
        RS.Open  SQL,connObj
        getCount =  CDbl(rs(0))
        RS.close
    end function
    
    function getHighRiskDeficiencies()
        dim strRet
        SQL = " SELECT md.section"
        SQL = SQL &  "   FROM moc_deficiencies md, moc_high_risk_deficiencies hrd"
        SQL = SQL &  "  WHERE md.section = hrd.section AND request_id = "  & req_id

        RS.Open  SQL,connObj
        
        if not RS.EOF then
            while not rs.EOF
                strRet = strRet & rs(0) & ","
                rs.MoveNext
            wend
        end if
        getHighRiskDeficiencies = strRet        
    end function
   
    function updateMsgForCrew(msg,inspectionDate)
        DIM sCrews
        on error resume next
        
        dim conDanaos
        Set conDanaos=server.CreateObject("ADODB.Connection")
	    conDanaos.Open "Provider=MSDAORA.1;User ID=wfadmin; Password=wfa0m1n;Data Source=AIX;"
        
        SQL = " SELECT   pv.persons_code, p.persons_name || ' ' || p.PERSONS_SURNAME FULLNAME,"
        SQL = SQL &  "  P.PERSONS_SURNAME, pv.joining_rank,to_char(pv.JOINING_DATE,'dd Mon yy') doj,"
        SQL = SQL &  "          DECODE (pv.joining_rank,"
        SQL = SQL &  "                  'MST', '1',"
        SQL = SQL &  "                  'C/O', '2',"
        SQL = SQL &  "                  'EIC', '3',"
        SQL = SQL &  "                  'C/E', '4',"
        SQL = SQL &  "                  '1/E', '5'"
        SQL = SQL &  "                 ) ID"
        SQL = SQL &  "     FROM tankpac.persons_voyages pv, tankpac.persons p"
        SQL = SQL &  "    WHERE pv.persons_code = p.persons_code"
        SQL = SQL &  "      AND pv.vessel_code = '" & VCODE & "'"
        SQL = SQL &  "      AND (    pv.joining_date < TO_DATE ('" & inspectionDate & "', 'DD-Mon-YYYY')"
        SQL = SQL &  "      AND NVL (pv.sign_off_date, SYSDATE) > TO_DATE ('" & inspectionDate & "', 'DD-Mon-YYYY')"
        SQL = SQL &  "      )"        
        SQL = SQL &  "      AND UPPER (pv.joining_rank) IN ('MST', 'C/O', 'C/E', 'EIC', '1/E')"
        SQL = SQL &  " ORDER BY ID, joining_rank"
        
        RS.Open  SQL,conDanaos
        
        while not rs.EOF
            SELECT CASE RS("id") 
                CASE "1" 
                    msg = replace(msg,"MST_NAME","Capt. " & RS("PERSONS_SURNAME"))
                    msg = replace(msg,"MST_CODE",RS("persons_code"))
                CASE "2" 
                    msg = replace(msg,"CO_NAME",RS("joining_rank") & " " & RS("FULLNAME"))
                    msg = replace(msg,"CO_CODE",RS("persons_code"))
                CASE "3","4"
                    msg = replace(msg,"CE_NAME",RS("joining_rank") & " " & RS("FULLNAME"))
                    msg = replace(msg,"CE_NAME",RS("joining_rank") & " " & RS("FULLNAME"))

                    msg = replace(msg,"CE_CODE",RS("persons_code"))
                    msg = replace(msg,"CE_CODE",RS("persons_code"))

                CASE "5" 
                    msg = replace(msg,"AE_NAME",RS("joining_rank") & " " & RS("FULLNAME"))
                    msg = replace(msg,"AE_CODE",RS("persons_code"))
            END SELECT
            
            rs.MoveNext
        wend
                
        RS.Close
        
        updateMsgForCrew = msg
    end function   
    
    
    function getApproveMsg()
        Dim msg 
        on error resume next
        msg = msg & "Dear MST_NAME"  & vbCrLf & vbCrLf
        msg = msg & "We are pleased to advise you that"
        msg = msg & " the members of your VMT have qualified"
        msg = msg & " for the incentive award subsequent to the"
        msg = msg & " " & MOC_NAME & " Inspection completed at "
        msg = msg & INSP_PORT & " on " & INSP_DATE_DISP & ". " 
        msg = msg & vbCrLf        
        msg = msg & "Please make arrangements for a total of USD "
        msg = msg & "TOTAL_AMT"
        msg = msg & " award money to be shared equally among the VMT members"
        msg = msg & " during the month end and keep FPD Accounts informed."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & vbCrLf & " 1. MST_NAME SC No. MST_CODE"
        msg = msg & vbCrLf & " 2. CE_NAME SC No. CE_CODE"
        msg = msg & vbCrLf & " 3. CO_NAME SC No. CO_CODE"
        msg = msg & vbCrLf & " 4. AE_NAME SC No. AE_CODE"
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "The Management appreciates your contribution and"
        msg = msg & " all the efforts made by each member of the VMT,"
        msg = msg & " in ensuring desired results during the above inspection."
        msg = msg & "We sincerely hope that your vessel will continue to excel"
        msg = msg & " during all forthcoming Oil Company inspections."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Well Done !"
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Regards"
        msg = msg & vbCrLf 
        msg = msg & "MSQV"    
        
        getApproveMsg = msg
    end function
    
    
    function getRejectMsg()
        Dim msg 
        msg = msg & "Dear MST_NAME" 
        msg = msg & vbCrLf
        msg = msg & "We regret to advise you that the members of your VMT"
        msg = msg & " have not qualified for the incentive award subsequent to the "
        msg = msg & MOC_NAME & " Inspection completed at "
        msg = msg & INSP_PORT & " on " & INSP_DATE_DISP & ". "
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "The Management appreciates all the hard work that has been"
        msg = msg & " put into the conduct of this inspection and sincerely hopes"
        msg = msg & " that under your able leadership, results will be in line to"
        msg = msg & " qualify for the incentive award during forthcoming inspections."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Please discuss areas of improvement with your Officers"
        msg = msg & " and Crew to ensure that future inspections yield desired results."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Regards" & vbCrLf 
        msg = msg & "MSQV "
    
        getRejectMsg = msg
    end function
%>
