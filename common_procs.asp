<%
	Function dispCheckBox(fieldName, dispList, valueList, selectedList, columnCount)
		retVal = "<TABLE WIDTH='100%' BORDER='0' CELLSPACING='0' CELLPADDING='0'>"
		retVal = retVal & "<TR CLASS='tabledata'>"

		cCnt = 1

		While InStr(1, dispList, "~") <> 0
			dItem = Mid(dispList, 1, InStr(1, dispList, "~") - 1)
			vItem = Mid(valueList, 1, InStr(1, valueList, "~") - 1)
			sItem = Mid(selectedList, 1, InStr(1, selectedList, "~") - 1)
			dispList = Mid(dispList, InStr(1, dispList, "~") + 1)
			valueList = Mid(valueList, InStr(1, valueList, "~") + 1)
			selectedList = Mid(selectedList, InStr(1, selectedList, "~") + 1)

			If cCnt mod (columnCount + 1) = 0 Then

				retVal = retVal & "</TR><TR CLASS='tabledata' HEIGHT='27'><TD>"

				If sItem = "1" Then
					retVal = retVal & "<INPUT TYPE='checkbox' NAME='" & fieldName & "' VALUE='" & vItem & "' CHECKED CLASS='chkbox'>&nbsp;"
				Else
					retVal = retVal & "<INPUT TYPE='checkbox' NAME='" & fieldName & "' VALUE='" & vItem & "' CLASS='chkbox'>&nbsp;"
				End If

				retVal = retVal & "<SPAN CLASS='tabledata'>" & dItem & "</SPAN>"
				retVal = retVal & "</TD>"
				cCnt = 1

			Else

				retVal = retVal & "<TD>"

				If sItem = "1" Then
					retVal = retVal & "<INPUT TYPE='checkbox' NAME='" & fieldName & "' VALUE='" & vItem & "' CHECKED CLASS='chkbox'>&nbsp;"
				Else
					retVal = retVal & "<INPUT TYPE='checkbox' NAME='" & fieldName & "' VALUE='" & vItem & "' CLASS='chkbox'>&nbsp;"
				End If

				retVal = retVal & "<SPAN CLASS='tabledata'>" & dItem & "</SPAN>"
				retVal = retVal & "</TD>"
				cCnt = cCnt + 1

			End If
		Wend

		retVal = retVal & "</TD>"
		retVal = retVal & "</TABLE>"

		dispCheckBox = retVal
	End Function

	Function dateTimeDisplay(inputDate)
		dim retVal
		retVal = Mid(WeekdayName(DatePart("w", inputDate)), 1, 3) & ", "
		retVal = retVal & DatePart("d", inputDate) & " "
		retVal = retVal & Mid(MonthName(DatePart("m", inputDate)), 1, 3) & " "
		retVal = retVal & DatePart("YYYY", inputDate) & " "
		retVal = retVal & DatePart("H", inputDate) & ":" & DatePart("n", inputDate) & ":" & DatePart("s", inputDate) & " "
	
		dateTimeDisplay = retVal
	End Function

	Function dispBlankIfZero(inputVal)
		If inputVal = "0" Then
			dispBlankIfZero = ""
		Else
			dispBlankIfZero = inputVal
		End If
	End Function

	Function setAppVars(user_id, user_name, access_level, fleet_code, extra1, extra2, extra3)

		ipaddr_in = Request.ServerVariables ("REMOTE_ADDR") & "."
		ipaddr_out = ""
	
		While instr(ipaddr_in, ".") <> 0 
			ipaddr_out = ipaddr_out & Mid(ipaddr_in, 1, instr(ipaddr_in, ".") - 1)
			ipaddr_in = Mid(ipaddr_in, instr(ipaddr_in, ".") + 1, 100)
		Wend

		v_app_values = Trim(user_id) & "~~"		' User Id
		v_app_values = v_app_values & Trim(user_name) & "~~"	' User Full Name
		v_app_values = v_app_values & Trim(access_level) & "~~"	'Access Level
		v_app_values = v_app_values & Trim(fleet_code) & "~~"	'Fleet Code
		v_app_values = v_app_values & Trim(extra1) & "~~"	'Left for future use
		v_app_values = v_app_values & Trim(extra2) & "~~"	'Left for future use
		v_app_values = v_app_values & Trim(extra3) & "~~"	'Left for future use

		Application("v_wls_app_" & ipaddr_out) = v_app_values

		setAppVars = v_app_values
	End Function

	Function getAppVar(varName)

		Dim appValArray(10)
		Select Case Ucase(Trim(varName))
			case "USER_ID"
				place = 1
			case "USER_NAME"
				place = 2
			case "ACCESS_LEVEL"
				place = 3
			case "FLEET_CODE"
				place = 4
			case "EXTRA1"
				place = 5
			case "EXTRA2"
				place = 6
			case "EXTRA3"
				place = 7
			case else
				place = 0
		End Select
				
		ipaddr_in = Request.ServerVariables ("REMOTE_ADDR") & "."
		ipaddr_out = ""

		While instr(ipaddr_in, ".") <> 0 
			ipaddr_out = ipaddr_out & Mid(ipaddr_in, 1, instr(ipaddr_in, ".") - 1)
			ipaddr_in = Mid(ipaddr_in, instr(ipaddr_in, ".") + 1, 100)
		Wend

		vAppValString = Application("v_wls_app_" & ipaddr_out)
		i = 1
		While Instr(1, vAppValString, "~~") <> 0
			appValArray(i) = Mid(vAppValString, 1, Instr(1, vAppValString, "~~") - 1)
			vAppValString = Mid(vAppValString, Instr(1, vAppValString, "~~") + 2)
			i = i + 1
		Wend
		getAppVar = appValArray(place)
	End Function

	Function getDate(inputDate)

		givenDate = inputDate
		
		If givenDate = "" Then
			givenDate = Now()
		End If

		v_day = DatePart("d", givenDate)
		v_month = DatePart("m", givenDate)
		v_year = DatePart("yyyy", givenDate)
	
		If Len(v_day) = 1 Then
			v_day = "0" & v_day
		End If

		If Len(v_month) = 1 Then
			v_month = "0" & v_month
		End If

		getDate = v_day & "/" & v_month & "/" & v_year
	End Function
	
Function FormatDateTimeValues(dtDate, iType)
    Dim sTempValue
    
    If dtDate = "" Then
        FormatDateTimeValues = "NO_DATE"
        Exit Function
    End If
    
    If Isnull(dtDate) then
		FormatDateTimeValues = "NO_DATE"
		Exit Function
	End IF
    
    If iType = 1 Then
        sTempValue = Day(dtDate) & " " & MonthName(Month(dtDate), True) & " " & Year(dtDate)
    ElseIf iType = 2 Then
        sTempValue = Day(dtDate) & "-" & MonthName(Month(dtDate),true) & "-" & Year(dtDate)
    ElseIf iType = 3 Then
        sTempValue = Day(dtDate) & "-" & Month(dtDate) & "-" & Year(dtDate) & "-" & Hour(dtDate) & "-" & Minute(dtDate) & "-" & Second(dtDate)
    ElseIf iType = 4 Then
        sTempValue = Day(dtDate) & "-" & MonthName(Month(dtDate), True) & "-" & Year(dtDate) & "-" & Hour(dtDate) & "-" & Minute(dtDate) & "-" & Second(dtDate)
    ElseIf iType = 5 Then
        sTempValue = Day(dtDate) & "-" & MonthName(Month(dtDate), False) & "-" & Year(dtDate) & "-" & Hour(dtDate) & "-" & Minute(dtDate) & "-" & Second(dtDate)
    End If

    FormatDateTimeValues = sTempValue
End Function
Function IIF(expr, trueVal, falseVal)
	If expr Then
		IIF = trueVal
	Else
		IIF = falseVal
	End If
End function
Function ToHTML(strText)
    Dim strHTML
    if isnull(strText) then
		ToHTML=""
		exit function
    end if
    strHTML = strText
    strHTML = Replace(strHTML,"<","&lt;")
    strHTML = Replace(strHTML,">","&gt;")
    strHTML = Replace(strHTML, vbCrLf, "<BR>")
    strHTML = Replace(strHTML, vbLf, "<BR>")
    strHTML = Replace(strHTML, vbCr, "<BR>")
    Do
        strHTML = Replace(strHTML, "  ", " &nbsp;")
    Loop While InStr(1, strHTML, "  ", vbTextCompare) > 1
    ToHTML = strHTML
End Function
'************ToggleColor***********
dim sColor1,sColor2
sColor1 = "white"
sColor2 = "lightyellow"
Sub InitToggleColor(s1,s2)
	sColor1 = s1
	sColor2 = s2
end sub
function ToggleColor(sCol)
	if sCol = sColor1 then
		sCol = sColor2
	else
		sCol = sColor1
	end if
	ToggleColor = sCol
end function
'************ToggleColor***********
%>