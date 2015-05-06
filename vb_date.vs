function checkDate(v_date,v_field_desc,v_form_name)
	dim dt,sDate
	sDate = v_date.value
	checkDate = ""
	if sDate="" then exit function
	
	if not IsDate(sDate) then
		msgbox "Please enter a valid date in " & v_field_desc,vbInformation,"Invalid date"
		window.event.srcElement.focus
		window.event.srcElement.select
		exit function
	end if
	dt = cdate(sDate)
	
	checkDate = day(dt) & "-" & monthname(month(dt),true) & "-" & year(dt)
end function

function valid_date(v_date,v_field_desc,v_form_name )	
	if (isnull(v_form_name)) then
		v_form_name = "form1"
	end if
	if (isnull(v_date.value) or v_date.value = "") then
		exit function
	end if
	
	v_split_date = Split(v_date.value,"/")
	v_ctr=1
	dim i
	for each i in v_split_date
		
		if v_ctr=1 then
			v_split_date_portion=i
		end if
		
		if v_ctr=2 then
			v_split_month_portion=i
			if IsNumeric(i) then
			v_to_change = true
			v_lng_month = clng(v_split_month_portion)
				select case v_lng_month 
					case 1 v_split_full_month_portion = "Jan"
					case 2 v_split_full_month_portion = "Feb"
					case 3 v_split_full_month_portion = "Mar"
					case 4 v_split_full_month_portion = "Apr"
					case 5 v_split_full_month_portion = "May"
					case 6 v_split_full_month_portion = "Jun"
					case 7 v_split_full_month_portion = "Jul"
					case 8 v_split_full_month_portion = "Aug"
					case 9 v_split_full_month_portion = "Sep"
					case 10 v_split_full_month_portion = "Oct"
					case 11 v_split_full_month_portion = "Nov"
					case 12 v_split_full_month_portion = "Dec"
				end select
			end if
		end if
		
		if v_ctr=3 then
			v_split_year_portion=i
			if clng(v_split_year_portion)<2000 then
				v_split_year_portion = "20"&i
			end if
		end if
		v_ctr=v_ctr+1
	next
	
	if v_to_change = true then
	v_date.value=v_split_date_portion&"-"&v_split_full_month_portion&"-"&v_split_year_portion
	end if
	
	k=v_form_name&"."&v_date.name&".focus"
	
	if mid(v_date.value,2,1)="-" then
		v_date.value = "0"&v_date.value
	end if
	
	if IsNumeric(mid(v_date.value,1,2)) and mid(v_date.value,3,1)="-" and mid(v_date.value,7,1)="-" and isnumeric(mid(v_date.value,8,4)) then
	else
		alert(v_field_desc&" : Please enter date in : DD-Mon-YYYY format. ie 15-Sep-2002")
		eval(k)
	exit function
	end if
	
	if (isdate(v_date.value) and clng(mid(v_date.value,8,4))>1800) then
	else
		if clng(mid(v_date.value,8,4))<1800 then
		alert(v_field_desc&" : Please enter date in : DD-Mon-YYYY format. ie 15-Sep-2002 (Check the Year Please)")
		eval(k)
		else
		alert(v_field_desc&" : The date is not a valid date")
		eval(k)
		end if
		exit function
	end if
	
	if instr(1,"JanFebMarAprMayJunJulAugSepOctNovDec",ucase(Mid(v_date.value,4,1))&lcase(Mid(v_date.value,5,2))) > 0 then
		if mid(v_date.value,4,3) <> (ucase(Mid(v_date.value,4,1))&lcase(Mid(v_date.value,5,2))) then
		v_date.value=mid(v_date.value,1,3)&ucase(Mid(v_date.value,4,1))&lcase(Mid(v_date.value,5,2))&mid(v_date.value,7,5)
		end if
	else
			alert(v_field_desc&" : Please enter date in : DD-Mon-YYYY format. ie 15-Sep-2002")
			eval(k)
		exit function
	end if
end function