// This Javascript modified by Sethu Subramanian Rengarajan on 02nd Sep, 2002
// This is verified only for "DD-MON-YYYY" Format.
// Please check the correctness, if you are using for other formats.

var weekend = [0,6];
var weekendColor = "#add8e6";
var fontface = "Verdana";
var fontsize = 1;

var gNow = new Date();
var ggWinCal;
isNav = (navigator.appName.indexOf("Netscape") != -1) ? true : false;
isIE = (navigator.appName.indexOf("Microsoft") != -1) ? true : false;

Calendar.Months = ["January", "February", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December"];

// Non-Leap year Month days..
Calendar.DOMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
// Leap year Month days..
Calendar.lDOMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function Calendar(p_item, p_WinCal, p_month, p_year, p_format) {
	if ((p_month == null) && (p_year == null))	return;

	if (p_WinCal == null)
		this.gWinCal = ggWinCal;
	else
		this.gWinCal = p_WinCal;
	
	if (p_month == null) {
		this.gMonthName = null;
		this.gMonth = null;
		this.gYearly = true;
	} else {
		this.gMonthName = Calendar.get_month(p_month);
		this.gMonth = new Number(p_month);
		this.gYearly = false;
	}

	this.gYear = p_year;
	this.gFormat = p_format;
	this.gBGColor = "white";
	this.gFGColor = "black";
	this.gTextColor = "black";
	this.gHeaderColor = "black";
	this.gReturnItem = p_item;
}

Calendar.get_month = Calendar_get_month;
Calendar.get_daysofmonth = Calendar_get_daysofmonth;
Calendar.calc_month_year = Calendar_calc_month_year;
Calendar.print = Calendar_print;

function Calendar_get_month(monthNo) {
	return Calendar.Months[monthNo];
}

function Calendar_get_daysofmonth(monthNo, p_year) {
	/* 
	Check for leap year ..
	1.Years evenly divisible by four are normally leap years, except for... 
	2.Years also evenly divisible by 100 are not leap years, except for... 
	3.Years also evenly divisible by 400 are leap years. 
	*/
	if ((p_year % 4) == 0) {
		if ((p_year % 100) == 0 && (p_year % 400) != 0)
			return Calendar.DOMonth[monthNo];
	
		return Calendar.lDOMonth[monthNo];
	} else
		return Calendar.DOMonth[monthNo];
}

function Calendar_calc_month_year(p_Month, p_Year, incr) {
	/* 
	Will return an 1-D array with 1st element being the calculated month 
	and second being the calculated year 
	after applying the month increment/decrement as specified by 'incr' parameter.
	'incr' will normally have 1/-1 to navigate thru the months.
	*/
	var ret_arr = new Array();
	
	if (incr == -1) {
		// B A C K W A R D
		if (p_Month == 0) {
			ret_arr[0] = 11;
			ret_arr[1] = parseInt(p_Year) - 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) - 1;
			ret_arr[1] = parseInt(p_Year);
		}
	} else if (incr == 1) {
		// F O R W A R D
		if (p_Month == 11) {
			ret_arr[0] = 0;
			ret_arr[1] = parseInt(p_Year) + 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) + 1;
			ret_arr[1] = parseInt(p_Year);
		}
	}
	
	return ret_arr;
}

function Calendar_print() {
	ggWinCal.print();
}

function Calendar_calc_month_year(p_Month, p_Year, incr) {
	/* 
	Will return an 1-D array with 1st element being the calculated month 
	and second being the calculated year 
	after applying the month increment/decrement as specified by 'incr' parameter.
	'incr' will normally have 1/-1 to navigate thru the months.
	*/
	var ret_arr = new Array();
	
	if (incr == -1) {
		// B A C K W A R D
		if (p_Month == 0) {
			ret_arr[0] = 11;
			ret_arr[1] = parseInt(p_Year) - 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) - 1;
			ret_arr[1] = parseInt(p_Year);
		}
	} else if (incr == 1) {
		// F O R W A R D
		if (p_Month == 11) {
			ret_arr[0] = 0;
			ret_arr[1] = parseInt(p_Year) + 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) + 1;
			ret_arr[1] = parseInt(p_Year);
		}
	}
	
	return ret_arr;
}

// This is for compatibility with Navigator 3, we have to create and discard one object before the prototype object exists.
new Calendar();

Calendar.prototype.getMonthlyCalendarCode = function() {
	var vCode = "";
	var vHeader_Code = "";
	var vData_Code = "";
	
	// Begin Table Drawing code here..
	vCode = vCode + "<TABLE width='100%' align=left BORDER=1 BGCOLOR=\"" + this.gBGColor + "\">";
	
	vHeader_Code = this.cal_header();
	vData_Code = this.cal_data();
	vCode = vCode + vHeader_Code + vData_Code;
	
	vCode = vCode + "</TABLE>";
	
	return vCode;
}

Calendar.prototype.show = function() {
	var vCode = "";
	
	this.gWinCal.document.open();

	// Setup the page...
	this.wwrite("<html>");
	this.wwrite("<head><title>Calendar</title>");
	this.wwrite("</head>");

	this.wwrite("<body " + 
		"link=\"" + this.gLinkColor + "\" " + 
		"vlink=\"" + this.gLinkColor + "\" " +
		"alink=\"" + this.gLinkColor + "\" " +
		"text=\"" + this.gTextColor + "\">");
	this.wwriteA("<SPAN ALIGN='center'><FONT FACE='" + fontface + "' SIZE=1 COLOR='green'><B>");
	this.wwriteA(this.gMonthName + " " + this.gYear);
	this.wwriteA("</B></SPAN><BR>");

	// Show navigation buttons
	var prevMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, -1);
	var prevMM = prevMMYYYY[0];
	var prevYYYY = prevMMYYYY[1];

	var nextMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, 1);
	var nextMM = nextMMYYYY[0];
	var nextYYYY = nextMMYYYY[1];
	
	//this.wwrite("<TABLE WIDTH='100%' BORDER=1 align=center CELLSPACING=0 CELLPADDING=0 BGCOLOR='#e0e0e0'><TR><TD ALIGN=center>");
	this.wwrite("<TABLE WIDTH='100%' BORDER=1 align=center CELLSPACING=0 CELLPADDING=0 BGCOLOR='palegoldenrod'><TR>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (parseInt(this.gYear)-1) + "', '" + this.gFormat + "'" +
		");" +
		"\" TITLE='Previous Year'><TD ALIGN=center TITLE='Previous Year'><FONT SIZE=-1>Y-</FONT></TD><\/A>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', '" + prevMM + "', '" + prevYYYY + "', '" + this.gFormat + "'" +
		");" +
		"\" TITLE='Previous Month'><TD ALIGN=center TITLE='Previous Month'><FONT SIZE=-1>M-</FONT></TD><\/A>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', '" + nextMM + "', '" + nextYYYY + "', '" + this.gFormat + "'" +
		");" +
		"\"><TD ALIGN=center TITLE='Next Month'><FONT SIZE=-1>M+</FONT></TD><\/A>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (parseInt(this.gYear)+1) + "', '" + this.gFormat + "'" +
		");" +
		"\"><TD ALIGN=center TITLE='Next Year'><FONT SIZE=-1>Y+</FONT></TD><\/A></TR></TABLE>");
	this.wwrite("<TABLE><TR HEIGHT='2px'><TD></TD></TR></TABLE>");

	// Get the complete calendar code for the month..
	vCode = this.getMonthlyCalendarCode();
	this.wwrite(vCode);

	this.wwrite("</font></body></html>");
	this.gWinCal.document.close();
}

Calendar.prototype.showY = function() {
	var vCode = "";
	var i;
	var vr, vc, vx, vy;		// Row, Column, X-coord, Y-coord
	var vxf = 285;			// X-Factor
	var vyf = 200;			// Y-Factor
	var vxm = 10;			// X-margin
	var vym;				// Y-margin
	if (isIE)	vym = 75;
	else if (isNav)	vym = 25;
	
	this.gWinCal.document.open();

	this.wwrite("<html>");
	this.wwrite("<head><title>Calendar</title>");
	this.wwrite("<style type='text/css'>\n<!--");
	for (i=0; i<12; i++) {
		vc = i % 3;
		if (i>=0 && i<= 2)	vr = 0;
		if (i>=3 && i<= 5)	vr = 1;
		if (i>=6 && i<= 8)	vr = 2;
		if (i>=9 && i<= 11)	vr = 3;
		
		vx = parseInt(vxf * vc) + vxm;
		vy = parseInt(vyf * vr) + vym;

		this.wwrite(".lclass" + i + " {position:absolute;top:" + vy + ";left:" + vx + ";}");
	}
	this.wwrite("-->\n</style>");
	this.wwrite("</head>");

	this.wwrite("<body bgcolor='palegoldenrod'" + 
		"link=\"" + this.gLinkColor + "\" " + 
		"vlink=\"" + this.gLinkColor + "\" " +
		"alink=\"" + this.gLinkColor + "\" " +
		"text=\"" + this.gTextColor + "\">");
	this.wwrite("<FONT FACE='" + fontface + "' SIZE=1><B>");
	this.wwrite("Year : " + this.gYear);
	this.wwrite("</B><BR>");

	// Show navigation buttons
	var prevYYYY = parseInt(this.gYear) - 1;
	var nextYYYY = parseInt(this.gYear) + 1;
	
	this.wwrite("<TABLE WIDTH='100%' BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR='#e0e0e0'><TR><TD ALIGN=center>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', null, '" + prevYYYY + "', '" + this.gFormat + "'" +
		");" +
		"\" alt='Prev Year'><<<\/A></TD><TD ALIGN=center>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"javascript:window.print();\">Print</A></TD><TD ALIGN=center>");
	this.wwrite("<A STYLE='TEXT-DECORATION: none' HREF=\"" +
		"javascript:window.opener.Build(" + 
		"'" + this.gReturnItem + "', null, '" + nextYYYY + "', '" + this.gFormat + "'" +
		");" +
		"\">>><\/A></TD></TR></TABLE><BR>");

	// Get the complete calendar code for each month..
	var j;
	for (i=11; i>=0; i--) {
		if (isIE)
			this.wwrite("<DIV ID=\"layer" + i + "\" CLASS=\"lclass" + i + "\">");
		else if (isNav)
			this.wwrite("<LAYER ID=\"layer" + i + "\" CLASS=\"lclass" + i + "\">");

		this.gMonth = i;
		this.gMonthName = Calendar.get_month(this.gMonth);
		vCode = this.getMonthlyCalendarCode();
		this.wwrite(this.gMonthName + "/" + this.gYear + "<BR>");
		this.wwrite(vCode);

		if (isIE)
			this.wwrite("</DIV>");
		else if (isNav)
			this.wwrite("</LAYER>");
	}

	this.wwrite("</font><BR></body></html>");
	this.gWinCal.document.close();
}

Calendar.prototype.wwrite = function(wtext) {
	this.gWinCal.document.writeln(wtext);
}

Calendar.prototype.wwriteA = function(wtext) {
	this.gWinCal.document.write(wtext);
}

Calendar.prototype.cal_header = function() {
	var vCode = "";
	
	vCode = vCode + "<TR BGCOLOR='maroon'>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Sun</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Mon</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Tue</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Wed</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Thu</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='14%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Fri</B></FONT></TD>";
	vCode = vCode + "<TD WIDTH='16%' ALIGN='center'><FONT SIZE='1' FACE='" + fontface + "' COLOR='" + this.gHeaderColor + "'><B>Sat</B></FONT></TD>";
	vCode = vCode + "</TR>";
	
	return vCode;
}

Calendar.prototype.cal_data = function() {
	var vDate = new Date();
	vDate.setDate(1);
	vDate.setMonth(this.gMonth);
	vDate.setFullYear(this.gYear);

	var vFirstDay=vDate.getDay();
	var vDay=1;
	var vLastDay=Calendar.get_daysofmonth(this.gMonth, this.gYear);
	var vOnLastDay=0;
	var vCode = "";

	/*
	Get day for the 1st of the requested month/year..
	Place as many blank cells before the 1st day of the month as necessary. 
	*/

	vCode = vCode + "<TR>";
	for (i=0; i<vFirstDay; i++) {
		vCode = vCode + "<TD WIDTH='14%' VALIGN='center' ALIGN='center'" + this.write_weekend_string(i) + "><B><FONT SIZE='1' FACE='" + fontface + "'>&nbsp;</FONT></B></TD>";
	}

	// Write rest of the 1st week
	for (j=vFirstDay; j<7; j++) {
		vCode = vCode + 
			"<A STYLE='TEXT-DECORATION: none' HREF='#' " + 
				"onClick=\"self.opener.document." + this.gReturnItem + ".value='" + 
				this.format_data(vDay) + 
				"';window.close();\">" + "<TD WIDTH='14%' ALIGN='center'" + this.write_weekend_string(j) + "><FONT SIZE='1' FACE='" + fontface + "'>" +
				this.format_day(vDay) + 
			"</TD></A>" + 
			"</FONT>";
		vDay=vDay + 1;
	}
	vCode = vCode + "</TR>";

	// Write the rest of the weeks
	for (k=2; k<7; k++) {
		vCode = vCode + "<TR>";

		for (j=0; j<7; j++) {
			vCode = vCode +  
				"<A STYLE='TEXT-DECORATION: none' HREF='#' " + 
					"onClick=\"self.opener.document." + this.gReturnItem + ".value='" + 
					this.format_data(vDay) + 
					"';window.close();\">" + "<TD WIDTH='14%' ALIGN='center'" + this.write_weekend_string(j) + "><FONT SIZE='1' FACE='" + fontface + "'>" + 
				this.format_day(vDay) + 
				"</TD></A>" + 
				"</FONT>";
			vDay=vDay + 1;

			if (vDay > vLastDay) {
				vOnLastDay = 1;
				break;
			}
		}

		if (j == 6)
			vCode = vCode + "</TR>";
		if (vOnLastDay == 1)
			break;
	}
	
	// Fill up the rest of last week with proper blanks, so that we get proper square blocks
	for (m=1; m<(7-j); m++) {
		if (this.gYearly)
			vCode = vCode + "<TD WIDTH='14%' ALIGN='center'" + this.write_weekend_string(j+m) + 
			"><B><FONT SIZE='1' FACE='" + fontface + "'>&nbsp;</FONT></B></TD>";
		else
			vCode = vCode + "<TD WIDTH='14%' ALIGN='center'" + this.write_weekend_string(j+m) + 
			"><B><FONT SIZE='1' FACE='" + fontface + "'>&nbsp;</FONT></B></TD>";
	}
	
	return vCode;
}

Calendar.prototype.format_day = function(vday) {
	var vNowDay = gNow.getDate();
	var vNowMonth = gNow.getMonth();
	var vNowYear = gNow.getFullYear();

	if (vday == vNowDay && this.gMonth == vNowMonth && this.gYear == vNowYear)
		return ("<FONT COLOR=\"RED\"><B>" + vday + "</B></FONT>");
	else
		return (vday);
}

Calendar.prototype.write_weekend_string = function(vday) {
	var i;

	// Return special formatting for the weekend day.
	for (i=0; i<weekend.length; i++) {
		if (vday == weekend[i])
			return (" BGCOLOR=\"" + weekendColor + "\"");
	}
	
	return "";
}

Calendar.prototype.format_data = function(p_day) {
	var vData;
	var vMonth = 1 + this.gMonth;
	vMonth = (vMonth.toString().length < 2) ? "0" + vMonth : vMonth;
	// Commented by Sethu Subramanian Rengarajan change it to Title Case on 2nd Sep, 2002
	// var vMon = Calendar.get_month(this.gMonth).substr(0,3).toUpperCase();
	var vMon = Calendar.get_month(this.gMonth).substr(0,1).toUpperCase()+Calendar.get_month(this.gMonth).substr(1,2).toLowerCase();
	var vFMon = Calendar.get_month(this.gMonth).toUpperCase();
	var vY4 = new String(this.gYear);
	var vY2 = new String(this.gYear.substr(2,2));
	var vDD = (p_day.toString().length < 2) ? "0" + p_day : p_day;

	switch (this.gFormat) {
		case "MM\/DD\/YYYY" :
			vData = vMonth + "\/" + vDD + "\/" + vY4;
			break;
		case "MM\/DD\/YY" :
			vData = vMonth + "\/" + vDD + "\/" + vY2;
			break;
		case "MM-DD-YYYY" :
			vData = vMonth + "-" + vDD + "-" + vY4;
			break;
		case "MM-DD-YY" :
			vData = vMonth + "-" + vDD + "-" + vY2;
			break;

		case "DD\/MON\/YYYY" :
			vData = vDD + "\/" + vMon + "\/" + vY4;
			break;
		case "DD\/MON\/YY" :
			vData = vDD + "\/" + vMon + "\/" + vY2;
			break;
		case "DD-MON-YYYY" :
			vData = vDD + "-" + vMon + "-" + vY4;
			break;
		case "DD-Mon-YYYY" :
			vData = vDD + "-" + vMon + "-" + vY4;
			break;
		case "DD-MON-YY" :
			vData = vDD + "-" + vMon + "-" + vY2;
			break;

		case "DD\/MONTH\/YYYY" :
			vData = vDD + "\/" + vFMon + "\/" + vY4;
			break;
		case "DD\/MONTH\/YY" :
			vData = vDD + "\/" + vFMon + "\/" + vY2;
			break;
		case "DD-MONTH-YYYY" :
			vData = vDD + "-" + vFMon + "-" + vY4;
			break;
		case "DD-MONTH-YY" :
			vData = vDD + "-" + vFMon + "-" + vY2;
			break;

		case "DD\/MM\/YYYY" :
			vData = vDD + "\/" + vMonth + "\/" + vY4;
			break;
		case "DD\/MM\/YY" :
			vData = vDD + "\/" + vMonth + "\/" + vY2;
			break;
		case "DD-MM-YYYY" :
			vData = vDD + "-" + vMonth + "-" + vY4;
			break;
		case "DD-MM-YY" :
			vData = vDD + "-" + vMonth + "-" + vY2;
			break;

		default :
			vData = vMonth + "\/" + vDD + "\/" + vY4;
	}

	return vData;
}

function Build(p_item, p_month, p_year, p_format) {
	var p_WinCal = ggWinCal;
	gCal = new Calendar(p_item, p_WinCal, p_month, p_year, p_format);

	// Customize your Calendar here..
	//gCal.gBGColor="white";
	gCal.gBGColor="palegoldenrod";
	gCal.gLinkColor="black";
	gCal.gTextColor="black";
	gCal.gHeaderColor="white";

	// Choose appropriate show function
	if (gCal.gYearly)	gCal.showY();
	else	gCal.show();
}

function show_calendar(v_field_name,v_exist_date) {
	/* 
		p_month : 0-11 for Jan-Dec; 12 for All Months.
		p_year	: 4-digit year
		p_format: Date format (mm/dd/yyyy, dd/mm/yy, ...)
		p_item	: Return Item.
	*/
	p_item = arguments[0];
	if (arguments[1] == null)
		p_month = new String(gNow.getMonth());
	else
		p_month = arguments[1];
	if (arguments[2] == "" || arguments[2] == null)
		p_year = new String(gNow.getFullYear().toString());
	else
		p_year = arguments[2];
	if (arguments[3] == null)
		p_format = "DD-MON-YYYY";
	else
		p_format = arguments[3];
	if (v_field_name.value == '' || v_field_name.value==null) {
				p_month=new String(gNow.getMonth());
				p_year = new String(gNow.getFullYear().toString());
	}

	w = screen.width;
        if(w<=800){ 
	vWinCal = window.open("", "Calendar", 
		"width=220,height=200,status=no,resizable=no,top=150,left=250");
 	}else{
	vWinCal = window.open("", "Calendar", 
		"width=275,height=200,status=no,resizable=no,top=60,left=430");
	}
	vWinCal.opener = self;
	ggWinCal = vWinCal;
	
	//if date is 1 jan 2005 pad 0 in front
	if(eval(v_field_name).value.length == 10)
		eval(v_field_name).value = "0" + eval(v_field_name).value;
	if ((v_exist_date !=0) && (p_format="DD-MON-YYYY")) {
	// alert(eval(v_field_name).value);
		//p_month="5";
		p_month = eval(v_field_name).value.substr(3,3);
		//alert(p_month);
		//p_month="5"
		
		switch (p_month.toUpperCase()) {
		case "JAN" :
			p_month = "0"
			break;
		case "FEB" :
			p_month = "1"
			break;
		case "MAR" :
			p_month = "2"
			break;
		case "APR" :
			p_month = "3"
			break;
		case "MAY" :
			p_month = "4"
			break;
		case "JUN" :
			p_month = "5"
			break;
		case "JUL" :
			p_month = "6"
			break;
		case "AUG" :
			p_month = "7"
			break;
		case "SEP" :
			p_month = "8"
			break;
		case "OCT" :
			p_month = "9"
			break;
		case "NOV" :
			p_month = "10"
			break;
		case "DEC" :
			p_month = "11"
			break;
		default :
			p_month = "0";
	}
		
		p_year=eval(v_field_name).value.substr(7,4);
		if(p_year.length==3)
			p_year=eval(v_field_name).value.substr(6,4);
		//alert(p_year);
		//alert(p_month);
		}
	Build(p_item, p_month, p_year, p_format);
	
}
/*
Yearly Calendar Code Starts here
*/
function show_yearly_calendar(p_item, p_year, p_format) {
	// Load the defaults..
	if (p_year == null || p_year == "")
		p_year = new String(gNow.getFullYear().toString());
	if (p_format == null || p_format == "")
		p_format = "MM/DD/YYYY";

	var vWinCal = window.open("", "Calendar", "scrollbars=yes");
	vWinCal.opener = self;
	ggWinCal = vWinCal;

	Build(p_item, null, p_year, p_format);
}

function convert_date(field1)
{
var fLength = field1.value.length; // Length of supplied field in characters.
var divider_values = new Array ('-','.','/',' ',':','_',','); // Array to hold permitted date seperators.  Add in '\' value
var array_elements = 7; // Number of elements in the array - divider_values.
var day1 = new String(null); // day value holder
var month1 = new String(null); // month value holder
var year1 = new String(null); // year value holder
var divider1 = null; // divider holder
var outdate1 = null; // formatted date to send back to calling field holder
var counter1 = 0; // counter for divider looping 
var divider_holder = new Array ('0','0','0'); // array to hold positions of dividers in dates
var s = String(field1.value); // supplied date value variable

//If field is empty do nothing
if ( fLength == 0 ) {
   return true;
}

// Deal with today or now
if ( field1.value.toUpperCase() == 'NOW' || field1.value.toUpperCase() == 'TODAY' ) {
   
	var newDate1 = new Date();
	
  		if (navigator.appName == "Netscape") {
    		var myYear1 = newDate1.getYear() + 1900;
  		}
  		else {
  			var myYear1 =newDate1.getYear();
  		}
  
	var myMonth1 = newDate1.getMonth()+1;  
	var myDay1 = newDate1.getDate();
	field1.value = myDay1 + "/" + myMonth1 + "/" + myYear1;
	fLength = field1.value.length;//re-evaluate string length.
	s = String(field1.value)//re-evaluate the string value.
}

//Check the date is the required length
if ( fLength != 0 && (fLength < 6 || fLength > 11) ) {
	invalid_date(field1);
	return false;   
	}

// Find position and type of divider in the date
for ( var i=0; i<3; i++ ) {
	for ( var x=0; x<array_elements; x++ ) {
		if ( s.indexOf(divider_values[x], counter1) != -1 ) {
			divider1 = divider_values[x];
			divider_holder[i] = s.indexOf(divider_values[x], counter1);
		   //alert(i + " divider1 = " + divider_holder[i]);
			counter1 = divider_holder[i] + 1;
			//alert(i + " counter1 = " + counter1);
			break;
		}
 	}
 }

// if element 2 is not 0 then more than 2 dividers have been found so date is invalid.
if ( divider_holder[2] != 0 ) {
   invalid_date(field1);
	return false;   
}

// See if no dividers are present in the date string.
if ( divider_holder[0] == 0 && divider_holder[1] == 0 ) { 
   
		//continue processing
		if ( fLength == 6 ) {//ddmmyy
   		day1 = field1.value.substring(0,2);
     		month1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,6);
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
			
		else if ( fLength == 7 ) {//ddmmmy
   		day1 = field1.value.substring(0,2);
  			month1 = field1.value.substring(2,5);
  			year1 = field1.value.substring(5,7);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
  			if ( (year1 = validate_year(year1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
		else if ( fLength == 8 ) {//ddmmyyyy
   		day1 = field1.value.substring(0,2);
  			month1 = field1.value.substring(2,4);
  			year1 = field1.value.substring(4,8);
			}
		else if ( fLength == 9 ) {//ddmmmyyyy
   		day1 = field1.value.substring(0,2);
  			month1 = field1.value.substring(2,5);
  			year1 = field1.value.substring(5,9);
  			if ( (month1 = convert_month(month1)) == false ) {
   			invalid_date(field1);
				return false; 
				}
			}
		
		if ( (outdate1 = validate_date(day1,month1,year1)) == false ) {
   		alert("The value " + field1.value + " is not a vaild date.\n\r" +  
			"Please enter a valid date in the format dd/mm/yyyy");
			field1.value="";
			field1.focus();
			field1.select();
			return false;
			}

		field1.value = outdate1;
		return true;// All OK
		}
		
// 2 dividers are present so continue to process	
if ( divider_holder[0] != 0 && divider_holder[1] != 0 ) { 	
  	day1 = field1.value.substring(0, divider_holder[0]);
  	month1 = field1.value.substring(divider_holder[0] + 1, divider_holder[1]);
  	//alert(month1);
  	year1 = field1.value.substring(divider_holder[1] + 1, field1.value.length);
	}

if ( isNaN(day1) && isNaN(year1) ) { // Check day and year are numeric
	invalid_date(field1);
	return false;  
   }

if ( day1.length == 1 ) { //Make d day dd
   day1 = '0' + day1;  
}

if ( month1.length == 1 ) {//Make m month mm
	month1 = '0' + month1;   
}

if ( year1.length == 2 ) {//Make yy year yyyy
   if ( (year1 = validate_year(year1)) == false ) {
   	invalid_date(field1);
		return false;  
		}
}

if ( month1.length == 3 || month1.length == 4 ) {//Make mmm month mm
   if ( (month1 = convert_month(month1)) == false) {
   	alert("month1" + month1);
   	invalid_date(field1);
   	return false;  
   }
}

// Date components are OK
if ( (day1.length == 2 || month1.length == 2 || year1.length == 4) == false) {
   invalid_date(field1);
   return false;
}

//Validate the date
if ( (outdate1 = validate_date(day1, month1, year1)) == false ) {
   alert("The value " + field1.value + " is not a vaild date.\n\r" +  
	"Please enter a valid date in the format dd/mm/yyyy");
	field1.value="";
	field1.focus();
	field1.select();
	return false;
}

// Redisplay the date in dd/mm/yyyy format
field1.value = outdate1;
return true;//All is well

}

function convert_month(monthIn) {

var month_values = new Array ("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC");

monthIn = monthIn.toUpperCase(); 

if ( monthIn.length == 3 ) {
	for ( var i=0; i<12; i++ ) 
		{
   	if ( monthIn == month_values[i] ) 
   		{
			monthIn = i + 1;
			if ( i != 10 && i != 11 && i != 12 ) 
				{
   			monthIn = '0' + monthIn;
				}
			return monthIn;
			}
		}
	}

else if ( monthIn.length == 4 && monthIn == 'SEPT') {
   monthIn = '09';
   return monthIn;
	}
	
else {
	return false;
	} 
}

function invalid_date(inField) 
{
alert("The value " + inField.value + " is not in a vaild date format.\n\r" + 
        "Please enter date in the format dd/mm/yyyy");
inField.value="";
inField.focus();
inField.select();
return true   
}

function validate_date(day2, month2, year2)                                                      {                                                                                                
var DayArray = new Array(31,28,31,30,31,30,31,31,30,31,30,31);                                   var MonthArray = new Array("01","02","03","04","05","06","07","08","09","10","11","12");         
var inpDate = day2 + month2 + year2;                                                             
var filter=/^[0-9]{2}[0-9]{2}[0-9]{4}$/;                                                         

//Check ddmmyyyy date supplied
if (! filter.test(inpDate))                                                                        {                                                                                                return false;                                                                                    }                                                                                              
/* Check Valid Month */                                                                          filter=/01|02|03|04|05|06|07|08|09|10|11|12/ ;                                                   
if (! filter.test(month2))                                                                         {                                                                                              
  return false;                                                                                  
  }                                                                                              

/* Check For Leap Year */                                                                        
                                                 
var N = Number(year2);                                                                           if ( ( N%4==0 && N%100 !=0 ) || ( N%400==0 ) )                                                   {                                                                                                
                            
   DayArray[1]=29;                                                                               }                                                                                                                                                                                                 
                              
/* Check for valid days for month */                                                             for(var ctr=0; ctr<=11; ctr++)                                                                   
{                                                                                                   if (MonthArray[ctr]==month2)                                                                  
	{                                                                                                                           
      if (day2<= DayArray[ctr] && day2 >0 )                                                              {
        inpDate = day2 + '/' + month2 + '/' + year2;       
        return inpDate;
        }       
      else                                                                                               {                                                                                                return false;                                                                                    }                                                                                               }                                                                                            }                                                                                                                                                               
}

function validate_year(inYear) 
{
if ( inYear < 10 ) 
	{
   inYear = "20" + inYear;
   return inYear;
	}
else if ( inYear >= 10 )
	{
   inYear = "19" + inYear;
   return inYear;
	}
else 
	{
	return false;
	}   
}
var w = screen.width; 
var today=new Date();
var expiry=new Date(today.getTime() + 28 * 24 * 60 * 60 * 1000);
	
function setCookie(name,value){
	document.cookie=name + "=" + escape(value) + "; expires=" + expiry.toGMT
}	
setCookie("scrWidth",w)
function getCookie(name) {
	var re=new RegExp(name + "=([^;]+)");
	var value=re.exec(document.cookie);
	return (value != null) ? unescape(value[1]) : null;
}