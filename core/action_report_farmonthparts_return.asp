<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<%		
    act = Request("type")
	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	jobmonth = request("jobmonth")
	jobyear = request("jobyear")
	job_tech_model = request("job_tech_model")
	job_tech_type = request("job_tech_type")
	stepno = request("stepno")
	
	
	if stepno = "1" then 
	
	msg1 = "FAR Report by Spare Parts has been reset and prepare to generate new result....."
		   
	url = "action_report_farmonthparts.asp?type=rpt_farparts_year&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & TotalSales & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type 	
	
	elseif stepno = "2" then 
	
	   sql = "select top 1 parts_code from tblrpr_farParts where fa_done='N' "
	   parts_code = selectid(sql)
	   
	   sql = "select count(id) as sTotal from tblrpr_farParts "
	   sTotal = selectid(sql)
	   
	   sql = "select count(id) as TotalDone from tblrpr_farParts where fa_done='Y' "
	   TotalDone = selectid(sql)
	   
	   msg1 =  "Generating FAR by Spare Parts in the progress....(Done Items: " & TotalDone & " of " & sTotal & ")... System is processing Spare Part Record for : [" & parts_code & "]"
		
	   url = "action_report_farmonthparts.asp?type=rpt_farparts_year&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & TotalSales & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type 	
	
	end if

%>

<html>
<head>
<title>FAR Report Processing....</title>
<meta http-equiv="refresh" content="3; URL=<%=url%>">
<meta name="keywords" content="automatic redirection">
</head>
<body>
<h2>FAR Spare Parts Report Processing, please wait until process complete.

<%=msg1%></h2>
</body>
</html>

