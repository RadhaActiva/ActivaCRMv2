<!-- #include file="database/datastore.asp" -->
<%
act = Request("type")

redim jobmonth_list(13), jobyear_list(13)
redim fa_MD_over_month(13), fa_MD_under_month(13), fa_DS_over_month(13), fa_DS_under_month(13), fa_WI_over_month(13), fa_WI_under_month(13), fa_CF_over_month(13), fa_CF_under_month(13)


'----------------------------------------------------------------------------------------------------    
	
Select Case act

 
'----------------------------------------------------------------------------------------------------    
  Case "rpt_farmonth_parts_reset"

	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	job_tech_model = request("job_tech_model")
	
	
	if ordertype = "" then 
	   ordertype = "desc"
	end if
	
	if request("job_date_from") <> "" then
	   job_date_from = request("job_date_from")
	else
	   job_date_from = chkdate(DateAdd("d",-90,date()))
	end if
	
	if request("job_date_to") <> "" then
	   job_date_to = request("job_date_to")
	else
	   job_date_to = chkdate(date())
	end if
	
	if request("jobmonth") <> "" then
	   jobmonth = request("jobmonth")
	else
	   jobmonth = month(date())
	end if
	
	if request("jobyear") <> "" then
	   jobyear = request("jobyear")
	else
	   jobyear = year(date())
	end if  
  	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_farParts"
	CUD(sql)
	

	sql = "INSERT INTO tblrpr_farParts (parts_code, parts_desc)  SELECT distinct tbljob_parts.jobp_partcode, tbljob_parts.jobp_desc " & _ 
          "FROM tbljob_parts INNER JOIN tbljob ON tbljob_parts.job_code = tbljob.job_code " & _ 
	      "where tbljob.job_id is not null " & _
		  "and  year(tbljob.job_posteddate) = '" & jobyear & "' " & _
		  "and tbljob.job_status='Posted' and job_tech_type = '" & job_tech_type & "' "
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	CUD(sql)

	response.write "tblrpr_farParts reset done"

	url = "action_report_farmonthparts_return.asp?type=rpt_farparts_year&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & request("TotalSales") & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type & "&stepno=1"	
	
 
'----------------------------------------------------------------------------------------------------    
  Case "rpt_farparts_year"
  
	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	job_tech_model = request("job_tech_model")
	
	
	if ordertype = "" then 
	   ordertype = "desc"
	end if
	
	if request("job_date_from") <> "" then
	   job_date_from = request("job_date_from")
	else
	   job_date_from = chkdate(DateAdd("d",-90,date()))
	end if
	
	if request("job_date_to") <> "" then
	   job_date_to = request("job_date_to")
	else
	   job_date_to = chkdate(date())
	end if
	
	if request("jobmonth") <> "" then
	   jobmonth = request("jobmonth")
	else
	   jobmonth = month(date())
	end if
	
	if request("jobyear") <> "" then
	   jobyear = request("jobyear")
	else
	   jobyear = year(date())
	end if
	
	currentdate = "01-" & convertmonth(jobmonth) & "-" & jobyear
		
	jobmonth1 = jobmonth
	jobyear1 =  jobyear
	jobmonth_list(1) = jobmonth
	jobyear_list(1) = jobyear 
	
	jobmonth2 = month(DateAdd("m",-1,currentdate))
	jobyear2 =  year(DateAdd("m",-1,currentdate))
	jobmonth_list(2) = month(DateAdd("m",-1,currentdate))
	jobyear_list(2) = year(DateAdd("m",-1,currentdate))
	
	jobmonth3 = month(DateAdd("m",-2,currentdate))
	jobyear3 =  year(DateAdd("m",-2,currentdate))
	jobmonth_list(3) = month(DateAdd("m",-2,currentdate))
	jobyear_list(3) = year(DateAdd("m",-2,currentdate))
	
	jobmonth4 = month(DateAdd("m",-3,currentdate))
	jobyear4 =  year(DateAdd("m",-3,currentdate))
	jobmonth_list(4) = month(DateAdd("m",-3,currentdate))
	jobyear_list(4) = year(DateAdd("m",-3,currentdate))
	
	jobmonth5 = month(DateAdd("m",-4,currentdate))
	jobyear5 =  year(DateAdd("m",-4,currentdate))
	jobmonth_list(5) = month(DateAdd("m",-4,currentdate))
	jobyear_list(5) = year(DateAdd("m",-4,currentdate))
	
	jobmonth6 = month(DateAdd("m",-5,currentdate))
	jobyear6 =  year(DateAdd("m",-5,currentdate))
	jobmonth_list(6) = month(DateAdd("m",-5,currentdate))
	jobyear_list(6) = year(DateAdd("m",-5,currentdate))
	
	jobmonth7 = month(DateAdd("m",-6,currentdate))
	jobyear7 =  year(DateAdd("m",-6,currentdate))
	jobmonth_list(7) = month(DateAdd("m",-6,currentdate))
	jobyear_list(7) = year(DateAdd("m",-6,currentdate))
	
	jobmonth8 = month(DateAdd("m",-7,currentdate))
	jobyear8 =  year(DateAdd("m",-7,currentdate))
	jobmonth_list(8) = month(DateAdd("m",-7,currentdate))
	jobyear_list(8) = year(DateAdd("m",-7,currentdate))
	
	jobmonth9 = month(DateAdd("m",-8,currentdate))
	jobyear9 =  year(DateAdd("m",-8,currentdate))
	jobmonth_list(9) = month(DateAdd("m",-8,currentdate))
	jobyear_list(9) = year(DateAdd("m",-8,currentdate))
	
	jobmonth10 = month(DateAdd("m",-9,currentdate))
	jobyear10 =  year(DateAdd("m",-9,currentdate))
	jobmonth_list(10) = month(DateAdd("m",-9,currentdate))
	jobyear_list(10) = year(DateAdd("m",-9,currentdate))
	
	jobmonth11 = month(DateAdd("m",-10,currentdate))
	jobyear11 =  year(DateAdd("m",-10,currentdate))
	jobmonth_list(11) = month(DateAdd("m",-10,currentdate))
	jobyear_list(11) = year(DateAdd("m",-10,currentdate))
	
	jobmonth12 = month(DateAdd("m",-11,currentdate))
	jobyear12 =  year(DateAdd("m",-11,currentdate))
	jobmonth_list(12) = month(DateAdd("m",-11,currentdate))
	jobyear_list(12) = year(DateAdd("m",-11,currentdate))
	
	if request("job_tech_model") <> "" then
	   job_tech_model = replace(request("job_tech_model"), " ", "")
	   arrjob_tech_model = split(job_tech_model,",")
	   job_tech_model = replace(job_tech_model, ",", "','")
	   
	   listjob_tech_model = listjob_tech_model & job_tech_model
	   
	else
	   listjob_tech_model = ""
	   arrjob_tech_model = split("0,0",",")
	   
	end if
	
	
	function checkModelList(strv)
	for k = 0 to ubound(arrjob_tech_model)
		if arrjob_tech_model(k) = strv then 
		   checkModelList = true
		   exit for
		else
		   checkModelList = false
		end if
	next
	end function

	
	stillgotRecord = "No"
	
    ''' Loop tblrpr_farmonth
	sql1 = "SELECT top 1 id, parts_code, parts_desc, fa_month1_over, fa_month1_under, fa_month2_over, fa_month2_under, fa_month3_over, fa_month3_under, " & _
		   "fa_month_total_over, fa_month_total_under, fa_MD_over, fa_MD_under, fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
		   "FROM tblrpr_farParts where fa_done='N' "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

	'fa_month1_over
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month1_over = selectid(sql2)
	
	'fa_month1_under
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month1_under = selectid(sql2)

	'fa_month2_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth2 & "' and year(tbljob.job_posteddate) = '" & jobyear2 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month2_Over = selectid(sql2)
	
	'fa_month2_under,
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth2 & "' and year(tbljob.job_posteddate) = '" & jobyear2 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month2_Under = selectid(sql2)
	 
	'fa_month3_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth3 & "' and year(tbljob.job_posteddate) = '" & jobyear3 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month3_over = selectid(sql2)
	
	'fa_month3_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth3 & "' and year(tbljob.job_posteddate) = '" & jobyear3 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month3_under = selectid(sql2)
	
	'fa_month4_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth4 & "' and year(tbljob.job_posteddate) = '" & jobyear4 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month4_over = selectid(sql2)
	
	'fa_month4_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth4 & "' and year(tbljob.job_posteddate) = '" & jobyear4 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month4_under = selectid(sql2)

	'fa_month5_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth5 & "' and year(tbljob.job_posteddate) = '" & jobyear5 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month5_over = selectid(sql2)
	
	'fa_month5_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth5 & "' and year(tbljob.job_posteddate) = '" & jobyear5 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month5_under = selectid(sql2)	
	
	'fa_month6_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth6 & "' and year(tbljob.job_posteddate) = '" & jobyear6 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if		
	fa_month6_over = selectid(sql2)
	
	'fa_month6_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth6 & "' and year(tbljob.job_posteddate) = '" & jobyear6 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month6_under = selectid(sql2)	
	
	'fa_month7_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth7 & "' and year(tbljob.job_posteddate) = '" & jobyear7 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month7_over = selectid(sql2)
	
	'fa_month7_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth7 & "' and year(tbljob.job_posteddate) = '" & jobyear7 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month7_under = selectid(sql2)	
	
	'fa_month8_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth8 & "' and year(tbljob.job_posteddate) = '" & jobyear8 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month8_over = selectid(sql2)
	
	'fa_month8_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth8 & "' and year(tbljob.job_posteddate) = '" & jobyear8 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month8_under = selectid(sql2)		
	
	'fa_month9_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth9 & "' and year(tbljob.job_posteddate) = '" & jobyear9 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month9_over = selectid(sql2)
	
	'fa_month9_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth9 & "' and year(tbljob.job_posteddate) = '" & jobyear9 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month9_under = selectid(sql2)	
	
	'fa_month10_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth10 & "' and year(tbljob.job_posteddate) = '" & jobyear10 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month10_over = selectid(sql2)
	
	'fa_month10_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth10 & "' and year(tbljob.job_posteddate) = '" & jobyear10 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month10_under = selectid(sql2)
	
	'fa_month11_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth11 & "' and year(tbljob.job_posteddate) = '" & jobyear11 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month11_over = selectid(sql2)
	
	'fa_month11_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth11 & "' and year(tbljob.job_posteddate) = '" & jobyear11 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month11_under = selectid(sql2)
	
	'fa_month12_over, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth12 & "' and year(tbljob.job_posteddate) = '" & jobyear12 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month12_over = selectid(sql2)
	
	'fa_month12_under, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
          "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth12 & "' and year(tbljob.job_posteddate) = '" & jobyear12 & "' " & _
		  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month12_under = selectid(sql2)
	
	'fa_month_total_over, 
	fa_month_total_over = ChkNumber(fa_month1_over) + ChkNumber(fa_month2_over) + ChkNumber(fa_month3_over) + ChkNumber(fa_month4_over) + ChkNumber(fa_month5_over) + ChkNumber(fa_month6_over) + ChkNumber(fa_month7_over) + ChkNumber(fa_month8_over) + ChkNumber(fa_month9_over) + ChkNumber(fa_month10_over) + ChkNumber(fa_month11_over) + ChkNumber(fa_month12_over)
	
	'fa_month_total_under, 
	fa_month_total_under = ChkNumber(fa_month1_under) + ChkNumber(fa_month2_under) + ChkNumber(fa_month3_under) + ChkNumber(fa_month4_under) + ChkNumber(fa_month5_under) + ChkNumber(fa_month6_under) + ChkNumber(fa_month7_under) + ChkNumber(fa_month8_under) + ChkNumber(fa_month9_under) + ChkNumber(fa_month10_under) + ChkNumber(fa_month11_under) + ChkNumber(fa_month12_under) 
	
	'fa_MD_over_month
	for i = 1 to 12 
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='MD' and tbljob.job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if		
		fa_MD_over_month(i) = selectid(sql2)
	next
	
	'fa_MD_under_month
	for i = 1 to 12 
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='MD' and tbljob.job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if	
		fa_MD_under_month(i) = selectid(sql2)
	next
	
	'fa_DS_over,
	for i = 1 to 12 
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='DS' and tbljob.job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if	
		fa_DS_over_month(i) = selectid(sql2)
	next
	
	'fa_DS_under, 
	for i = 1 to 12  
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='DS' and tbljob.job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if	
		fa_DS_under_month(i) = selectid(sql2)
	next
	
	'fa_WI_over, 
	for i = 1 to 12  
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='WI' and tbljob.job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if
		fa_WI_over_month(i) = selectid(sql2)
	next
	
	'fa_WI_under
	for i = 1 to 12  
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='WI' and tbljob.job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if
		fa_WI_under_month(i) = selectid(sql2)
	next
    
	'fa_CF_over, 
	for i = 1 to 12  
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='CF' and tbljob.job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if
		fa_CF_over_month(i) = selectid(sql2)
	next
	
	'fa_CF_under, 
	for i = 1 to 12  
		sql2 = "SELECT DISTINCT  count(tbljob.job_id) as totaljob FROM tbljob_parts INNER JOIN " & _
			  "tbljob ON tbljob_parts.job_code = tbljob.job_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and tbljob_parts.jobp_partcode = '" & rs1("parts_code") & "' and tbljob.job_hq_category_code='CF' and tbljob.job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
		end if
	fa_CF_under_month(i) = selectid(sql2)
	next
	
	fa_MD_over=0
	fa_MD_under=0
	fa_DS_over=0
	fa_DS_under=0
	fa_WI_over=0
	fa_WI_under=0
	fa_CF_over=0
	fa_CF_under=0
	
	for i = 1 to 12  
	    fa_MD_over = fa_MD_over + chknumber(fa_MD_over_month(i))
		fa_MD_under = fa_MD_under + chknumber(fa_MD_under_month(i))
		fa_DS_over = fa_DS_over + chknumber(fa_DS_over_month(i))
		fa_DS_under = fa_DS_under + chknumber(fa_DS_under_month(i))
		fa_WI_over = fa_WI_over + chknumber(fa_WI_over_month(i))
		fa_WI_under = fa_WI_under + chknumber(fa_WI_under_month(i))
		fa_CF_over = fa_CF_over + chknumber(fa_CF_over_month(i))
		fa_CF_under = fa_CF_under + chknumber(fa_CF_under_month(i))
	next
		
	sql4 = "Update tblrpr_farParts set " & _
	"fa_month1_over = " & fa_month1_over & ", fa_month1_under=" & fa_month1_under & ", " & _
	"fa_month2_over=" & fa_month2_over & ", fa_month2_under=" & fa_month2_under & ", " & _
	"fa_month3_over=" & fa_month3_over & ", fa_month3_under=" & fa_month3_under & ", " & _
	"fa_month4_over=" & fa_month4_over & ", fa_month4_under=" & fa_month4_under & ", " & _
	"fa_month5_over=" & fa_month5_over & ", fa_month5_under=" & fa_month5_under & ", " & _
	"fa_month6_over=" & fa_month6_over & ", fa_month6_under=" & fa_month6_under & ", " & _
	"fa_month7_over=" & fa_month7_over & ", fa_month7_under=" & fa_month7_under & ", " & _
	"fa_month8_over=" & fa_month8_over & ", fa_month8_under=" & fa_month8_under & ", " & _
	"fa_month9_over=" & fa_month9_over & ", fa_month9_under=" & fa_month9_under & ", " & _
	"fa_month10_over=" & fa_month10_over & ", fa_month10_under=" & fa_month10_under & ", " & _
	"fa_month11_over=" & fa_month11_over & ", fa_month11_under=" & fa_month11_under & ", " & _
	"fa_month12_over=" & fa_month12_over & ", fa_month12_under=" & fa_month12_under & ", " & _
	"fa_month_total_over=" & fa_month_total_over & ", fa_month_total_under=" & fa_month_total_under & ", " & _
	"fa_MD_over=" & fa_MD_over & ", fa_MD_under=" & fa_MD_under & ", " & _
	"fa_DS_over=" & fa_DS_over & ", fa_DS_under=" & fa_DS_under & ", " & _
	"fa_WI_over=" & fa_WI_over & ", fa_WI_under=" & fa_WI_under & ", " & _
	"fa_CF_over=" & fa_CF_over & ", fa_CF_under=" & fa_CF_under & ", " & _
	"fa_done='Y' where id = " & rs1("id") 
	
	response.write sql4 & "<br><br>"
	CUD(sql4)
	stillgotRecord = "Yes"
	

	rs1.movenext
	wend
	rs1.close
	
    job_tech_model = replace(job_tech_model, "'", "")
	
	if stillgotRecord = "No" then 
	
	   response.write "Process Completed, <a href='rm_rpt_farmonth_spareparts.asp?orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & TotalSales & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type & "&loginerr=Job has been updated.#articletitle'>View Result</a>"
	   response.end
	
	else
	
	   url = "action_report_farmonthparts_return.asp?type=rpt_farparts_year&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & TotalSales & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type & "&stepno=2"	
	
	end if
	
	
    		 		 		 			 
'----------------------------------------------------------------------------------------------------    

		 
End Select
Response.Clear
Response.Redirect(url)
%>

