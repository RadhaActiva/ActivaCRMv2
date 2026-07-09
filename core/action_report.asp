<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<%
act = Request("type")

dim jobmonth_list(13), jobyear_list(13)
dim fa_MD_over_month(13), fa_MD_under_month(13), fa_DS_over_month(13), fa_DS_under_month(13), fa_WI_over_month(13), fa_WI_under_month(13), fa_CF_over_month(13), fa_CF_under_month(13)


Function ChkZero(str)	  
	 if str = "" or isnull(str) or len(str) = 0 then 
	    ChkZero = "0"
	 else
	    ChkZero = str
	 end if
End Function

'----------------------------------------------------------------------------------------------------    
	
Select Case act
 
'----------------------------------------------------------------------------------------------------    
  Case "rpt_farmonth_reset"
  
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_farmonth"
	CUD(sql)
	
	if request("job_tech_type") = "All" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_status='Y'"
	CUD(sql)
	
	elseif request("job_tech_type") = "CF" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='CF' and fr_status='Y'"
	CUD(sql)
	
	elseif request("job_tech_type") = "WH" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='WH' and fr_status='Y' "
	CUD(sql)
	
	end if
	
    url = "rm_rpt_farmonth.asp?type=showresult&jobmonth=" & request("jobmonth") & "&jobyear=" & request("jobyear") & "&job_tech_type=" & request("job_tech_type") & "&loginerr=Report has been reset.#articletitle" 	

	   
'----------------------------------------------------------------------------------------------------    
  Case "rpt_farmonth"
  
	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	updatemonth = request("updatemonth")
	
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
	
    ''' Loop tblrpr_farmonth
	sql1 = "SELECT id, faulth_code, faulth_desc, fa_month1_over, fa_month1_under, fa_month2_over, fa_month2_under, fa_month3_over, fa_month3_under, " & _
		   "fa_month4_over, fa_month4_under, fa_month5_over, fa_month5_under, fa_month6_over, fa_month6_under, fa_month7_over, fa_month7_under,  " & _
		   "fa_month8_over, fa_month8_under, fa_month9_over, fa_month9_under, fa_month10_over, fa_month10_under, fa_month11_over, fa_month11_under,  " & _
		   "fa_month12_over, fa_month12_under, fa_month_total_over, fa_month_total_under, fa_MD1_over, fa_MD1_under, fa_MD2_over, fa_MD2_under, fa_MD3_over,  " & _
		   "fa_MD3_under, fa_MD4_over, fa_MD4_under, fa_MD5_over, fa_MD5_under, fa_MD6_over, fa_MD6_under, fa_MD7_over, fa_MD7_under, fa_MD8_over,  " & _
		   "fa_MD8_under, fa_MD9_over, fa_MD9_under, fa_MD10_over, fa_MD10_under, fa_MD11_over, fa_MD11_under, fa_MD12_over, fa_MD12_under, fa_DS1_over,  " & _
		   "fa_DS1_under, fa_DS2_over, fa_DS2_under, fa_DS3_over, fa_DS3_under, fa_DS4_over, fa_DS4_under, fa_DS5_over, fa_DS5_under, fa_DS6_over,  " & _
		   "fa_DS6_under, fa_DS7_over, fa_DS7_under, fa_DS8_over, fa_DS8_under, fa_DS9_over, fa_DS9_under, fa_DS10_over, fa_DS10_under, fa_DS11_over, " & _ 
		   "fa_DS11_under, fa_DS12_over, fa_DS12_under, fa_WI1_over, fa_WI1_under, fa_WI2_over, fa_WI2_under, fa_WI3_over, fa_WI3_under, fa_WI4_over,  " & _
		   "fa_WI4_under, fa_WI5_over, fa_WI5_under, fa_WI6_over, fa_WI6_under, fa_WI7_over, fa_WI7_under, fa_WI8_over, fa_WI8_under, fa_WI9_over,  " & _
		   "fa_WI9_under, fa_WI10_over, fa_WI10_under, fa_WI11_over, fa_WI11_under, fa_WI12_over, fa_WI12_under, fa_CF1_over, fa_CF1_under, fa_CF2_over,  " & _
		   "fa_CF2_under, fa_CF3_over, fa_CF3_under, fa_CF4_over, fa_CF4_under, fa_CF5_over, fa_CF5_under, fa_CF6_over, fa_CF6_under, fa_CF7_over,  " & _
		   "fa_CF7_under, fa_CF8_over, fa_CF8_under, fa_CF9_over, fa_CF9_under, fa_CF10_over, fa_CF10_under, fa_CF11_over, fa_CF11_under, fa_CF12_over,  " & _
		   "fa_CF12_under FROM tblrpr_farmonth order by id "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

	'fa_month1_over
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	
	if updatemonth = "jobmonth1" then 
	   fa_month1_over = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_month2_over = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_month3_over = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_month4_over = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_month5_over = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_month6_over = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_month7_over = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_month8_over = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_month9_over = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_month10_over = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_month11_over = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_month12_over = selectid(sql2)    
	end if                  
	
	'fa_month1_under
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	
	if updatemonth = "jobmonth1" then 
	   fa_month1_under = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_month2_under = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_month3_under = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_month4_under = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_month5_under = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_month6_under = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_month7_under = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_month8_under = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_month9_under = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_month10_under = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_month11_under = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_month12_under = selectid(sql2)    
	end if 
	
	'fa_MD_over_month
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		
	if updatemonth = "jobmonth1" then 
	   fa_MD_over_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_MD_over_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_MD_over_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_MD_over_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_MD_over_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_MD_over_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_MD_over_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_MD_over_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_MD_over_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_MD_over_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_MD_over_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_MD_over_month12 = selectid(sql2)    
	end if 
	
	'fa_MD_under_month
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if

	if updatemonth = "jobmonth1" then 
	   fa_MD_under_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_MD_under_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_MD_under_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_MD_under_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_MD_under_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_MD_under_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_MD_under_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_MD_under_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_MD_under_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_MD_under_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_MD_under_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_MD_under_month12 = selectid(sql2)    
	end if 
	
	'fa_DS_over,
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
		end if

	if updatemonth = "jobmonth1" then 
	   fa_DS_over_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_DS_over_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_DS_over_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_DS_over_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_DS_over_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_DS_over_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_DS_over_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_DS_over_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_DS_over_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_DS_over_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_DS_over_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_DS_over_month12 = selectid(sql2)    
	end if 
		
	'fa_DS_under, 
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		
	if updatemonth = "jobmonth1" then 
	   fa_DS_under_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_DS_under_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_DS_under_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_DS_under_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_DS_under_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_DS_under_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_DS_under_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_DS_under_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_DS_under_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_DS_under_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_DS_under_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_DS_under_month12 = selectid(sql2)    
	end if 

	'fa_WI_over, 
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		
	if updatemonth = "jobmonth1" then 
	   fa_WI_over_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_WI_over_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_WI_over_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_WI_over_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_WI_over_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_WI_over_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_WI_over_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_WI_over_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_WI_over_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_WI_over_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_WI_over_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_WI_over_month12 = selectid(sql2)    
	end if 
		
	'fa_WI_under
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		
	if updatemonth = "jobmonth1" then 
	   fa_WI_under_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_WI_under_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_WI_under_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_WI_under_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_WI_under_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_WI_under_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_WI_under_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_WI_under_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_WI_under_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_WI_under_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_WI_under_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_WI_under_month12 = selectid(sql2)    
	end if 
    
	'fa_CF_over, 	
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			   "where tbljob.job_id is not null " & _
			   "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
			   "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if

	if updatemonth = "jobmonth1" then 
	   fa_CF_over_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_CF_over_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_CF_over_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_CF_over_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_CF_over_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_CF_over_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_CF_over_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_CF_over_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_CF_over_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_CF_over_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_CF_over_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_CF_over_month12 = selectid(sql2)    
	end if 
		
	'fa_CF_under,
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
	end if
	
	if updatemonth = "jobmonth1" then 
	   fa_CF_under_month1 = selectid(sql2)
	elseif updatemonth = "jobmonth2" then   
	   fa_CF_under_month2 = selectid(sql2) 
	elseif updatemonth = "jobmonth3" then   
	   fa_CF_under_month3 = selectid(sql2) 
	elseif updatemonth = "jobmonth4" then   
	   fa_CF_under_month4 = selectid(sql2) 
	elseif updatemonth = "jobmonth5" then   
	   fa_CF_under_month5 = selectid(sql2) 
	elseif updatemonth = "jobmonth6" then   
	   fa_CF_under_month6 = selectid(sql2) 
	elseif updatemonth = "jobmonth7" then   
	   fa_CF_under_month7 = selectid(sql2) 
	elseif updatemonth = "jobmonth8" then   
	   fa_CF_under_month8 = selectid(sql2) 
	elseif updatemonth = "jobmonth9" then   
	   fa_CF_under_month9 = selectid(sql2) 
	elseif updatemonth = "jobmonth10" then   
	   fa_CF_under_month10 = selectid(sql2) 
	elseif updatemonth = "jobmonth11" then   
	   fa_CF_under_month11 = selectid(sql2)
	elseif updatemonth = "jobmonth12" then   
	   fa_CF_under_month12 = selectid(sql2)    
	end if 	
	
		
	sql4 = "Update tblrpr_farmonth set " 
	
	if updatemonth = "jobmonth1" then 
	sql4 = sql4 & "fa_month1_over = " & fa_month1_over & ", fa_month1_under=" & fa_month1_under & ", " 
	elseif updatemonth = "jobmonth2" then 
	sql4 = sql4 & "fa_month2_over = " & fa_month2_over & ", fa_month2_under=" & fa_month2_under & ", " 
	elseif updatemonth = "jobmonth3" then 
	sql4 = sql4 & "fa_month3_over = " & fa_month3_over & ", fa_month3_under=" & fa_month3_under & ", " 
	elseif updatemonth = "jobmonth4" then 
	sql4 = sql4 & "fa_month4_over = " & fa_month4_over & ", fa_month4_under=" & fa_month4_under & ", " 
	elseif updatemonth = "jobmonth5" then 
	sql4 = sql4 & "fa_month5_over = " & fa_month5_over & ", fa_month5_under=" & fa_month5_under & ", " 
	elseif updatemonth = "jobmonth6" then 
	sql4 = sql4 & "fa_month6_over = " & fa_month6_over & ", fa_month6_under=" & fa_month6_under & ", " 
	elseif updatemonth = "jobmonth7" then 
	sql4 = sql4 & "fa_month7_over = " & fa_month7_over & ", fa_month7_under=" & fa_month7_under & ", " 
	elseif updatemonth = "jobmonth8" then 
	sql4 = sql4 & "fa_month8_over = " & fa_month8_over & ", fa_month8_under=" & fa_month8_under & ", " 
	elseif updatemonth = "jobmonth9" then 
	sql4 = sql4 & "fa_month9_over = " & fa_month9_over & ", fa_month9_under=" & fa_month9_under & ", " 
	elseif updatemonth = "jobmonth10" then 
	sql4 = sql4 & "fa_month10_over = " & fa_month10_over & ", fa_month10_under=" & fa_month10_under & ", " 
	elseif updatemonth = "jobmonth11" then 
	sql4 = sql4 & "fa_month11_over = " & fa_month11_over & ", fa_month11_under=" & fa_month11_under & ", " 
	elseif updatemonth = "jobmonth12" then 
	sql4 = sql4 & "fa_month12_over = " & fa_month12_over & ", fa_month12_under=" & fa_month12_under & ", " 
	end if

	if updatemonth = "jobmonth1" then 
	sql4 = sql4 & _
	"fa_MD1_over=" & fa_MD_over_month1 & ", fa_MD1_under=" & fa_MD_under_month1 & ", " & _
	"fa_DS1_over=" & fa_DS_over_month1 & ", fa_DS1_under=" & fa_DS_under_month1 & ", " & _
	"fa_WI1_over=" & fa_WI_over_month1 & ", fa_WI1_under=" & fa_WI_under_month1 & ", " & _
	"fa_CF1_over=" & fa_CF_over_month1 & ", fa_CF1_under=" & fa_CF_under_month1 & " " 
	
	elseif updatemonth = "jobmonth2" then 
	sql4 = sql4 & _
	"fa_MD2_over=" & fa_MD_over_month2 & ", fa_MD2_under=" & fa_MD_under_month2 & ", " & _
	"fa_DS2_over=" & fa_DS_over_month2 & ", fa_DS2_under=" & fa_DS_under_month2 & ", " & _
	"fa_WI2_over=" & fa_WI_over_month2 & ", fa_WI2_under=" & fa_WI_under_month2 & ", " & _
	"fa_CF2_over=" & fa_CF_over_month2 & ", fa_CF2_under=" & fa_CF_under_month2 & " " 
	
	elseif updatemonth = "jobmonth3" then 
	sql4 = sql4 & _
	"fa_MD3_over=" & fa_MD_over_month3 & ", fa_MD3_under=" & fa_MD_under_month3 & ", " & _
	"fa_DS3_over=" & fa_DS_over_month3 & ", fa_DS3_under=" & fa_DS_under_month3 & ", " & _
	"fa_WI3_over=" & fa_WI_over_month3 & ", fa_WI3_under=" & fa_WI_under_month3 & ", " & _
	"fa_CF3_over=" & fa_CF_over_month3 & ", fa_CF3_under=" & fa_CF_under_month3 & " " 

	elseif updatemonth = "jobmonth4" then 
	sql4 = sql4 & _
	"fa_MD4_over=" & fa_MD_over_month4 & ", fa_MD4_under=" & fa_MD_under_month4 & ", " & _
	"fa_DS4_over=" & fa_DS_over_month4 & ", fa_DS4_under=" & fa_DS_under_month4 & ", " & _
	"fa_WI4_over=" & fa_WI_over_month4 & ", fa_WI4_under=" & fa_WI_under_month4 & ", " & _
	"fa_CF4_over=" & fa_CF_over_month4 & ", fa_CF4_under=" & fa_CF_under_month4 & " " 
	
	elseif updatemonth = "jobmonth5" then 
	sql4 = sql4 & _
	"fa_MD5_over=" & fa_MD_over_month5 & ", fa_MD5_under=" & fa_MD_under_month5 & ", " & _
	"fa_DS5_over=" & fa_DS_over_month5 & ", fa_DS5_under=" & fa_DS_under_month5 & ", " & _
	"fa_WI5_over=" & fa_WI_over_month5 & ", fa_WI5_under=" & fa_WI_under_month5 & ", " & _
	"fa_CF5_over=" & fa_CF_over_month5 & ", fa_CF5_under=" & fa_CF_under_month5 & " "

	elseif updatemonth = "jobmonth6" then 
	sql4 = sql4 & _
	"fa_MD6_over=" & fa_MD_over_month6 & ", fa_MD6_under=" & fa_MD_under_month6 & ", " & _
	"fa_DS6_over=" & fa_DS_over_month6 & ", fa_DS6_under=" & fa_DS_under_month6 & ", " & _
	"fa_WI6_over=" & fa_WI_over_month6 & ", fa_WI6_under=" & fa_WI_under_month6 & ", " & _
	"fa_CF6_over=" & fa_CF_over_month6 & ", fa_CF6_under=" & fa_CF_under_month6 & " "
	
	elseif updatemonth = "jobmonth7" then 
	sql4 = sql4 & _
	"fa_MD7_over=" & fa_MD_over_month7 & ", fa_MD7_under=" & fa_MD_under_month7 & ", " & _
	"fa_DS7_over=" & fa_DS_over_month7 & ", fa_DS7_under=" & fa_DS_under_month7 & ", " & _
	"fa_WI7_over=" & fa_WI_over_month7 & ", fa_WI7_under=" & fa_WI_under_month7 & ", " & _
	"fa_CF7_over=" & fa_CF_over_month7 & ", fa_CF7_under=" & fa_CF_under_month7 & " "

	elseif updatemonth = "jobmonth8" then 
	sql4 = sql4 & _
	"fa_MD8_over=" & fa_MD_over_month8 & ", fa_MD8_under=" & fa_MD_under_month8 & ", " & _
	"fa_DS8_over=" & fa_DS_over_month8 & ", fa_DS8_under=" & fa_DS_under_month8 & ", " & _
	"fa_WI8_over=" & fa_WI_over_month8 & ", fa_WI8_under=" & fa_WI_under_month8 & ", " & _
	"fa_CF8_over=" & fa_CF_over_month8 & ", fa_CF8_under=" & fa_CF_under_month8 & " "
	
	elseif updatemonth = "jobmonth9" then 
	sql4 = sql4 & _
	"fa_MD9_over=" & fa_MD_over_month9 & ", fa_MD9_under=" & fa_MD_under_month9 & ", " & _
	"fa_DS9_over=" & fa_DS_over_month9 & ", fa_DS9_under=" & fa_DS_under_month9 & ", " & _
	"fa_WI9_over=" & fa_WI_over_month9 & ", fa_WI9_under=" & fa_WI_under_month9 & ", " & _
	"fa_CF9_over=" & fa_CF_over_month9 & ", fa_CF9_under=" & fa_CF_under_month9 & " "

	elseif updatemonth = "jobmonth10" then 
	sql4 = sql4 & _
	"fa_MD10_over=" & fa_MD_over_month10 & ", fa_MD10_under=" & fa_MD_under_month10 & ", " & _
	"fa_DS10_over=" & fa_DS_over_month10 & ", fa_DS10_under=" & fa_DS_under_month10 & ", " & _
	"fa_WI10_over=" & fa_WI_over_month10 & ", fa_WI10_under=" & fa_WI_under_month10 & ", " & _
	"fa_CF10_over=" & fa_CF_over_month10 & ", fa_CF10_under=" & fa_CF_under_month10 & " "
	
	elseif updatemonth = "jobmonth11" then 
	sql4 = sql4 & _
	"fa_MD11_over=" & fa_MD_over_month11 & ", fa_MD11_under=" & fa_MD_under_month11 & ", " & _
	"fa_DS11_over=" & fa_DS_over_month11 & ", fa_DS11_under=" & fa_DS_under_month11 & ", " & _
	"fa_WI11_over=" & fa_WI_over_month11 & ", fa_WI11_under=" & fa_WI_under_month11 & ", " & _
	"fa_CF11_over=" & fa_CF_over_month11 & ", fa_CF11_under=" & fa_CF_under_month11 & " "
	
	elseif updatemonth = "jobmonth12" then 
	sql4 = sql4 & _
	"fa_MD12_over=" & fa_MD_over_month12 & ", fa_MD12_under=" & fa_MD_under_month12 & ", " & _
	"fa_DS12_over=" & fa_DS_over_month12 & ", fa_DS12_under=" & fa_DS_under_month12 & ", " & _
	"fa_WI12_over=" & fa_WI_over_month12 & ", fa_WI12_under=" & fa_WI_under_month12 & ", " & _
	"fa_CF12_over=" & fa_CF_over_month12 & ", fa_CF12_under=" & fa_CF_under_month12 & " "
	end if
		
	sql4 = sql4 & _
	" where id = " & rs1("id") 
	'response.write sql4
	'response.End()
	CUD(sql4)
	
	sql4 = "update tblrpr_farmonth set " & _
	       "fa_month_total_over = fa_month1_over + fa_month2_over + fa_month3_over + fa_month4_over + fa_month5_over + fa_month6_over + fa_month7_over + fa_month8_over + fa_month9_over + fa_month10_over + fa_month11_over + fa_month12_over, " & _
		   "fa_month_total_under = fa_month1_under + fa_month2_under + fa_month3_under + fa_month4_under + fa_month5_under + fa_month6_under + fa_month7_under + fa_month8_under + fa_month9_under + fa_month10_under + fa_month11_under + fa_month12_under, " & _  
		   "fa_MD_over = fa_MD1_over + fa_MD2_over + fa_MD3_over + fa_MD4_over + fa_MD5_over + fa_MD6_over + fa_MD7_over + fa_MD8_over + fa_MD9_over + fa_MD10_over + fa_MD11_over + fa_MD12_over, " & _
		   "fa_MD_under = fa_MD1_under + fa_MD2_under + fa_MD3_under + fa_MD4_under + fa_MD5_under + fa_MD6_under + fa_MD7_under + fa_MD8_under + fa_MD9_under + fa_MD10_under + fa_MD11_under + fa_MD12_under, " & _	
		   "fa_DS_over = fa_DS1_over + fa_DS2_over + fa_DS3_over + fa_DS4_over + fa_DS5_over + fa_DS6_over + fa_DS7_over + fa_DS8_over + fa_DS9_over + fa_DS10_over + fa_DS11_over + fa_DS12_over, " & _
		   "fa_DS_under = fa_DS1_under + fa_DS2_under + fa_DS3_under + fa_DS4_under + fa_DS5_under + fa_DS6_under + fa_DS7_under + fa_DS8_under + fa_DS9_under + fa_DS10_under + fa_DS11_under + fa_DS12_under, " & _	
		   "fa_WI_over = fa_WI1_over + fa_WI2_over + fa_WI3_over + fa_WI4_over + fa_WI5_over + fa_WI6_over + fa_WI7_over + fa_WI8_over + fa_WI9_over + fa_WI10_over + fa_WI11_over + fa_WI12_over, " & _
		   "fa_WI_under = fa_WI1_under + fa_WI2_under + fa_WI3_under + fa_WI4_under + fa_WI5_under + fa_WI6_under + fa_WI7_under + fa_WI8_under + fa_WI9_under + fa_WI10_under + fa_WI11_under + fa_WI12_under, " & _	
		   "fa_CF_over = fa_CF1_over + fa_CF2_over + fa_CF3_over + fa_CF4_over + fa_CF5_over + fa_CF6_over + fa_CF7_over + fa_CF8_over + fa_CF9_over + fa_CF10_over + fa_CF11_over + fa_CF12_over, " & _
		   "fa_CF_under = fa_CF1_under + fa_CF2_under + fa_CF3_under + fa_CF4_under + fa_CF5_under + fa_CF6_under + fa_CF7_under + fa_CF8_under + fa_CF9_under + fa_CF10_under + fa_CF11_under + fa_CF12_under " & _
		   " where id = " & rs1("id") 
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
    job_tech_model = replace(job_tech_model, "'", "")
	
	url = "rm_rpt_farmonth.asp?type=showresult&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & request("TotalSales") & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type & "&updatemonth=" & updatemonth & "&loginerr=Job has been updated.#articletitle" 	


'----------------------------------------------------------------------------------------------------    
  Case "rpt_farmonth_year"
  
	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	
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
	
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_farmonth"
	CUD(sql)
	
	if request("job_tech_type") = "All" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_status='Y'"
	CUD(sql)
	
	elseif request("job_tech_type") = "CF" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='CF' and fr_status='Y'"
	CUD(sql)
	
	elseif request("job_tech_type") = "WH" then 
	sql = "INSERT INTO tblrpr_farmonth (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='WH' and fr_status='Y' "
	CUD(sql)
	
	end if
	
	
    ''' Loop tblrpr_farmonth
	sql1 = "SELECT id, faulth_code, faulth_desc, fa_month1_over, fa_month1_under, fa_month2_over, fa_month2_under, fa_month3_over, fa_month3_under, " & _
		   "fa_month_total_over, fa_month_total_under, fa_MD_over, fa_MD_under, fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
		   "FROM tblrpr_farmonth order by id "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

	'fa_month1_over
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month1_over = selectid(sql2)
	
	'fa_month1_under
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth1 & "' and year(tbljob.job_posteddate) = '" & jobyear1 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month1_under = selectid(sql2)
	
	'fa_month2_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth2 & "' and year(tbljob.job_posteddate) = '" & jobyear2 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month2_Over = selectid(sql2)
	
	'fa_month2_under,
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth2 & "' and year(tbljob.job_posteddate) = '" & jobyear2 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month2_Under = selectid(sql2)
	 
	'fa_month3_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth3 & "' and year(tbljob.job_posteddate) = '" & jobyear3 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	
	fa_month3_over = selectid(sql2)
	
	'fa_month3_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth3 & "' and year(tbljob.job_posteddate) = '" & jobyear3 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month3_under = selectid(sql2)
	
	'fa_month4_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth4 & "' and year(tbljob.job_posteddate) = '" & jobyear4 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month4_over = selectid(sql2)
	
	'fa_month4_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth4 & "' and year(tbljob.job_posteddate) = '" & jobyear4 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month4_under = selectid(sql2)

	'fa_month5_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth5 & "' and year(tbljob.job_posteddate) = '" & jobyear5 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month5_over = selectid(sql2)
	
	'fa_month5_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth5 & "' and year(tbljob.job_posteddate) = '" & jobyear5 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month5_under = selectid(sql2)	
	
	'fa_month6_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth6 & "' and year(tbljob.job_posteddate) = '" & jobyear6 & "' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month6_over = selectid(sql2)
	
	'fa_month6_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth6 & "' and year(tbljob.job_posteddate) = '" & jobyear6 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month6_under = selectid(sql2)	
	
	'fa_month7_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth7 & "' and year(tbljob.job_posteddate) = '" & jobyear7 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month7_over = selectid(sql2)
	
	'fa_month7_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth7 & "' and year(tbljob.job_posteddate) = '" & jobyear7 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month7_under = selectid(sql2)	
	
	'fa_month8_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth8 & "' and year(tbljob.job_posteddate) = '" & jobyear8 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month8_over = selectid(sql2)
	
	'fa_month8_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth8 & "' and year(tbljob.job_posteddate) = '" & jobyear8 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month8_under = selectid(sql2)		
	
	'fa_month9_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth9 & "' and year(tbljob.job_posteddate) = '" & jobyear9 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month9_over = selectid(sql2)
	
	'fa_month9_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth9 & "' and year(tbljob.job_posteddate) = '" & jobyear9 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month9_under = selectid(sql2)	
	
	'fa_month10_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth10 & "' and year(tbljob.job_posteddate) = '" & jobyear10 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month10_over = selectid(sql2)
	
	'fa_month10_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth10 & "' and year(tbljob.job_posteddate) = '" & jobyear10 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month10_under = selectid(sql2)
	
	'fa_month11_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth11 & "' and year(tbljob.job_posteddate) = '" & jobyear11 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month11_over = selectid(sql2)
	
	'fa_month11_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth11 & "' and year(tbljob.job_posteddate) = '" & jobyear11 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month11_under = selectid(sql2)
	
	'fa_month12_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth12 & "' and year(tbljob.job_posteddate) = '" & jobyear12 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if	
	fa_month12_over = selectid(sql2)
	
	'fa_month12_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth12 & "' and year(tbljob.job_posteddate) = '" & jobyear12 & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month12_under = selectid(sql2)
	
	'fa_month_total_over, 
	fa_month_total_over = ChkNumber(fa_month1_over) + ChkNumber(fa_month2_over) + ChkNumber(fa_month3_over) + ChkNumber(fa_month4_over) + ChkNumber(fa_month5_over) + ChkNumber(fa_month6_over) + ChkNumber(fa_month7_over) + ChkNumber(fa_month8_over) + ChkNumber(fa_month9_over) + ChkNumber(fa_month10_over) + ChkNumber(fa_month11_over) + ChkNumber(fa_month12_over)
	
	'fa_month_total_under, 
	fa_month_total_under = ChkNumber(fa_month1_under) + ChkNumber(fa_month2_under) + ChkNumber(fa_month3_under) + ChkNumber(fa_month4_under) + ChkNumber(fa_month5_under) + ChkNumber(fa_month6_under) + ChkNumber(fa_month7_under) + ChkNumber(fa_month8_under) + ChkNumber(fa_month9_under) + ChkNumber(fa_month10_under) + ChkNumber(fa_month11_under) + ChkNumber(fa_month12_under) 
	
	'fa_MD_over_month
	for i = 1 to 12 
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_MD_over_month(i) = selectid(sql2)
	next
	
	'fa_MD_under_month
	for i = 1 to 12 
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_MD_under_month(i) = selectid(sql2)
	next
	
	'fa_DS_over,
	for i = 1 to 12  
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
		end if
		fa_DS_over_month(i) = selectid(sql2)
	next
	
	'fa_DS_under, 
	for i = 1 to 12  
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_DS_under_month(i) = selectid(sql2)
	next
	
	'fa_WI_over, 
	for i = 1 to 12  
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_WI_over_month(i) = selectid(sql2)
	next
	
	'fa_WI_under
	for i = 1 to 12  
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			  "where tbljob.job_id is not null " & _
			  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_WI_under_month(i) = selectid(sql2)
	next
    
	'fa_CF_over, 
	for i = 1 to 12  
		sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
			   "where tbljob.job_id is not null " & _
			   "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
			   "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
		if job_tech_type <> "" then 
		   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
		end if
		if job_tech_model <> "" then 
		   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
		end if
		fa_CF_over_month(i) = selectid(sql2)
	next
	
	'fa_CF_under, 
	for i = 1 to 12  
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and  month(tbljob.job_posteddate) = '" & jobmonth_list(i) & "' and year(tbljob.job_posteddate) = '" & jobyear_list(i) & "' " & _
		  "and job_tech_faulty_reason='" & rs1("faulth_desc") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
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
		
	sql4 = "Update tblrpr_farmonth set " & _
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
	"fa_CF_over=" & fa_CF_over & ", fa_CF_under=" & fa_CF_under & " " & _
	"where id = " & rs1("id") 
	
	'response.write sql4
	'response.End()
	
	
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
    job_tech_model = replace(job_tech_model, "'", "")
	
	url = "rm_rpt_farmonth_year.asp?type=showresult&orderby=" & orderby & "&ordertype=" & ordertype & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&TotalSales=" & request("TotalSales") & "&job_tech_model=" & job_tech_model & "&job_tech_type=" & job_tech_type & "&loginerr=Job has been updated.#articletitle" 	
    		 		 		 			 
'----------------------------------------------------------------------------------------------------    
  Case "rpt_fararea"
  
	job_tech_type = request("job_tech_type")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	
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
	
	if request("job_tech_model") <> "" then
	   job_tech_model = replace(request("job_tech_model"), " ", "")
	   arrjob_tech_model = split(job_tech_model,",")
	   job_tech_model = replace(job_tech_model, ",", "','")
	   listjob_tech_model = listjob_tech_model & job_tech_model
	else
	   listjob_tech_model = ""
	   arrjob_tech_model = split("0,0",",")
	end if
	
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_fararea"
	CUD(sql)
	sql = "INSERT INTO tblrpr_fararea (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_status='Y'"
	CUD(sql)
	
	
	dim fa_state_over(15), fa_state_under(15)
	
    ''' Loop tblrpr_farmonth
	sql1 = "SELECT id, faulth_code, faulth_desc, fa_state1_over, fa_state1_under, fa_state2_over, fa_state2_under, fa_state3_over, fa_state3_under, " & _
			"fa_state_total_over, fa_state_total_under, fa_MD_over, fa_MD_under, fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
			"FROM tblrpr_fararea order by id "
	set rs1 = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

		    i = 1
		    sql2 = "select state_id, state_name from tblstate order by state_cnty_id"
			rs2.ActiveConnection = strconnect
			rs2.Source = sql2
			rs2.CursorLocation  = 3
			rs2.Open
			while not rs2.eof 
			        'fa_state_over
					sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
							"where tbljob.job_id is not null " & _
							"and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
							"and tbljob.job_cust_state_id = '" & rs2("state_id") & "' " & _
							"and tbljob.job_status='Posted'" & _
							"and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Over' "
							if job_tech_type <> "" then 
							sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
							end if
							if job_tech_model <> "" then 
							sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
							end if
					fa_state_over(i) = selectid(sql2)
					
					'fa_state_Under
					sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
							"where tbljob.job_id is not null " & _
							"and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
							"and tbljob.job_cust_state_id = '" & rs2("state_id") & "' " & _
							"and tbljob.job_status='Posted'" & _
							"and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_actual_wrty_status='Under' "
							if job_tech_type <> "" then 
							sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
							end if
							if job_tech_model <> "" then 
							sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model & "') "
							end if
					fa_state_under(i) = selectid(sql2)
					
					i = i + 1
			rs2.movenext
			wend
			rs2.close
		
			
	fa_state_total_over = 0
	fa_state_total_under = 0
	for i = 1 to 15 
		if not isnull(fa_state_over(i)) then
		   fa_state_total_over = fa_state_total_over + cint(fa_state_over(i))
		end if
		
		if not isnull(fa_state_under(i)) then
		   fa_state_total_under = fa_state_total_under + cint(fa_state_under(i))
		end if
		
		response.write i & ": " & fa_state_under(i) & "<br>"
		
	next
	
	'fa_MD_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Over' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_MD_over = selectid(sql2)
	
	'fa_MD_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Under' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_MD_under = selectid(sql2)
	
	'fa_DS_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Over' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_DS_over = selectid(sql2)
	
	'fa_DS_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Under' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_DS_under = selectid(sql2)
	
	'fa_WI_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Over' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_WI_over = selectid(sql2)
	
	'fa_WI_under
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Under' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_WI_under = selectid(sql2)
    
	'fa_CF_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Over' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_CF_over = selectid(sql2)
	
	'fa_CF_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status='Posted' " & _
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Under' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_CF_under = selectid(sql2)
	
	sql4 = "Update tblrpr_fararea set " & _
	"fa_state1_over = " & fa_state_over(1) & ", fa_state1_under=" & fa_state_under(1) & ", " & _
	"fa_state2_over=" & fa_state_over(2) & ", fa_state2_under=" & fa_state_under(2) & ", " & _
	"fa_state3_over=" & fa_state_over(3) & ", fa_state3_under=" & fa_state_under(3) & ", " & _
	"fa_state4_over=" & fa_state_over(4) & ", fa_state4_under=" & fa_state_under(4) & ", " & _
	"fa_state5_over=" & fa_state_over(5) & ", fa_state5_under=" & fa_state_under(5) & ", " & _
	"fa_state6_over=" & fa_state_over(6) & ", fa_state6_under=" & fa_state_under(6) & ", " & _
	"fa_state7_over=" & fa_state_over(7) & ", fa_state7_under=" & fa_state_under(7) & ", " & _
	"fa_state8_over=" & fa_state_over(8) & ", fa_state8_under=" & fa_state_under(8) & ", " & _
	"fa_state9_over=" & fa_state_over(9) & ", fa_state9_under=" & fa_state_under(9) & ", " & _
	"fa_state10_over=" & fa_state_over(10) & ", fa_state10_under=" & fa_state_under(10) & ", " & _
	"fa_state11_over=" & fa_state_over(11) & ", fa_state11_under=" & fa_state_under(11) & ", " & _
	"fa_state12_over=" & fa_state_over(12) & ", fa_state12_under=" & fa_state_under(12) & ", " & _
	"fa_state13_over=" & fa_state_over(13) & ", fa_state13_under=" & fa_state_under(13) & ", " & _
	"fa_state14_over=" & fa_state_over(14) & ", fa_state14_under=" & fa_state_under(14) & ", " & _
	"fa_state15_over=" & fa_state_over(15) & ", fa_state15_under=" & fa_state_under(15) & ", " & _
	"fa_state_total_over=" & fa_state_total_over & ", fa_state_total_under=" & fa_state_total_under & ", " & _
	"fa_MD_over=" & fa_MD_over & ", fa_MD_under=" & fa_MD_under & ", fa_DS_over=" & fa_DS_over & ", fa_DS_under=" & fa_DS_under & ", " & _
	"fa_WI_over=" & fa_WI_over & ", fa_WI_under=" & fa_WI_under & ", fa_CF_over=" & fa_CF_over & " " & _
	"where id = " & rs1("id") 
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	

	url = "rm_rpt_fararea.asp?type=showresult&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & _
	      "&TotalSales=" & TotalSales & "&job_date_to=" & job_date_to & "&job_tech_model=" & request("job_tech_model") & "&job_tech_type=" & job_tech_type & "&loginerr=Report has been updated.#articletitle" 	


'----------------------------------------------------------------------------------------------------    
  Case "rpt_farDateRange"
  
	job_tech_type = request("job_tech_type")
	Searchor_date = request("Searchor_date")
	orderby = request("orderby")
	ordertype = request("ordertype")
	TotalSales = request("TotalSales")
	
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
	
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_fardaterange"
	CUD(sql)
	
	if request("job_tech_type") = "All" then 
	sql = "INSERT INTO tblrpr_fardaterange (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_status='Y' "
	CUD(sql)
	
	elseif request("job_tech_type") = "CF" then 
	sql = "INSERT INTO tblrpr_fardaterange (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='CF' and fr_status='Y'"
	CUD(sql)
	
	elseif request("job_tech_type") = "WH" then 
	sql = "INSERT INTO tblrpr_fardaterange (faulth_code, faulth_desc)  Select fr_code, fr_description from tblfaultyreason where fr_type='WH' and fr_status='Y'"
	CUD(sql)
	
	end if
	
	
    ''' Loop tblrpr_fardaterange
	sql1 = "SELECT id, faulth_code, faulth_desc, fa_month_total_over, fa_month_total_under, fa_MD_over, fa_MD_under, fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
		   "FROM tblrpr_fardaterange order by id "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 
	
	'fa_month_total_over
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and tbljob.job_actual_wrty_status='Over' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month_total_over = selectid(sql2)
	
	'fa_month_total_under
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and tbljob.job_tech_faulty_code='" & rs1("faulth_code") & "' and tbljob.job_actual_wrty_status='Under' and tbljob.job_status='Posted' "
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and tbljob.job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_month_total_under = selectid(sql2)
	
	'fa_MD_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_MD_over = selectid(sql2)
	
	'fa_MD_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='MD' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_MD_under = selectid(sql2)
	
	
	'fa_DS_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_DS_over = selectid(sql2)
	
	'fa_DS_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='DS' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_DS_under = selectid(sql2)
	
	'fa_WI_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_WI_over = selectid(sql2)
	
	'fa_WI_under
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='WI' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_WI_under = selectid(sql2)
    
	'fa_CF_over, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Over' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_CF_over = selectid(sql2)
	
	'fa_CF_under, 
    sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _		  
		  "and job_tech_faulty_code='" & rs1("faulth_code") & "' and job_hq_category_code='CF' and job_actual_wrty_status='Under' and tbljob.job_status='Posted'"
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_tech_model in ( '" & job_tech_model& "') "
	end if
	fa_CF_under = selectid(sql2)
		
	sql4 = "Update tblrpr_fardaterange set " & _
	"fa_month_total_over=" & fa_month_total_over & ", fa_month_total_under=" & fa_month_total_under & ", " & _
	"fa_MD_over=" & fa_MD_over & ", fa_MD_under=" & fa_MD_under & ", " & _
	"fa_DS_over=" & fa_DS_over & ", fa_DS_under=" & fa_DS_under & ", " & _
	"fa_WI_over=" & fa_WI_over & ", fa_WI_under=" & fa_WI_under & ", " & _
	"fa_CF_over=" & fa_CF_over & ", fa_CF_under=" & fa_CF_under & " " & _
	"where id = " & rs1("id")
	CUD(sql4)

	fa_MD_over=0
	fa_DS_over=0
	fa_WI_over=0
	fa_CF_over=0
	rs1.movenext
	wend
	rs1.close

	url = "rm_rpt_fardaterange.asp?type=showresult&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&TotalSales=" & request("TotalSales") & "&job_tech_model=" & request("job_tech_model") & "&job_tech_type=" & job_tech_type & "&loginerr=Job has been updated.#articletitle" 	
  	
'----------------------------------------------------------------------------------------------------    
  Case "rpt_monthtech"
  
	orderby = request("orderby")
	ordertype = request("ordertype")
	job_actual_wrty_status = request("job_actual_wrty_status")
	
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
		
	if request("job_tech_code") <> "" then
	   job_tech_code = replace(request("job_tech_code"), " ", "")
	   arrjob_tech_code = split(job_tech_code,",")
	   job_tech_code = replace(job_tech_code, ",", "','")
	   
	   listjob_tech_code = listjob_tech_code & job_tech_code
	   
	else
	   listjob_tech_code = ""
	   arrjob_tech_code = split("0,0",",")
	end if
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_monthtech"
	CUD(sql)
	sql = "INSERT INTO tblrpr_monthtech (tech_code, tech_name)  Select tech_code, tech_name from tbltechnician where tech_code is not null and (tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC')"
	CUD(sql)
	
    ''' Loop tblrpr_monthtech
	sql1 = "SELECT id, tech_code, tech_name, pending_1to3d, pending_4to6d, pending_7above, done_1to3d, done_4to6d, done_7above, posted_qty " & _
	       "FROM tblrpr_monthtech where id is not null "
		   
	if job_tech_code <> "" then 
	   sql1 = sql1 & " and tech_code in ( '" & job_tech_code & "') "
	end if
		   
	sql1 = sql1 & " order by id "
	
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

	'pending_1to3d
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and DATEDIFF(day , tbljob.job_JS_receiveddate, tbljob.job_submitteddate) > -1 and DATEDIFF(day , tbljob.job_JS_receiveddate, tbljob.job_submitteddate) < 4"
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	pending_1to3d = selectid(sql2)
	
	'pending_4to6d
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and DATEDIFF(day, tbljob.job_JS_receiveddate, tbljob.job_submitteddate) > 3 and DATEDIFF(day, tbljob.job_JS_receiveddate, tbljob.job_submitteddate) < 7"
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	pending_4to6d = selectid(sql2)
	
	'pending_7above, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and DATEDIFF(day, tbljob.job_JS_receiveddate, tbljob.job_submitteddate) > 6 "
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	pending_7above = selectid(sql2)
	
	'done_1to3d,
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and DATEDIFF(day, tbljob.job_donedate, tbljob.job_JS_receiveddate) > -1 and DATEDIFF(day,  tbljob.job_donedate, tbljob.job_JS_receiveddate) < 4"
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	done_1to3d = selectid(sql2)
	 
	'done_4to6d, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and DATEDIFF(day, tbljob.job_donedate, tbljob.job_JS_receiveddate) > 3 and DATEDIFF(day, tbljob.job_donedate, tbljob.job_JS_receiveddate) < 7"
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	done_4to6d = selectid(sql2)
	
	'done_7above, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' " & _ 
		  "and DATEDIFF(day, tbljob.job_donedate, tbljob.job_JS_receiveddate) > 6 "
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	done_7above = selectid(sql2)
	
	'posted_qty, 
	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_posteddate >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_posteddate <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
		  "and tbljob.job_status<>'Cancel' " & _
		  "and tbljob.job_tech_code = '" & rs1("tech_code") & "' and tbljob.job_status='Posted' " 
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	posted_qty = selectid(sql2)
	
	
	sql4 = "Update tblrpr_monthtech set " & _
	"pending_1to3d = " & ChkNumber0(pending_1to3d) & ", pending_4to6d=" & ChkNumber0(pending_4to6d) & ", pending_7above=" & ChkNumber0(pending_7above) & ", done_1to3d=" & ChkNumber0(done_1to3d) & ", " & _
	"done_4to6d=" & ChkNumber0(done_4to6d) & ", done_7above=" & ChkNumber0(done_7above) & ", posted_qty=" & ChkNumber0(posted_qty) & " where id = " & rs1("id") 
	CUD(sql4)
	
	rs1.movenext
	wend
	rs1.close
	

	url = "rm_rpt_monthtech.asp?type=showresult&orderby=" & orderby & "&ordertype=" & ordertype & "&job_tech_code=" & request("job_tech_code") & _
	       "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & _
		   "&job_actual_wrty_status=" & job_actual_wrty_status & "&loginerr=Report has been updated.#articletitle" 	
  
'----------------------------------------------------------------------------------------------------    
case "monthtechIncentive" 'this is meant for over-warranty jobs,ignore installation jobs
		job_date_from = request("job_date_from")
		job_date_to = request("job_date_to")
		job_tech_type = request("job_tech_type")
		job_tech_code = request("job_tech_code")
		jobyear = request("jobyear")
		jobmonth  = request("jobmonth")

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
		
	totaljobs=0
				
	'''Generate monthcommisionIHT table.
	sql = "Delete from tbltech_incentive where inc_month=" & jobmonth & " and inc_year=" & jobyear & " and inc_tech_code='" & job_tech_code & "'"
	CUD(sql)

	'only select jobs that has been posted and ticked for claims submission by technician and jobs with no approved claims yet
	sql2 = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, " & _
		"tbljob.job_Model, tbljob.job_model_desc,tbljob.job_tech_code, tbljob.job_tech_type, " & _
		"tbljob.job_tech_model_desc, tbljob.job_faulty_reason_cs, tbljob.job_overwty_allowance,tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt,tbljob.job_totalAmt, tbljob.job_inv_no, tbljob.job_inv_date, " & _
		"tblinvoice.inv_date, tblinvoice.inv_no ,tbltechnician.tech_type " & _
		"FROM tbljob INNER JOIN tblinvoice ON tbljob.job_inv_no = tblinvoice.inv_no INNER JOIN " & _
        "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code " & _
		"where tbljob.job_id is not null and tbljob.job_status='Posted' and tbljob.job_actual_wrty_status='Over' and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL " & _
		" and tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "  & _
  	    " and tbljob.job_tech_code = '" & job_tech_code & "' and tbljob.job_faulty_reason_cs <> 'Installation'" 'ignore installation jobs

	   set rs1 = server.CreateObject("adodb.recordset")
	   rs1.ActiveConnection = strconnect
	   rs1.Source = sql2
	   rs1.CursorLocation  = 3
	   rs1.Open
	   while not rs1.eof 
	
		totaljobs = totaljobs+1

		technician_tech_type = rs1("tech_type")
		job_jobcode = rs1("job_code")
		job_count = rs1("job_count")
		job_model = rs1("job_Model")

		'updated logic on the 060524 for water storage rates.
	 
		'if job_model = "TAE07-800" or job_model =  "TAE07-810" or job_model = "TAE07-811" or job_model = "TAE07-812" then 'only for water storage tanks
		'		sql = "select water_tank_over_wrty_percent from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'" 
		'		labourPer = selectid(sql)
		'else 'calc for non water storage tanks
			if (job_count = 1) then
				if technician_tech_type = "IHT" then
					sql = "select s_unit1_over_wrty_contractor from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'"  'based on amount 
					labourPer = selectid(sql)
				else
					sql = "select s_unit1_over_wrty_percent from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'"  'based on percentage
					labourPer = selectid(sql)
				end if
			end if
			if (job_count > 1) then
				if technician_tech_type = "IHT" then
					sql = "select s_unit2_over_wrty_contractor from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'"  'based on amount 
					labourPer = selectid(sql)
				else
					sql = "select s_unit2_over_wrty_percent from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'"  'based on percentage
					labourPer = selectid(sql)
				end if
			end if		
		'end if
	
		sql = "select s_over_wrty_spare_cost_percent from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'" 
		sparepartPer = selectid(sql)

		sparepayout = rs1("job_totalPartsAmt") * (sparepartPer/100)

		'20/06/2025 - to change the logic here (labourpayout)
		'if labourcost is zero and the checkbox is ticked then apply 45.00  (2nd job) and 75.00 (1st job) accodingly based on tech type

		job_totallabourAmt=rs1("job_totallabourAmt")

		if rs1("job_totallabourAmt") = 0 then 'over-warranty job that is not charged for labour (exceptional) but tech should still be paid
			if rs1("job_overwty_allowance") = "Yes" then
				if (job_count = 1) then
					sql = "select s_unit1_cust_over_wrty from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'" 
					job_totallabourAmt = selectid(sql) 'auto-fetch std labor rate from table since labour not entered by user
				end if
				if (job_count > 1) then
					sql = "select s_unit2_cust_over_wrty from tbltech_service_fee where s_tech_type ='" & technician_tech_type & "'" 
					job_totallabourAmt = selectid(sql) 'auto-fetch std labor rate from table since labour not entered by user
				end if	
			end if
		end if
		
		'labourpayout = round(rs1("job_totallabourAmt") * (labourPer/100))

		if technician_tech_type = "IHT" then	
			labourpayout = labourPer 'fixed amount
		else
			labourpayout = job_totallabourAmt * (labourPer/100)
		end if
		
		totalB_Payout = labourpayout + sparepayout
		totalNet_received = rs1("job_totalAmt") - totalB_Payout
	
		sql = "INSERT INTO tbltech_incentive (inc_tech_code, inc_month, inc_year, inc_invoice_no, inc_jobno, inc_modelno, inc_received_amt,inc_labor_charge_amt, inc_labor_overwriting,inc_labor_payout, " & _
	     " inc_part_charge, inc_part_overwriting,  inc_part_payout, inc_total_payout, inc_net_received) " & _
	     " values ('" & job_tech_code & "', " & jobmonth & ", " & jobyear & ",'" & rs1("job_inv_no") & "', '" & rs1("job_code") & "', '" & rs1("job_Model") & "' ," & chknumber2(rs1("job_totalAmt")) & "," & chknumber2(rs1("job_totallabourAmt")) & ", " & labourPer & ", " & _
	     ""& chknumber2(labourpayout) & "," & chknumber2(rs1("job_totalPartsAmt")) & ", " & sparepartPer & ", " & chknumber2(sparepayout) & "," & chknumber2(totalB_Payout) & ", " & chknumber2(totalNet_received) & ")" 
		CUD(sql)

	   rs1.movenext
	   wend
	   rs1.close
	
	sql2="update tbltech_claim set tc_overwrty_amt = (select sum(inc_total_payout) as total_over_wrty_payout from tbltech_incentive  "  & _
		 "where  inc_tech_code= '" & job_tech_code & "' and inc_month = '" & jobmonth & "' and inc_year='" & jobyear & "') where tc_tech_code= '" & job_tech_code & "' and tc_year = '" & jobyear & "' and tc_month='" & jobmonth & "'"
	CUD(sql2)

	sql22="update tbltech_claim set tc_overwrty_qty = " & totaljobs & " where tc_tech_code= '" & job_tech_code & "' and tc_year = '" & jobyear & "' and tc_month='" & jobmonth & "'"
	CUD(sql22)

	url = "rm_rpt_IncentiveTechnician_new.asp?type=showresult&job_tech_type=" & job_tech_type & "&job_tech_code=" & job_tech_code & "&tech_type=" & tech_type & "&job_date_from=" & job_date_from & " &job_date_to=" & job_date_to & " &jobmonth="&jobmonth&" &jobyear="& jobyear &" &loginerr=Report has been updated.#articletitle" 	
	'response.write url
	'response.end
'----------------------------------------------------------------------------------------------------    
  Case "monthcommisionIHT"

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
		
	'Generate monthcommisionIHT table.
	'don't delete the data if the claim has been checked/verified
	sql = "Delete from tblrpr_techcommission where rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_tech_type='IHT' and rpc_checkedby is NULL"
	CUD(sql)
	sql = "INSERT INTO tblrpr_techcommission (rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type)  Select '" & jobmonth & "' as jobmonth,'" & jobyear & "' as jobyear, tech_code, tech_name, tech_type " & _
	       " from tbltechnician where tech_code is not null and tech_type='IHT' " & _
	       "and tech_code not in (select rpc_tech_code from tblrpr_techcommission where rpc_checkedby is not NULL and rpc_month='" & jobmonth & "' and rpc_year='" & jobyear & "')"
	CUD(sql)

    ''' Loop monthcommisionIHT
	sql2 = "Select rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		  "rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,rpc_others2,  " & _
		  "rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_deduction_desc1,rpc_deduction_desc2,rpc_total_WH_qty,rpc_total_CF_qty " & _
		  "FROM tblrpr_techcommission  where rpc_tech_type='IHT' and rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_checkedby is NULL order by rpc_id "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql2
	rs1.CursorLocation  = 3
	rs1.Open

	while not rs1.eof 

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

	'IHT incentives covers both 1st and 2nd job but for under wrty only
		   sql2 = "SELECT count(tbljob.job_id) as totaljob1 FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted'  and tbljob.job_actual_wrty_status='Under' "& _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
			   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL " & _ 
			  	  "and tbljob.job_faulty_reason_cs <> 'Installation' and tbljob.job_count = '1'"
				  '"and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
  				  '"and month(tbljob.job_posteddate) = " & jobmonth & " and year(tbljob.job_posteddate) = " & jobyear & " and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL" 'and tbljob.job_count=1"
	
			rpc_serviceQty1 = selectid(sql2)
			rpc_serviceAmt1 = 0			

			sql3 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted'  and tbljob.job_actual_wrty_status='Under' "& _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
			   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL " & _ 
			  	  "and tbljob.job_faulty_reason_cs <> 'Installation' and tbljob.job_count > '1'"
				  '"and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
  				  '"and month(tbljob.job_posteddate) = " & jobmonth & " and year(tbljob.job_posteddate) = " & jobyear & " and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL" 'and tbljob.job_count=1"
	
			rpc_serviceQty2 = selectid(sql3)
			rpc_serviceAmt2 = 0			
			'rpc_serviceQtyTotal = ChkNumber(rpc_serviceQty1)+ChkNumber(rpc_serviceQty2)

			service_uw_fee1=0
			service_uw_fee2=0
			
			sql14 = "select s_unit1_under_wrty from tbltech_service_fee where s_tech_type='IHT'"
			service_uw_fee1 = selectid(sql14)

			sql15 = "select s_unit2_under_wrty from tbltech_service_fee where s_tech_type='IHT'"
			service_uw_fee2 = selectid(sql15)

			'logic to take into account, if technician has serviced water storage tank as this is a separate figure not be be considered part of totaljob
			'10052024 - claims for water storage only for underwarranty. Remove over-warranty
			'sql25= "select job_code,job_tech_model,job_faulty_reason_cs from tbljob where job_tech_code='" & rs1("rpc_tech_code") & "' and job_tech_model in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') " & _
			'"and tbljob.job_id is not null and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in('Under') and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL "
			'set rs25 = server.CreateObject("adodb.recordset")
			'rs25.ActiveConnection = strconnect
			'rs25.Source = sql25
			'rs25.CursorLocation  = 3
			'rs25.Open
			'storage_water_tank_text =""
			
			'water_storage_services_amt = 0
			'water_storage_qty = 0
			'while not rs25.eof 
			'	storage_water_tank_text = storage_water_tank_text + " / " + rs25("job_code") + " "  + "(" + rs25("job_tech_model") + ")"
			'	sql20 = "select s_storage_service from tbltech_service_fee where s_tech_type='IHT'"
			'	water_storage_services_amt=water_storage_services_amt + selectid(sql20)
			'	water_storage_qty = water_storage_qty + 1
			'rs25.movenext
			'wend
			'rpc_others=water_storage_services_amt 'calculates 80 per water storage
			'rs25.close

			rpc_others_desc =""
			'if len(storage_water_tank_text) > 2 then
			'	rpc_others_desc = storage_water_tank_text
			'	rpc_others_desc = replace(rpc_others_desc,"/","",1,1)
			'end if

			'if rpc_serviceQtyTotal > 30 then 
	           'balanceafter30 = rpc_serviceQtyTotal-30
				rpc_serviceAmt1 = rpc_serviceQty1 * service_uw_fee1 ''qty within 30 apply different rate		   
				rpc_serviceAmt2 = rpc_serviceQty2 * service_uw_fee2  'qty more than 30 apply different rate
			   'rpc_serviceQty2 = balanceafter30			   
			   'rpc_serviceQty1 = ChkNumber((rpc_serviceQtyTotal)-balanceafter30)
			'else
			 '  rpc_serviceAmt1 = ChkNumber(rpc_serviceQtyTotal)*service_uw_fee1
			 '  rpc_serviceQty1 = ChkNumber(rpc_serviceQtyTotal)
			 '  rpc_serviceAmt2 = 0
			 '  rpc_serviceQty2 = 0
			'end if
			
			rpc_serviceAmtTotal = rpc_serviceAmt1 + rpc_serviceAmt2
			
			'sql6 = "select s_allow_service from tbltech_service_fee where s_tech_type='IHT'"
			sql6 = "select tech_salary from tbltechnician where tech_code = '" & rs1("rpc_tech_code") & "'" '191224 - fetch salary from tech table
			rpc_techfees = selectid(sql6)
			
			
			sql7 = "select s_allow_car from tbltech_service_fee where s_tech_type='IHT'"
			rpc_car_allow = selectid(sql7)
			
			sql8 = "select s_allow_phone from tbltech_service_fee where s_tech_type='IHT'"
			rpc_phone_allow = selectid(sql8)

			'check if claim month is diffrent form submitted month

			sql16 ="select tc_month_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_month = selectid(sql16)

			sql15 ="select tc_year_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_year = selectid(sql15)
	
			processing_mth_diff = "No"
	        'assign the processing month as the new search

				if  jobmonth <> pay_month then
					if pay_year <> "" then 
						jobyear = pay_year
					end if

					if pay_month <> "" then					
						jobmonth = pay_month
					end if
					processing_mth_diff = "Yes"
				end if
	
			'revised logic 24112025 fetch from manual entry
			rpc_parking=0
			rpc_petrol=0
			rpc_toll=0
			rpc_target_incentive = 0
			job_period = ""
			claims_rec = ""

			sql20 = "select top 1 claims_id from tbltech_claim_manual where tech_code = '" & rs1("rpc_tech_code") & "' and completed = 'No' order by claims_id,tech_code"
			claims_rec = selectid(sql20)
			if claims_rec <> "" then 
					sql21 = "select total_petrol from tbltech_claim_manual where tech_code = '" & rs1("rpc_tech_code") & "' and claims_id = " & claims_rec & ""
					rpc_petrol = selectid(sql21)

					sql22 = "select total_toll from tbltech_claim_manual where tech_code = '" & rs1("rpc_tech_code") & "' and claims_id = " & claims_rec & ""
					rpc_toll = selectid(sql22)

					sql23 = "select total_parking from tbltech_claim_manual where tech_code = '" & rs1("rpc_tech_code") & "' and claims_id = " & claims_rec & ""
					rpc_parking = selectid(sql23)

					sql24 = "select total_incentive from tbltech_claim_manual where tech_code = '" & rs1("rpc_tech_code") & "' and claims_id = " & claims_rec & ""
					rpc_target_incentive = selectid(sql24)
		
					job_period = jobmonth & "/" & jobyear
					sql25 = "Update tbltech_claim_manual set completed = 'Yes', period = '" & job_period & "' where claims_id = " & claims_rec & ""
					CUD(sql25)
			end if 
			'------ old code-----------'
			'rpc_parking=0
			'sql14 = "select sum(tc_total_parking) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			'rpc_parking = selectid(sql14)
	
			'rpc_petrol=0
			'sql3 = "select sum(tc_total_petrol) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			'rpc_petrol = selectid(sql3)
	
			'sql2 = "select tc_total_toll from tbltech_claim where tc_month =  " & jobmonth & " and tc_year = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
		 	'sql2 = "select sum(tc_total_toll) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			'rpc_toll = selectid(sql2)
		
			rpc_hotel=0
			'sql4 = "select sum(th_claim_amount) from tbltech_claim_hotel where th_month =  " & jobmonth & " and th_year = " & jobyear & " and th_tech_code ='" & rs1("rpc_tech_code") & "'"
			sql4 = "select sum(tc_total_hotel) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_hotel = selectid(sql4)	

			rpc_service_allow = 0
			
			rpc_overwarranty_qty_1 = 0
			rpc_overwarranty_qty_1_amt = 0
			
			rpc_overwarranty_qty_2 = 0
			rpc_overwarranty_qty_2_amt = 0
			rpc_overwarranty_qty = 0
			
			rpc_others2=0
			rpc_deduction_ow_qty = 0
			rpc_deduction_ow = 0 '1st deduction amt/existing field
			
			rpc_deduction_sparepart_qty = 0
			rpc_deduction_sparepart = 0 '2nd deduction amt/existing field

			rpc_deduction_desc1 = ""
			rpc_deduction_desc2 = ""			
			rpc_others_desc2 = "Installation Service"
			'rpc_other2_service = 0  'installation comm rate for tech - obselete
			rpc_other2_WH_service_rate = 0  'installation comm rate for tech
			rpc_other2_CF_service_rate = 0  'installation comm rate for tech 
	        rpc_others2_qty= 0  ' total installation for tech

			total_WH_qty = 0 
			total_CF_qty = 0 

			sql5 = "select sum(tc_overwrty_amt) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_overwarranty_fee = selectid(sql5)			
			
			sql30 = "select tc_overwrty_qty from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_overwarranty_qty = selectid(sql30)			
			
	        'deduction data
			sql6 = "select tc_deduc1 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc1 = selectid(sql6)
	
			sql7 = "select sum(tc_deducamt1) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_ow = selectid(sql7)

			sql8 = "select tc_deduc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc2 = selectid(sql8)

			sql9 = "select sum(tc_deducamt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_sparepart = selectid(sql9)

			sql14 = "select s_installation_service_CF from tbltech_service_fee where s_tech_type='IHT'" 'rate for CF
			'rpc_other2_service = selectid(sql14)
			rpc_other2_CF_service_rate = selectid(sql14)

			sql16 = "select s_installation_service_WH from tbltech_service_fee where s_tech_type='IHT'" 'rate for WH
			rpc_other2_WH_service_rate = selectid(sql16)

			'191225 need to calculate separately for WH/CF

			sql13 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted'  and  tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
			   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL " & _ 
			  	  "and tbljob.job_faulty_reason_cs = 'Installation'  and tbljob.job_tech_type = 'CF'"
			
			total_CF_qty = selectid(sql13)

			sql15 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted'  and  tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
			   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL " & _ 
			  	  "and tbljob.job_faulty_reason_cs = 'Installation'  and tbljob.job_tech_type = 'WH'"
			
			total_WH_qty = selectid(sql15)

			rpc_others2_qty = total_CF_qty + total_WH_qty
			'rpc_others2 =	rpc_other2_service * rpc_others2_qty
			rpc_others2 = (total_CF_qty * rpc_other2_CF_service_rate) + (total_WH_qty * rpc_other2_WH_service_rate)
			
			rpc_deduction_total= rpc_deduction_ow + rpc_deduction_sparepart
			rpc_serviceAmtTotal=ChkNumber2Decimal(rpc_serviceAmtTotal)
			rpc_techfee=ChkNumber2Decimal(rpc_techfee)
			rpc_car_allow=ChkNumber2Decimal(rpc_car_allow)
			rpc_phone_allow=ChkNumber2Decimal(rpc_phone_allow)
			rpc_toll=ChkNumber2Decimal(rpc_toll)
			rpc_parking=ChkNumber2Decimal(rpc_parking)
			rpc_petrol=ChkNumber2Decimal(rpc_petrol)
			rpc_hotel=ChkNumber2Decimal(rpc_hotel)
			rpc_service_allow=ChkNumber2Decimal(rpc_service_allow)
			rpc_overwarranty_fee=ChkNumber2Decimal(rpc_overwarranty_fee)
			rpc_others=ChkNumber2Decimal(rpc_others)
			rpc_others2=ChkNumber2Decimal(rpc_others2)
			rpc_deduction_total=ChkNumber2Decimal(rpc_deduction_total)
	
			rpc_total = rpc_serviceAmtTotal+rpc_techfees+rpc_car_allow+rpc_phone_allow+rpc_toll+rpc_parking+rpc_petrol+rpc_hotel+rpc_service_allow+rpc_overwarranty_fee+rpc_others+rpc_others2 - rpc_deduction_total

			sql4 = "Update tblrpr_techcommission set " & _
					"rpc_serviceQty1=" & ChkNumber(rpc_serviceQty1) & ", rpc_serviceAmt1=" & ChkNumber(rpc_serviceAmt1) & ", rpc_serviceQty2=" & ChkNumber(rpc_serviceQty2) & ", " & _
					"rpc_serviceAmt2=" & ChkNumber(rpc_serviceAmt2) & ", rpc_serviceQtyTotal=" & ChkNumber(rpc_serviceQtyTotal) & ", rpc_serviceAmtTotal=" & ChkNumber(rpc_serviceAmtTotal) & ",  " & _
					"rpc_techfees=" & ChkNumber(rpc_techfees) & ", rpc_car_allow=" & ChkNumber(rpc_car_allow) & ", rpc_phone_allow=" & ChkNumber(rpc_phone_allow) & ",  " & _
					"rpc_toll=" & ChkNumber(rpc_toll) & ", rpc_parking=" & ChkNumber(rpc_parking) & ", rpc_petrol=" & ChkNumber(rpc_petrol) & ",  " & _
					"rpc_hotel=" & ChkNumber(rpc_hotel) & ", rpc_service_allow=" & ChkNumber(rpc_service_allow) & ", rpc_overwarranty_qty=" & ChkNumber(rpc_overwarranty_qty) & ",  " & _
					"rpc_overwarranty_fee=" & ChkNumber(rpc_overwarranty_fee) & ", rpc_others=" & ChkNumber(rpc_others) & ", rpc_deduction_ow_qty=" & ChkNumber(rpc_deduction_ow_qty) & ",  " & _
					"rpc_deduction_ow=" & ChkNumber(rpc_deduction_ow) & ",rpc_deduction_sparepart_qty=" & ChkNumber(rpc_deduction_sparepart_qty) & ",  " & _
					"rpc_deduction_sparepart=" & ChkNumber(rpc_deduction_sparepart) & ",rpc_deduction_desc1= '" & rpc_deduction_desc1 & "', rpc_deduction_desc2= '" & rpc_deduction_desc2 & "', rpc_deduction_total=" & ChkNumber(rpc_deduction_total) & ", rpc_total=" & ChkNumber(rpc_total) & ", " & _
					"rpc_others_desc= '" & rpc_others_desc & "',rpc_others_desc2= '" & rpc_others_desc2 & "',rpc_others2=" & ChkNumber(rpc_others2) & ",rpc_others2_qty=" & ChkNumber(rpc_others2_qty) & ", rpc_submitted_date = '" & ChkDateTimeMySQL(now()) & "', " & _
					"rpc_water_storage_qty = " & ChkNumber(water_storage_qty) & ", rpc_target_incentive = " & ChkNumber(rpc_target_incentive) & ",rpc_total_CF_qty = " & ChkNumber(total_CF_qty) & ",rpc_total_WH_qty = " & ChkNumber(total_WH_qty) & " where rpc_id = " & rs1("rpc_id") & ""
			CUD(sql4)

	rs1.movenext
	wend
	rs1.close	

	url = "rm_rpt_tech_monthcommisionIHT.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Report has been updated.#articletitle" 	 
 
 '----------------------------------------------------------------------------------------------------    
  Case "monthcommisionIHC"
	
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
		
	'Generate monthcommisionIHT table.
	'don't delete the data if the claim has been checked/verified
	sql = "Delete from tblrpr_techcommission where rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_tech_type='IHC' and rpc_checkedby is NULL"
	CUD(sql)
	sql = "INSERT INTO tblrpr_techcommission (rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type)  Select '" & jobmonth & "' as jobmonth,'" & jobyear & "' as jobyear, tech_code, tech_name, tech_type " & _
	       " from tbltechnician where tech_code is not null and tech_type='IHC' " & _
	       "and tech_code not in (select rpc_tech_code from tblrpr_techcommission where rpc_checkedby is not NULL and rpc_month='" & jobmonth & "' and rpc_year='" & jobyear & "')"
	CUD(sql)
	
    ''' Loop monthcommisionIHT
	sql2 = "Select rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		  "rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,rpc_others2,  " & _
		  "rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_deduction_desc1,rpc_deduction_desc2 " & _
		  "FROM tblrpr_techcommission  where rpc_tech_type='IHC' and rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_checkedby is NULL order by rpc_id "

	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql2
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

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

	service_uw_fee1=0
	service_uw_fee2=0
			
	sql14 = "select s_unit1_under_wrty from tbltech_service_fee where s_tech_type='IHC'"
	service_uw_fee1 = selectid(sql14)

	sql15 = "select s_unit2_under_wrty from tbltech_service_fee where s_tech_type='IHC'"
	service_uw_fee2 = selectid(sql15)
	
   'rpc_serviceQty1, 
   'IHC incentives covers the 1st job and 2nd job for under wrty only
	
    'update  - 01/12/2023
	'the logic below calculates based on the posted month ONLY
	'otherwise select based on job_claim_approved is NULL and once approved mark it as 'Yes', this way though posted on Nov can process in Dec as not based on posted date

	sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in ('Under') " & _
		  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
	   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_claim_approved is NULL " & _
		  "and tbljob.job_tech_model not in ('TAE07-800','`','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
	 'if water storage models exist, store them collectively in others_desc, the amt will be manually adjusted in claim form
	

	rpc_serviceQty1 = selectid(sql2)
	rpc_serviceAmt1 = ChkNumber(rpc_serviceQty1) * service_uw_fee1
	 
	sql24 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in ('Under','Over') " & _
		  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
	   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_claim_approved is NULL " & _
		  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  
	rpc_serviceQty3 = selectid(sql24) 'need to calc this separately for 1st qty for over. this is needed for allowance calc

	sql3 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null " & _
		  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in ('Under') " & _
		  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
	   	  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count > 1 and tbljob.job_claim_approved is NULL " & _
		  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc

	rpc_serviceQty2 = selectid(sql3)
	rpc_serviceAmt2 = ChkNumber(rpc_serviceQty2) * service_uw_fee2
	
	'rpc_serviceQtyTotal = ChkNumber(rpc_serviceQty1)+ChkNumber(rpc_serviceQty2)
	rpc_serviceQtyTotal = ChkNumber(rpc_serviceQty3) 'for IHC allowance are calc based on the 1st qty of under and over.
	rpc_serviceAmtTotal = ChkNumber(rpc_serviceAmt1)+ChkNumber(rpc_serviceAmt2)
	
	'10052024 - claims for water storage only for underwarranty. Remove over-warranty
	sql25= "select job_code,job_tech_model from tbljob where job_tech_code='" & rs1("rpc_tech_code") & "' and job_tech_model in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') " & _
	"and tbljob.job_id is not null and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in('Under') and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL "
	set rs25 = server.CreateObject("adodb.recordset")
	
	rs25.ActiveConnection = strconnect
	rs25.Source = sql25
	rs25.CursorLocation  = 3
	rs25.Open
	storage_water_tank_text =""
	water_storage_services_amt = 0
	water_storage_qty = 0
	while not rs25.eof 
			storage_water_tank_text = storage_water_tank_text + " / " + rs25("job_code") + " "  + "(" + rs25("job_tech_model") + ")"
			sql20 = "select s_storage_service from tbltech_service_fee where s_tech_type='IHC'"
			water_storage_services_amt=water_storage_services_amt + selectid(sql20)
			water_storage_qty = water_storage_qty + 1
	rs25.movenext
	wend
	rpc_others=water_storage_services_amt 'calculates 80 per water storage
	rs25.close

	rpc_others_desc =""
	if len(storage_water_tank_text) > 2 then	
		rpc_others_desc = storage_water_tank_text
		rpc_others_desc = replace(rpc_others_desc,"/","",1,1)	
	end if	

	'rpc_techfees
	rpc_techfees = 0
	rpc_car_allow = 0
	rpc_phone_allow = 0
	rpc_service_allow = 0	

	'rpc_car_allow
	if rpc_serviceQtyTotal > 79 and rpc_serviceQtyTotal < 100 then 
		sql16 = "select s_allow_excc_car_range1 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_car_allow = selectid(sql16)	  
	elseif rpc_serviceQtyTotal > 99 and rpc_serviceQtyTotal < 135 then 
		sql16 = "select s_allow_excc_car_range2 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_car_allow = selectid(sql16)	   
	elseif rpc_serviceQtyTotal >= 135 then 
		sql16 = "select s_allow_excc_car_range3 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_car_allow = selectid(sql16)
	else
	   rpc_car_allow = 0
	end if  

	'rpc_phone_allow
	if rpc_serviceQtyTotal > 79 and rpc_serviceQtyTotal < 100 then 
		sql17 = "select s_allow_excc_phone_range1 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_phone_allow = selectid(sql17)		
	elseif rpc_serviceQtyTotal > 99 and rpc_serviceQtyTotal < 135 then 	   
		sql17 = "select s_allow_excc_phone_range2 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_phone_allow = selectid(sql17)
	elseif rpc_serviceQtyTotal >= 135 then 	   
		sql17 = "select s_allow_excc_phone_range3 from tbltech_service_fee where s_tech_type='IHC'"
		rpc_phone_allow = selectid(sql17)
	else
	   rpc_phone_allow = 0
	end if  	
	
	rpc_toll=0
	rpc_parking=0
	rpc_petrol=0
	rpc_hotel=0
	rpc_service_allow=0
	
	'rpc_service_allow
	if rpc_serviceQtyTotal > 79 and rpc_serviceQtyTotal < 100 then 
	   sql18 = "select s_allow_excc_service_range1 from tbltech_service_fee where s_tech_type='IHC'"
	   rpc_service_allow = selectid(sql18)
	elseif rpc_serviceQtyTotal > 99 and rpc_serviceQtyTotal < 135 then  
	   sql18 = "select s_allow_excc_service_range2 from tbltech_service_fee where s_tech_type='IHC'"
	   rpc_service_allow = selectid(sql18)
	elseif rpc_serviceQtyTotal >= 135 then 		
	   sql18 = "select s_allow_excc_service_range3 from tbltech_service_fee where s_tech_type='IHC'"
	   rpc_service_allow = selectid(sql18)
	else
		rpc_service_allow = 0
	end if 

		
	sql16 ="select tc_month_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
	pay_month = selectid(sql16)

	sql15 ="select tc_year_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
	pay_year = selectid(sql15)
	
	processing_mth_diff = "No"
	'assign the processing month as the new search

	if  jobmonth <> pay_month then
				if pay_year <> "" then 
						jobyear = pay_year
				end if

				if pay_month <> "" then					
						jobmonth = pay_month
				end if
				processing_mth_diff = "Yes"				
	end if

	sql5 = "select sum(tc_overwrty_amt) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
	rpc_overwarranty_fee = selectid(sql5)

	sql30 = "select sum(tc_overwrty_qty) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
	rpc_overwarranty_qty = selectid(sql30)
	'rpc_overwarranty_fee = 0 'IHC will not have over wrtty cases, as it's already included in total service job 12/12 christina told to include again.

	'rpc_others
	'rpc_deduction_ow_qty
	rpc_deduction_ow_qty = 0
	'rpc_deduction_ow
	rpc_deduction_ow = 0
	
			sql6 = "select tc_deduc1 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc1 = selectid(sql6)
	
			sql7 = "select sum(tc_deducamt1) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_ow = selectid(sql7)

			sql8 = "select tc_deduc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc2 = selectid(sql8)

			sql9 = "select sum(tc_deducamt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_sparepart = selectid(sql9)

		'	sql11 = "select sum(tc_otheramt1) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
		'	rpc_others = selectid(sql11)

			sql12 = "select tc_otherdesc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others_desc2 = selectid(sql12)

			sql13 = "select sum(tc_otheramt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others2 = selectid(sql13)
	
	'rpc_deduction_total
     rpc_deduction_total = ChkNumber(rpc_deduction_ow) + ChkNumber(rpc_deduction_sparepart)
	
	'rpc_total
	rpc_total = (rpc_serviceAmtTotal+rpc_techfees+rpc_car_allow+rpc_phone_allow+rpc_toll+rpc_parking+rpc_petrol+rpc_hotel+rpc_service_allow+rpc_overwarranty_fee+rpc_others+rpc_others2) - rpc_deduction_total
	
	sql4 = "Update tblrpr_techcommission set " & _
			"rpc_serviceQty1=" & ChkNumber(rpc_serviceQty1) & ", rpc_serviceAmt1=" & ChkNumber(rpc_serviceAmt1) & ", rpc_serviceQty2=" & ChkNumber(rpc_serviceQty2) & ", " & _
			"rpc_serviceAmt2=" & ChkNumber(rpc_serviceAmt2) & ", rpc_serviceQtyTotal=" & ChkNumber(rpc_serviceQtyTotal) & ", rpc_serviceAmtTotal=" & ChkNumber(rpc_serviceAmtTotal) & ",  " & _
			"rpc_techfees=" & ChkNumber(rpc_techfees) & ", rpc_car_allow=" & ChkNumber(rpc_car_allow) & ", rpc_phone_allow=" & ChkNumber(rpc_phone_allow) & ",  " & _
			"rpc_toll=" & ChkNumber(rpc_toll) & ", rpc_parking=" & ChkNumber(rpc_parking) & ", rpc_petrol=" & ChkNumber(rpc_petrol) & ",  " & _
			"rpc_hotel=" & ChkNumber(rpc_hotel) & ", rpc_service_allow=" & ChkNumber(rpc_service_allow) & ", rpc_overwarranty_qty=" & ChkNumber(rpc_overwarranty_qty) & ",  " & _
			"rpc_overwarranty_fee=" & ChkNumber(rpc_overwarranty_fee) & ", rpc_others=" & ChkNumber(rpc_others) & ", rpc_deduction_ow_qty=" & ChkNumber(rpc_deduction_ow_qty) & ",  " & _
			"rpc_deduction_ow=" & ChkNumber(rpc_deduction_ow) & ",rpc_deduction_sparepart_qty=" & ChkNumber(rpc_deduction_sparepart_qty) & ",  " & _
			"rpc_deduction_sparepart=" & ChkNumber(rpc_deduction_sparepart) & ", rpc_deduction_total=" & ChkNumber(rpc_deduction_total) & ", rpc_total=" & ChkNumber(rpc_total) & ", " & _
			"rpc_deduction_desc1= '" & rpc_deduction_desc1 & "', rpc_deduction_desc2= '" & rpc_deduction_desc2 & "', " & _
			"rpc_others_desc= '" & rpc_others_desc & "',rpc_others_desc2= '" & rpc_others_desc2 & "',rpc_others2=" & ChkNumber(rpc_others2) & ", rpc_submitted_date = '" & ChkDateTimeMySQL(now()) & "', " & _
			"rpc_water_storage_qty = " & ChkNumber(water_storage_qty) & " where rpc_id = " & rs1("rpc_id") &""
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
	url = "rm_rpt_tech_monthcommisionIHC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Report has been updated.#articletitle" 	 
      	     	
'---------------------------------------------------------------------------------------------------- 
  
  Case "monthcommisionIC"

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
	
	'Generate monthcommisionIHT table.
	'don't delete the data if the claim has been checked/verified
	sql = "Delete from tblrpr_techcommission where rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_tech_type='IC' and rpc_checkedby is NULL"
	CUD(sql)
	sql = "INSERT INTO tblrpr_techcommission (rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type)  Select '" & jobmonth & "' as jobmonth,'" & jobyear & "' as jobyear, tech_code, tech_name, tech_type " & _
	       " from tbltechnician where tech_code is not null and tech_type='IC' " & _
	       "and tech_code not in (select rpc_tech_code from tblrpr_techcommission where rpc_checkedby is not NULL and rpc_month='" & jobmonth & "' and rpc_year='" & jobyear & "')"
	CUD(sql)
	
    ''' Loop monthcommisionIHT
	sql2 = "Select rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		  "rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,rpc_others2,  " & _
		  "rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_deduction_desc1,rpc_deduction_desc2 " & _
		  "FROM tblrpr_techcommission  where rpc_tech_type='IC' and rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_checkedby is NULL order by rpc_id "

	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql2
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

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

	service_uw_fee1=0
	service_uw_fee2=0
			
	sql14 = "select s_unit1_under_wrty from tbltech_service_fee where s_tech_type='IC'"
	service_uw_fee1 = selectid(sql14)

	sql15 = "select s_unit2_under_wrty from tbltech_service_fee where s_tech_type='IC'"
	service_uw_fee2 = selectid(sql15)

			'rpc_serviceQty1, 
			sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status='Under'" & _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
				  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_claim_approved is NULL " & _
				  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc

				 ' "and month(tbljob.job_posteddate) = " & jobmonth & " and year(tbljob.job_posteddate) = " & jobyear & " and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_claim_approved is NULL"
			rpc_serviceQty1 = selectid(sql2)
			rpc_serviceAmt1 = ChkNumber(rpc_serviceQty1) * service_uw_fee1
			
			'rpc_serviceQty2, 
			sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status='Under'" & _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
				  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count>1 and tbljob.job_claim_approved is NULL " & _
				  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
				  '"and month(tbljob.job_posteddate) = " & jobmonth & " and year(tbljob.job_posteddate) = " & jobyear & " and tbljob.job_submitforclaims='Yes' and tbljob.job_count>1 and tbljob.job_claim_approved is NULL"
			rpc_serviceQty2 = selectid(sql2)
			rpc_serviceAmt2 = ChkNumber(rpc_serviceQty2) * service_uw_fee2
			
			rpc_serviceQtyTotal = ChkNumber(rpc_serviceQty1)+ChkNumber(rpc_serviceQty2)
			rpc_serviceAmtTotal = ChkNumber(rpc_serviceAmt1)+ChkNumber(rpc_serviceAmt2)
			
			'rpc_techfees
			rpc_techfees = 0
			
			'rpc_car_allow
			rpc_car_allow = 0
			
			'rpc_phone_allow
			rpc_phone_allow = 0
			
			'rpc_toll 
			rpc_toll=0
			'rpc_parking
			rpc_parking=0
			'rpc_petrol
			rpc_petrol=0
			'rpc_hotel
			rpc_hotel=0
			'rpc_service_allow
			rpc_service_allow=0
			
			'rpc_service_allow
			rpc_service_allow = 0
			
			'rpc_overwarranty_qty
			sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted' " & _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
				  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_actual_wrty_status = 'Over' and tbljob.job_claim_approved is NULL"
				 ' "and month(tbljob.job_posteddate) = " & jobmonth & " and year(tbljob.job_posteddate) = " & jobyear & " and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_actual_wrty_status = 'Over' and tbljob.job_claim_approved is NULL"
			rpc_overwarranty_qty_1 = selectid(sql2)
			rpc_overwarranty_qty_1_amt = ChkNumber(rpc_overwarranty_qty_1)  * 7
			
			rpc_overwarranty_qty_2 = 0
			rpc_overwarranty_qty_2_amt = 0
			'rpc_overwarranty_qty
			rpc_overwarranty_qty = ChkNumber(rpc_overwarranty_qty_1) + ChkNumber(rpc_overwarranty_qty_2)
			
			'rpc_overwarranty_fee   
			rpc_overwarranty_fee = ChkNumber(rpc_overwarranty_qty_1_amt) + ChkNumber(rpc_overwarranty_qty_2_amt)

			
			sql30 = "select sum(tc_overwrty_qty) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_overwarranty_qty = selectid(sql30) 'use this, the same as other technicians


			'logic to take into account, if technician has serviced water storage tank as this is a separate figure not be be considered part of totaljob
			'10052024 - claims for water storage only for underwarranty. Remove over-warranty
			sql25= "select job_code,job_tech_model from tbljob where job_tech_code='" & rs1("rpc_tech_code") & "' and job_tech_model in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') " & _
			"and tbljob.job_id is not null and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in('Under') and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL "
			set rs25 = server.CreateObject("adodb.recordset")
			rs25.ActiveConnection = strconnect
			rs25.Source = sql25
			rs25.CursorLocation  = 3
			rs25.Open
			storage_water_tank_text =""
			water_storage_services_amt = 0
			water_storage_services_amt = 0
			while not rs25.eof 
					storage_water_tank_text = storage_water_tank_text + " / " + rs25("job_code") + " "  + "(" + rs25("job_tech_model") + ")"
					sql20 = "select s_storage_service from tbltech_service_fee where s_tech_type='IHC'"
					water_storage_services_amt=water_storage_services_amt + selectid(sql20)
					water_storage_qty = water_storage_qty + 1
			rs25.movenext
			wend
			rpc_others=water_storage_services_amt 'calculates 80 per water storage
			rs25.close
			
			rpc_others_desc =""
			if len(storage_water_tank_text) > 2 then
				rpc_others_desc = storage_water_tank_text
				rpc_others_desc = replace(rpc_others_desc,"/","",1,1)
			end if
			
			'rpc_others
			'rpc_deduction_ow_qty
			rpc_deduction_ow_qty = 0
			'rpc_deduction_ow
			rpc_deduction_ow = 0

			sql16 ="select tc_month_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_month = selectid(sql16)

			sql15 ="select tc_year_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_year = selectid(sql15)
	
			processing_mth_diff = "No"
	        'assign the processing month as the new search

				if  jobmonth <> pay_month then
					if pay_year <> "" then 
						jobyear = pay_year
					end if

					if pay_month <> "" then					
						jobmonth = pay_month
					end if
					processing_mth_diff = "Yes"				
				end if

			
			'rpc_deduction_total
			sql6 = "select tc_deduc1 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc1 = selectid(sql6)
	
			sql7 = "select sum(tc_deducamt1) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_ow = selectid(sql7)

			sql8 = "select tc_deduc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc2 = selectid(sql8)

			sql9 = "select sum(tc_deducamt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_sparepart = selectid(sql9)
		
			sql12 = "select tc_otherdesc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others_desc2 = selectid(sql12)

			sql13 = "select sum(tc_otheramt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others2 = selectid(sql13)
	
	'rpc_deduction_total
     rpc_deduction_total = ChkNumber(rpc_deduction_ow) + ChkNumber(rpc_deduction_sparepart)
	
	'rpc_total
	rpc_total = (rpc_serviceAmtTotal+rpc_techfees+rpc_car_allow+rpc_phone_allow+rpc_toll+rpc_parking+rpc_petrol+rpc_hotel+rpc_service_allow+rpc_overwarranty_fee+rpc_others+rpc_others2) - rpc_deduction_total
	
			sql4 = "Update tblrpr_techcommission set " & _
					"rpc_serviceQty1=" & ChkNumber(rpc_serviceQty1) & ", rpc_serviceAmt1=" & ChkNumber(rpc_serviceAmt1) & ", rpc_serviceQty2=" & ChkNumber(rpc_serviceQty2) & ", " & _
					"rpc_serviceAmt2=" & ChkNumber(rpc_serviceAmt2) & ", rpc_serviceQtyTotal=" & ChkNumber(rpc_serviceQtyTotal) & ", rpc_serviceAmtTotal=" & ChkNumber(rpc_serviceAmtTotal) & ",  " & _
					"rpc_techfees=" & ChkNumber(rpc_techfees) & ", rpc_car_allow=" & ChkNumber(rpc_car_allow) & ", rpc_phone_allow=" & ChkNumber(rpc_phone_allow) & ",  " & _
					"rpc_toll=" & ChkNumber(rpc_toll) & ", rpc_parking=" & ChkNumber(rpc_parking) & ", rpc_petrol=" & ChkNumber(rpc_petrol) & ",  " & _
					"rpc_hotel=" & ChkNumber(rpc_hotel) & ", rpc_service_allow=" & ChkNumber(rpc_service_allow) & ", rpc_overwarranty_qty=" & ChkNumber(rpc_overwarranty_qty) & ",  " & _
					"rpc_overwarranty_fee=" & ChkNumber(rpc_overwarranty_fee) & ", rpc_others=" & ChkNumber(rpc_others) & ", rpc_deduction_ow_qty=" & ChkNumber(rpc_deduction_ow_qty) & ",  " & _
					"rpc_deduction_ow=" & ChkNumber(rpc_deduction_ow) & ",rpc_deduction_sparepart_qty=" & ChkNumber(rpc_deduction_sparepart_qty) & ",  " & _
					"rpc_deduction_sparepart=" & ChkNumber(rpc_deduction_sparepart) & ", rpc_deduction_total=" & ChkNumber(rpc_deduction_total) & ", rpc_total=" & ChkNumber(rpc_total) & ", " & _
					"rpc_deduction_desc1= '" & rpc_deduction_desc1 & "', rpc_deduction_desc2= '" & rpc_deduction_desc2 & "', " & _
					"rpc_others_desc= '" & rpc_others_desc & "',rpc_others_desc2= '" & rpc_others_desc2 & "',rpc_others2=" & ChkNumber(rpc_others2) & ", rpc_submitted_date = '" & ChkDateTimeMySQL(now()) & "', " & _			
					"rpc_water_storage_qty = " & ChkNumber(water_storage_qty) & " where rpc_id = " & rs1("rpc_id") 
			CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	

	url = "rm_rpt_tech_monthcommisionIC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Report has been updated.#articletitle" 	 
      	     	
'----------------------------------------------------------------------------------------------------    
  
  Case "monthcommisionTPC"
  
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
	
	'Generate monthcommisionIHT table.
	'don't delete the data if the claim has been checked/verified
	sql = "Delete from tblrpr_techcommission where rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_tech_type='TPC' and rpc_checkedby is NULL"
	CUD(sql)
	sql = "INSERT INTO tblrpr_techcommission (rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type)  Select '" & jobmonth & "' as jobmonth,'" & jobyear & "' as jobyear, tech_code, tech_name, tech_type " & _
	       " from tbltechnician where tech_code is not null and tech_type='TPC' " & _
	       "and tech_code not in (select rpc_tech_code from tblrpr_techcommission where rpc_checkedby is not NULL and rpc_month='" & jobmonth & "' and rpc_year='" & jobyear & "')"
	CUD(sql)
	
    ''' Loop monthcommisionIHT
	sql2 = "Select rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_tech_type, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		  "rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,rpc_others2,  " & _
		  "rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_deduction_desc1,rpc_deduction_desc2 " & _
		  "FROM tblrpr_techcommission  where rpc_tech_type='TPC' and rpc_month=" & jobmonth & " and rpc_year=" & jobyear & " and rpc_checkedby is NULL order by rpc_id "
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql2
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 

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
			sql14 = "select s_unit1_under_wrty from tbltech_service_fee where s_tech_type='TPC'"
			service_uw_fee1 = selectid(sql14)

			sql15 = "select s_unit2_under_wrty from tbltech_service_fee where s_tech_type='TPC'"
			service_uw_fee2 = selectid(sql15)

			'rpc_serviceQty1, 
			sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status='Under' " & _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
				  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count=1 and tbljob.job_claim_approved is NULL " & _
				  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
			
			rpc_serviceQty1 = selectid(sql2)
			rpc_serviceAmt1 = ChkNumber(rpc_serviceQty1) * service_uw_fee1
			
			'rpc_serviceQty2, 
			sql2 = "SELECT count(tbljob.job_id) as totaljob FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
				  "where tbljob.job_id is not null " & _
				  "and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status='Under' " & _
				  "and tbljob.job_tech_code = '" & rs1("rpc_tech_code") & "' " & _ 
				  "and tbljob.job_submitforclaims='Yes' and tbljob.job_count>1 and tbljob.job_claim_approved is NULL " & _
				  "and tbljob.job_tech_model not in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') "  'need to isolate water storage tank into other2 and not part of service calc
				 
			rpc_serviceQty2 = selectid(sql2)
			rpc_serviceAmt2 = ChkNumber(rpc_serviceQty2) * service_uw_fee2
			
			rpc_serviceQtyTotal = ChkNumber(rpc_serviceQty1)+ChkNumber(rpc_serviceQty2)
			rpc_serviceAmtTotal = ChkNumber(rpc_serviceAmt1)+ChkNumber(rpc_serviceAmt2)

			'logic to take into account, if technician has serviced water storage tank as this is a separate figure not be be considered part of totaljob
			'10052024 - claims for water storage only for underwarranty. Remove over-warranty
			sql25= "select job_code,job_tech_model from tbljob where job_tech_code='" & rs1("rpc_tech_code") & "' and job_tech_model in ('TAE07-800','TAE07-810','TAE07-811','TAE07-812') " & _
			"and tbljob.job_id is not null and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in('Under') and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL "
			set rs25 = server.CreateObject("adodb.recordset")
			rs25.ActiveConnection = strconnect
			rs25.Source = sql25
			rs25.CursorLocation  = 3
			rs25.Open
			storage_water_tank_text =""
			water_storage_services_amt = 0
			water_storage_qty = 0
			while not rs25.eof 
					storage_water_tank_text = storage_water_tank_text + " / " + rs25("job_code") + " "  + "(" + rs25("job_tech_model") + ")"
					sql20 = "select s_storage_service from tbltech_service_fee where s_tech_type='IHC'"
					water_storage_services_amt=water_storage_services_amt + selectid(sql20)
					water_storage_qty = water_storage_qty + 1
			rs25.movenext
			wend

			rpc_others=water_storage_services_amt 'calculates 80 per water storage
			rs25.close
			rpc_others_desc =""
			if len(storage_water_tank_text) > 2 then
				rpc_others_desc = storage_water_tank_text
				rpc_others_desc = replace(rpc_others_desc,"/","",1,1)
			end if
	

			'rpc_techfees
			rpc_techfees = 0
			
			'rpc_car_allow
			rpc_car_allow = 0
			
			'rpc_phone_allow
			rpc_phone_allow = 0

			processing_mth_dif	= "No"
	        'assign the processing month as the new search

				if  jobmonth <> pay_month then
					if pay_year <> "" then 
						jobyear = pay_year
					end if

					if pay_month <> "" then					
						jobmonth = pay_month
					end if
					processing_mth_diff = "Yes"
				end if
			
			'rpc_toll 
			rpc_toll=0
			'rpc_parking
			rpc_parking=0
			'rpc_petrol
			rpc_petrol=0
			sql3 = "select sum(tc_total_extramileage) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_petrol = selectid(sql3)
	
			'rpc_hotel
			rpc_hotel=0
			'rpc_service_allow
			rpc_service_allow=0
			
			'rpc_service_allow
			rpc_service_allow = 0
			
			'rpc_others
			rpc_others2=0
			'rpc_deduction_ow_qty
			rpc_deduction_ow_qty = 0
			'rpc_deduction_ow
			rpc_deduction_ow = 0
			rpc_overwarranty_fee = 0


			sql16 ="select tc_month_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_month = selectid(sql16)

			sql15 ="select tc_year_process from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			pay_year = selectid(sql15)	
				
			sql6 = "select tc_deduc1 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc1 = selectid(sql6)
	
			sql7 = "select sum(tc_deducamt1) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_ow = selectid(sql7)

			sql8 = "select tc_deduc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_desc2 = selectid(sql8)

			sql9 = "select sum(tc_deducamt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_deduction_sparepart = selectid(sql9)

			sql12 = "select tc_otherdesc2 from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others_desc2 = selectid(sql12)

			sql13 = "select sum(tc_otheramt2) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_others2 = selectid(sql13)
			 rpc_deduction_total = ChkNumber(rpc_deduction_ow) + ChkNumber(rpc_deduction_sparepart)
			
			sql5 = "select sum(tc_overwrty_amt) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_overwarranty_fee = selectid(sql5)

			sql30 = "select sum(tc_overwrty_qty) from tbltech_claim where tc_month_process =  " & jobmonth & " and tc_year_process = " & jobyear & " and tc_tech_code ='" & rs1("rpc_tech_code") & "'"
			rpc_overwarranty_qty = selectid(sql30)

			'rpc_total
			rpc_total = (ChkNumber(rpc_serviceAmtTotal)+ChkNumber(rpc_techfees+rpc_car_allow)+ChkNumber(rpc_phone_allow)+ChkNumber(rpc_toll)+ChkNumber(rpc_parking)+ChkNumber(rpc_petrol)+ChkNumber(rpc_hotel)+ChkNumber(rpc_service_allow)+ChkNumber(rpc_overwarranty_fee)+ChkNumber(rpc_others)+ChkNumber(rpc_others2)) - rpc_deduction_total

			sql4 = "Update tblrpr_techcommission set " & _
					"rpc_serviceQty1=" & ChkNumber(rpc_serviceQty1) & ", rpc_serviceAmt1=" & ChkNumber(rpc_serviceAmt1) & ", rpc_serviceQty2=" & ChkNumber(rpc_serviceQty2) & ", " & _
					"rpc_serviceAmt2=" & ChkNumber(rpc_serviceAmt2) & ", rpc_serviceQtyTotal=" & ChkNumber(rpc_serviceQtyTotal) & ", rpc_serviceAmtTotal=" & ChkNumber(rpc_serviceAmtTotal) & ",  " & _
					"rpc_techfees=" & ChkNumber(rpc_techfees) & ", rpc_car_allow=" & ChkNumber(rpc_car_allow) & ", rpc_phone_allow=" & ChkNumber(rpc_phone_allow) & ",  " & _
					"rpc_toll=" & ChkNumber(rpc_toll) & ", rpc_parking=" & ChkNumber(rpc_parking) & ", rpc_petrol=" & ChkNumber(rpc_petrol) & ",  " & _
					"rpc_hotel=" & ChkNumber(rpc_hotel) & ", rpc_service_allow=" & ChkNumber(rpc_service_allow) & ", rpc_overwarranty_qty=" & ChkNumber(rpc_overwarranty_qty) & ",  " & _
					"rpc_overwarranty_fee=" & ChkNumber(rpc_overwarranty_fee) & ", rpc_others=" & ChkNumber(rpc_others) & ", rpc_deduction_ow_qty=" & ChkNumber(rpc_deduction_ow_qty) & ",  " & _
					"rpc_deduction_ow=" & ChkNumber(rpc_deduction_ow) & ",rpc_deduction_sparepart_qty=" & ChkNumber(rpc_deduction_sparepart_qty) & ",  " & _
					"rpc_deduction_sparepart=" & ChkNumber(rpc_deduction_sparepart) & ", rpc_deduction_total=" & ChkNumber(rpc_deduction_total) & ", rpc_total=" & ChkNumber(rpc_total) & ", " & _
					"rpc_deduction_desc1= '" & rpc_deduction_desc1 & "', rpc_deduction_desc2= '" & rpc_deduction_desc2 & "', " & _
					"rpc_others_desc= '" & rpc_others_desc & "',rpc_others_desc2= '" & rpc_others_desc2 & "',rpc_others2=" & ChkNumber(rpc_others2) & ", rpc_submitted_date = '" & ChkDateTimeMySQL(now()) & "', " & _			
					"rpc_water_storage_qty = " & ChkNumber(water_storage_qty) & " where rpc_id = " & rs1("rpc_id") 
	
	'if rs1("rpc_tech_code") = "RASC-D-001" then
	'	response.write sql4 
	'	response.write "/"
	'	response.write rpc_total
	'	response.end
	'end if

	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	

	url = "rm_rpt_tech_monthcommisionTPC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Report has been updated.#articletitle" 	  

'----------------------------------------------------------------------------------------------------    
  Case "inventorystock"
  
	orderby = request("orderby")
	ordertype = request("ordertype")

	if ordertype = "" then 
	   ordertype = "desc"
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
	sql = "Delete from tblrpr_techinventory where rpi_month=" & jobmonth & " and rpi_year=" & jobyear
	CUD(sql)
	sql = "INSERT INTO tblrpr_techinventory (rpi_month, rpi_year, rpi_item_code, rpi_item_name)  Select '" & jobmonth & "','" & jobyear & "', md_code, md_desc from tblmodel where md_category='Parts'"
	CUD(sql)
	
	dim rpi_tech_in(50), rpi_tech_out(50)
	
    ''' Loop tblrpr_techinventory
	sql1 = "SELECT rpi_id, rpi_month, rpi_year, rpi_item_code, rpi_item_name, rpi_tech01_in, rpi_tech01_out, rpi_tech02_in, rpi_tech02_out, " & _
	"rpi_tech03_in, rpi_tech03_out, rpi_tech04_in, rpi_tech04_out, rpi_tech05_in, rpi_tech05_out, rpi_tech06_in, rpi_tech06_out, rpi_tech07_in, rpi_tech07_out, rpi_tech08_in, rpi_tech08_out,  " & _
	"rpi_tech09_in, rpi_tech09_out, rpi_tech10_in, rpi_tech10_out, rpi_tech11_in, rpi_tech11_out, rpi_tech12_in, rpi_tech12_out, rpi_tech13_in, rpi_tech13_out,  " & _
	"rpi_tech14_in, rpi_tech14_out, rpi_tech15_in, rpi_tech15_out, rpi_tech16_in, rpi_tech16_out, rpi_tech17_in, rpi_tech17_out, rpi_tech18_in, rpi_tech18_out,  " & _
	"rpi_tech19_in, rpi_tech19_out, rpi_tech20_in, rpi_tech20_out, rpi_tech21_in, rpi_tech21_out, rpi_tech22_in, rpi_tech22_out, rpi_tech23_in, rpi_tech23_out,  " & _
	"rpi_tech24_in, rpi_tech24_out, rpi_tech25_in, rpi_tech25_out, rpi_tech26_in, rpi_tech26_out, rpi_tech27_in, rpi_tech27_out, rpi_tech28_in, rpi_tech28_out,  " & _
	"rpi_tech29_in, rpi_tech29_out, rpi_tech30_in, rpi_tech30_out, rpi_total_in, rpi_total_out " & _
	"FROM tblrpr_techinventory where rpi_month=" & jobmonth & " and rpi_year=" & jobyear & " order by rpi_id "
	set rs1 = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 
           
		    i = 1
		    sql2 = "select wh_code from tblwarehouse order by wh_code"
			rs2.ActiveConnection = strconnect
			rs2.Source = sql2
			rs2.CursorLocation  = 3
			rs2.Open
			while not rs2.eof 
			        'rpi_tech_in
					sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
							"where month(stk_date) = " & jobmonth & " and year(stk_date)= " & jobyear & " " & _
							"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
							"and stk_reference='" & rs2("wh_code") & "' " & _ 
							"and stk_qty > 0" 
	                rpi_tech_in(i) = selectid(sql2)
					
					'rpi_tech_out
					sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
							"where month(stk_date) = " & jobmonth & " and year(stk_date)= " & jobyear & " " & _
							"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
							"and stk_reference='" & rs2("wh_code") & "' " & _ 
							"and stk_qty < 0" 
	                temp = selectid(sql2)
					rpi_tech_out(i) = selectid(sql2)
					i = i + 1
			rs2.movenext
			wend
			rs2.close
		
			
	rpi_total_in = 0
	rpi_total_out = 0
	for i = 1 to 30 
		if rpi_tech_in(i) > 0 then
		   rpi_total_in = rpi_total_in + cint(rpi_tech_in(i))
		end if
		
		
		if rpi_tech_out(i) < 0  then
		   rpi_total_out = rpi_total_out + cint(rpi_tech_out(i))
		end if
		
		'response.write i & ": " & rpi_tech_out(i) & "<br>"
		
	next
	
	sql4 = "UPDATE tblrpr_techinventory SET " & _
	    "rpi_tech01_in=" & ChkNumber(rpi_tech_in(1)) & ", rpi_tech01_out=" & ChkNumber(rpi_tech_out(1)) & ", rpi_tech02_in=" & ChkNumber(rpi_tech_in(2)) & ", rpi_tech02_out=" & ChkNumber(rpi_tech_out(2)) & ", " & _
		"rpi_tech03_in=" & ChkNumber(rpi_tech_in(3)) & ", rpi_tech03_out=" & ChkNumber(rpi_tech_out(3)) & ", rpi_tech04_in=" & ChkNumber(rpi_tech_in(4)) & ", rpi_tech04_out=" & ChkNumber(rpi_tech_out(4)) & ", " & _
		"rpi_tech05_in=" & ChkNumber(rpi_tech_in(5)) & ", rpi_tech05_out=" & ChkNumber(rpi_tech_out(5)) & ", " & _
		"rpi_tech06_in=" & ChkNumber(rpi_tech_in(6)) & ", rpi_tech06_out=" & ChkNumber(rpi_tech_out(6)) & ", rpi_tech07_in=" & ChkNumber(rpi_tech_in(7)) & ", rpi_tech07_out=" & ChkNumber(rpi_tech_out(7)) & ", " & _
		"rpi_tech08_in=" & ChkNumber(rpi_tech_in(8)) & ", rpi_tech08_out=" & ChkNumber(rpi_tech_out(8)) & ", rpi_tech09_in=" & ChkNumber(rpi_tech_in(9)) & ", rpi_tech09_out=" & ChkNumber(rpi_tech_out(9)) & ", " & _
		"rpi_tech10_in=" & ChkNumber(rpi_tech_in(10)) & ", rpi_tech10_out=" & ChkNumber(rpi_tech_out(10)) & ", rpi_tech11_in=" & ChkNumber(rpi_tech_in(11)) & ", rpi_tech11_out=" & ChkNumber(rpi_tech_out(11)) & ", " & _ 
		"rpi_tech12_in=" & ChkNumber(rpi_tech_in(12)) & ", rpi_tech12_out=" & ChkNumber(rpi_tech_out(12)) & ", rpi_tech13_in=" & ChkNumber(rpi_tech_in(13)) & ", rpi_tech13_out=" & ChkNumber(rpi_tech_out(13)) & ", " & _
		"rpi_tech14_in=" & ChkNumber(rpi_tech_in(14)) & ", rpi_tech14_out=" & ChkNumber(rpi_tech_out(14)) & ", rpi_tech15_in=" & ChkNumber(rpi_tech_in(15)) & ", rpi_tech15_out=" & ChkNumber(rpi_tech_out(15)) & ", " & _
		"rpi_tech16_in=" & ChkNumber(rpi_tech_in(16)) & ", rpi_tech16_out=" & ChkNumber(rpi_tech_out(16)) & ", rpi_tech17_in=" & ChkNumber(rpi_tech_in(17)) & ", rpi_tech17_out=" & ChkNumber(rpi_tech_out(17)) & ", " & _
		"rpi_tech18_in=" & ChkNumber(rpi_tech_in(18)) & ", rpi_tech18_out=" & ChkNumber(rpi_tech_out(18)) & ", rpi_tech19_in=" & ChkNumber(rpi_tech_in(19)) & ", rpi_tech19_out=" & ChkNumber(rpi_tech_out(19)) & ", " & _
		"rpi_tech20_in=" & ChkNumber(rpi_tech_in(20)) & ", rpi_tech20_out=" & ChkNumber(rpi_tech_out(20)) & ", rpi_tech21_in=" & ChkNumber(rpi_tech_in(21)) & ", rpi_tech21_out=" & ChkNumber(rpi_tech_out(21)) & ", " & _
		"rpi_tech22_in=" & ChkNumber(rpi_tech_in(22)) & ", rpi_tech22_out=" & ChkNumber(rpi_tech_out(22)) & ", rpi_tech23_in=" & ChkNumber(rpi_tech_in(23)) & ", rpi_tech23_out=" & ChkNumber(rpi_tech_out(23)) & ", " & _
		"rpi_tech24_in=" & ChkNumber(rpi_tech_in(24)) & ", rpi_tech24_out=" & ChkNumber(rpi_tech_out(24)) & ", rpi_tech25_in=" & ChkNumber(rpi_tech_in(25)) & ", rpi_tech25_out=" & ChkNumber(rpi_tech_out(25)) & ", " & _
		"rpi_tech26_in=" & ChkNumber(rpi_tech_in(26)) & ", rpi_tech26_out=" & ChkNumber(rpi_tech_out(26)) & ", rpi_tech27_in=" & ChkNumber(rpi_tech_in(27)) & ", rpi_tech27_out=" & ChkNumber(rpi_tech_out(27)) & ", " & _
		"rpi_tech28_in=" & ChkNumber(rpi_tech_in(28)) & ", rpi_tech28_out=" & ChkNumber(rpi_tech_out(28)) & ", rpi_tech29_in=" & ChkNumber(rpi_tech_in(29)) & ", rpi_tech29_out=" & ChkNumber(rpi_tech_out(29)) & ", " & _
		"rpi_tech30_in=" & ChkNumber(rpi_tech_in(30)) & ", rpi_tech30_out=" & ChkNumber(rpi_tech_out(30)) & ", rpi_total_in=" & ChkNumber(rpi_total_in) & ", rpi_total_out=" & ChkNumber(rpi_total_out) & " " & _
	    "WHERE rpi_id = " & rs1("rpi_id") 
		
		'response.write sql4 & "<br>"
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
	url = "rm_rpt_inventory_shockin.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&orderby=" & orderby & "&loginerr=Report has been updated.#articletitle" 	

'----------------------------------------------------------------------------------------------------    
  Case "inventoryWarehouseLocation"
	 Server.ScriptTimeout=10000
	'On Error Resume Next
  
	orderby = request("orderby")
	ordertype = request("ordertype")

	if ordertype = "" then 
	   ordertype = "desc"
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
	sql = "Delete from tblrpr_techinventory where rpi_month=" & jobmonth & " and rpi_year=" & jobyear
	CUD(sql)
	sql = "INSERT INTO tblrpr_techinventory (rpi_month, rpi_year, rpi_item_code, rpi_item_name)  Select '" & jobmonth & "','" & jobyear & "', md_code, md_desc from tblmodel where md_category='Parts'"
	CUD(sql)
	
	'redim rpi_tech_in(30), rpi_tech_out(30)
	redim rpi_tech_in(100), rpi_tech_out(100)
	
    ''' Loop tblrpr_techinventory
	sql1 = "SELECT rpi_id, rpi_month, rpi_year, rpi_item_code, rpi_item_name, rpi_tech01_in, rpi_tech01_out, rpi_tech02_in, rpi_tech02_out, " & _
	"rpi_tech03_in, rpi_tech03_out, rpi_tech04_in, rpi_tech04_out, rpi_tech05_in, rpi_tech05_out, rpi_tech06_in, rpi_tech06_out, rpi_tech07_in, rpi_tech07_out, rpi_tech08_in, rpi_tech08_out,  " & _
	"rpi_tech09_in, rpi_tech09_out, rpi_tech10_in, rpi_tech10_out, rpi_tech11_in, rpi_tech11_out, rpi_tech12_in, rpi_tech12_out, rpi_tech13_in, rpi_tech13_out,  " & _
	"rpi_tech14_in, rpi_tech14_out, rpi_tech15_in, rpi_tech15_out, rpi_tech16_in, rpi_tech16_out, rpi_tech17_in, rpi_tech17_out, rpi_tech18_in, rpi_tech18_out,  " & _
	"rpi_tech19_in, rpi_tech19_out, rpi_tech20_in, rpi_tech20_out, rpi_tech21_in, rpi_tech21_out, rpi_tech22_in, rpi_tech22_out, rpi_tech23_in, rpi_tech23_out,  " & _
	"rpi_tech24_in, rpi_tech24_out, rpi_tech25_in, rpi_tech25_out, rpi_tech26_in, rpi_tech26_out, rpi_tech27_in, rpi_tech27_out, rpi_tech28_in, rpi_tech28_out,  " & _
	"rpi_tech29_in, rpi_tech29_out, rpi_tech30_in, rpi_tech30_out, rpi_total_in, rpi_total_out " & _
	"FROM tblrpr_techinventory where rpi_month=" & jobmonth & " and rpi_year=" & jobyear & " order by rpi_id "
	set rs1 = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	'response.write sql1

	while not rs1.eof 
           
		    i = 1
		    sql2 = "select wh_code from tblwarehouse order by wh_code"
			rs2.ActiveConnection = strconnect
			rs2.Source = sql2
			rs2.CursorLocation  = 3
			rs2.Open
			while not rs2.eof 
	
					'rpi_tech_in
					sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
							"where month(stk_date) = " & jobmonth & " and year(stk_date)= " & jobyear & " " & _
							"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
							"and stk_reference='" & rs2("wh_code") & "' " & _ 
							"and stk_qty > 0" 
	                
						rpi_tech_in(i) = selectid(sql2)
		'response.write rpi_tech_in(i)
	'response.write " // "
					'response.write sql2
					'response.write i

					'rpi_tech_out
					sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
							"where month(stk_date) = " & jobmonth & " and year(stk_date)= " & jobyear & " " & _
							"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
							"and stk_reference='" & rs2("wh_code") & "' " & _ 
							"and stk_qty < 0" 
	                temp = selectid(sql2)
					rpi_tech_out(i) = selectid(sql2)
					i = i + 1
	
			rs2.movenext
			wend
			rs2.close
					
	rpi_total_in = 0
	rpi_total_out = 0
	for i = 1 to 30 
		if rpi_tech_in(i) > 0 then
		   rpi_total_in = rpi_total_in + cint(rpi_tech_in(i))
		end if
		
		
		if rpi_tech_out(i) < 0  then
		   rpi_total_out = rpi_total_out + cint(rpi_tech_out(i))
		end if
		
		'response.write i & ": " & rpi_tech_out(i) & "<br>"
		
	next
	
	sql4 = "UPDATE tblrpr_techinventory SET " & _
	    "rpi_tech01_in=" & ChkNumber(rpi_tech_in(1)) & ", rpi_tech01_out=" & ChkNumber(rpi_tech_out(1)) & ", rpi_tech02_in=" & ChkNumber(rpi_tech_in(2)) & ", rpi_tech02_out=" & ChkNumber(rpi_tech_out(2)) & ", " & _
		"rpi_tech03_in=" & ChkNumber(rpi_tech_in(3)) & ", rpi_tech03_out=" & ChkNumber(rpi_tech_out(3)) & ", rpi_tech04_in=" & ChkNumber(rpi_tech_in(4)) & ", rpi_tech04_out=" & ChkNumber(rpi_tech_out(4)) & ", " & _
		"rpi_tech05_in=" & ChkNumber(rpi_tech_in(5)) & ", rpi_tech05_out=" & ChkNumber(rpi_tech_out(5)) & ", " & _
		"rpi_tech06_in=" & ChkNumber(rpi_tech_in(6)) & ", rpi_tech06_out=" & ChkNumber(rpi_tech_out(6)) & ", rpi_tech07_in=" & ChkNumber(rpi_tech_in(7)) & ", rpi_tech07_out=" & ChkNumber(rpi_tech_out(7)) & ", " & _
		"rpi_tech08_in=" & ChkNumber(rpi_tech_in(8)) & ", rpi_tech08_out=" & ChkNumber(rpi_tech_out(8)) & ", rpi_tech09_in=" & ChkNumber(rpi_tech_in(9)) & ", rpi_tech09_out=" & ChkNumber(rpi_tech_out(9)) & ", " & _
		"rpi_tech10_in=" & ChkNumber(rpi_tech_in(10)) & ", rpi_tech10_out=" & ChkNumber(rpi_tech_out(10)) & ", rpi_tech11_in=" & ChkNumber(rpi_tech_in(11)) & ", rpi_tech11_out=" & ChkNumber(rpi_tech_out(11)) & ", " & _ 
		"rpi_tech12_in=" & ChkNumber(rpi_tech_in(12)) & ", rpi_tech12_out=" & ChkNumber(rpi_tech_out(12)) & ", rpi_tech13_in=" & ChkNumber(rpi_tech_in(13)) & ", rpi_tech13_out=" & ChkNumber(rpi_tech_out(13)) & ", " & _
		"rpi_tech14_in=" & ChkNumber(rpi_tech_in(14)) & ", rpi_tech14_out=" & ChkNumber(rpi_tech_out(14)) & ", rpi_tech15_in=" & ChkNumber(rpi_tech_in(15)) & ", rpi_tech15_out=" & ChkNumber(rpi_tech_out(15)) & ", " & _
		"rpi_tech16_in=" & ChkNumber(rpi_tech_in(16)) & ", rpi_tech16_out=" & ChkNumber(rpi_tech_out(16)) & ", rpi_tech17_in=" & ChkNumber(rpi_tech_in(17)) & ", rpi_tech17_out=" & ChkNumber(rpi_tech_out(17)) & ", " & _
		"rpi_tech18_in=" & ChkNumber(rpi_tech_in(18)) & ", rpi_tech18_out=" & ChkNumber(rpi_tech_out(18)) & ", rpi_tech19_in=" & ChkNumber(rpi_tech_in(19)) & ", rpi_tech19_out=" & ChkNumber(rpi_tech_out(19)) & ", " & _
		"rpi_tech20_in=" & ChkNumber(rpi_tech_in(20)) & ", rpi_tech20_out=" & ChkNumber(rpi_tech_out(20)) & ", rpi_tech21_in=" & ChkNumber(rpi_tech_in(21)) & ", rpi_tech21_out=" & ChkNumber(rpi_tech_out(21)) & ", " & _
		"rpi_tech22_in=" & ChkNumber(rpi_tech_in(22)) & ", rpi_tech22_out=" & ChkNumber(rpi_tech_out(22)) & ", rpi_tech23_in=" & ChkNumber(rpi_tech_in(23)) & ", rpi_tech23_out=" & ChkNumber(rpi_tech_out(23)) & ", " & _
		"rpi_tech24_in=" & ChkNumber(rpi_tech_in(24)) & ", rpi_tech24_out=" & ChkNumber(rpi_tech_out(24)) & ", rpi_tech25_in=" & ChkNumber(rpi_tech_in(25)) & ", rpi_tech25_out=" & ChkNumber(rpi_tech_out(25)) & ", " & _
		"rpi_tech26_in=" & ChkNumber(rpi_tech_in(26)) & ", rpi_tech26_out=" & ChkNumber(rpi_tech_out(26)) & ", rpi_tech27_in=" & ChkNumber(rpi_tech_in(27)) & ", rpi_tech27_out=" & ChkNumber(rpi_tech_out(27)) & ", " & _
		"rpi_tech28_in=" & ChkNumber(rpi_tech_in(28)) & ", rpi_tech28_out=" & ChkNumber(rpi_tech_out(28)) & ", rpi_tech29_in=" & ChkNumber(rpi_tech_in(29)) & ", rpi_tech29_out=" & ChkNumber(rpi_tech_out(29)) & ", " & _
		"rpi_tech30_in=" & ChkNumber(rpi_tech_in(30)) & ", rpi_tech30_out=" & ChkNumber(rpi_tech_out(30)) & ", rpi_total_in=" & ChkNumber(rpi_total_in) & ", rpi_total_out=" & ChkNumber(rpi_total_out) & " " & _
	    "WHERE rpi_id = " & rs1("rpi_id") 
		
		'response.write sql4 & "<br>"
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
	url = "rm_rpt_inventory_warehouse_location.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&orderby=" & orderby & "&loginerr=Report has been updated.#articletitle" 	

'----------------------------------------------------------------------------------------------------    
  Case "inventorystock_single"
  
	orderby = request("orderby")
	ordertype = request("ordertype")
	wh_code = request("wh_code")
	searchitem = request("searchitem")
	searchvalue = request("searchvalue")

	if ordertype = "" then 
	   ordertype = "desc"
	end if
	
	if request("job_date_from") <> "" then
	   job_date_from = request("job_date_from")
	else
	   job_date_from = chkdate(DateAdd("d",-90,date()))
	end if
	
	if request("job_date_to") <> "" then
	   job_date_to = chkdate(DateAdd("d",1,request("job_date_to"))) 
	else
	   job_date_to = chkdate(DateAdd("d",1,date()))
	end if
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_techinventory_single"
	CUD(sql)
	'sql = "INSERT INTO tblrpr_techinventory_single (rpi_item_code, rpi_item_name)  select  stk_itm_code, md_desc from v_tblstocktran_item"
	'CUD(sql)
	
	
	sql = "INSERT INTO tblrpr_techinventory_single (rpi_item_code, rpi_item_name, rpi_total_bal) Select  tblmodel.md_code, tblmodel.md_desc, " & _
		"sum(tblstocktran.stk_qty) Totalstockqty " & _
		"from tblstocktran inner join tblmodel " & _
		"on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblstocktran.stk_reference ='" & wh_code & "' and tblstocktran.stk_reference <> '0' " & _
		"group by tblmodel.md_code, tblmodel.md_desc "
	CUD(sql)	
		
    ''' Loop tblrpr_techinventory
	sql1 = "SELECT rpi_id, rpi_month, rpi_year, rpi_item_code, rpi_item_name, rpi_total_in, rpi_total_out " & _
	       "FROM tblrpr_techinventory_single order by rpi_id "
	set rs1 = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 
		  
		'rpi_tech_in
		sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
				"where stk_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and stk_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
				"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
				"and stk_reference='" & wh_code & "' " & _ 
				"and stk_qty > 0" 
		rpi_tech_in = selectid(sql2)
		
		'rpi_tech_out
		sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
				"where stk_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and stk_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
				"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
				"and stk_reference='" & wh_code & "' " & _ 
				"and stk_qty < 0" 
		temp = selectid(sql2)
		rpi_tech_out = selectid(sql2)
			

	sql4 = "UPDATE tblrpr_techinventory_single SET " & _
	       "rpi_total_in=" & ChkNumber(rpi_tech_in) & ", rpi_total_out=" & ChkNumber(rpi_tech_out) & " " & _
	       "WHERE rpi_id = " & rs1("rpi_id") 
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
	url = "rm_rpt_inventory_warehouse.asp?type=showresult&job_date_from=" & request("job_date_from") & "&job_date_to=" & request("job_date_to") & "&orderby=" & orderby & "&wh_code=" & wh_code & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&loginerr=Report has been updated.#articletitle" 	

'----------------------------------------------------------------------------------------------------    
  Case "monthPnl"
  
	orderby = request("orderby")
	ordertype = request("ordertype")

	if ordertype = "" then 
	   ordertype = "desc"
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
	sql = "Delete from tblrpr_pnl where pnl_month=" & jobmonth & " and pnl_year=" & jobyear
	CUD(sql)
	sql = "INSERT INTO tblrpr_pnl (pnl_month, pnl_year) values (" & jobmonth & ", " & jobyear & ")"
	CUD(sql)
		
	'pnl_totalinvoice_qty
	sql2 = "SELECT count(inv_id) as pnl_totalinvoice_qty FROM tblinvoice " & _
			"where month(inv_date) = " & jobmonth & " and year(inv_date)= " & jobyear & " " & _
			"and inv_status='Posted'" 
	pnl_totalinvoice_qty = selectid(sql2)
	
	'inv_totalAmt
    sql2 = "SELECT sum(inv_totalAmt) as pnl_totalinvoice_amt FROM tblinvoice " & _
			"where month(inv_date) = " & jobmonth & " and year(inv_date)= " & jobyear & " " & _
			"and inv_status='Posted'" 
	pnl_totalinvoice_amt = selectid(sql2)
	
	'pnl_totalOutputGST
    sql2 = "SELECT sum(inv_gstAmt) as pnl_totalOutputGST FROM tblinvoice " & _
			"where month(inv_date) = " & jobmonth & " and year(inv_date)= " & jobyear & " " & _
			"and inv_status='Posted'" 
	pnl_totalOutputGST = selectid(sql2)
	
	'pnl_sparepartcost_amt
    sql2 = "SELECT sum(tblinvoice_detail.invd_subtotal) as pnl_sparepartcost_amt FROM tblinvoice_detail inner join tblinvoice " & _
	        "on tblinvoice_detail.invd_inv_no=tblinvoice.inv_no " & _
			"where month(tblinvoice.inv_date) = " & jobmonth & " and year(tblinvoice.inv_date)= " & jobyear & " " & _
			"and tblinvoice.inv_status='Posted' and tblinvoice_detail.invd_parttype='Parts'" 
	pnl_sparepartcost_amt = selectid(sql2)
	
   'pnl_labour_amt
    sql2 = "SELECT sum(tblinvoice_detail.invd_subtotal) as pnl_labour_amt FROM tblinvoice_detail inner join tblinvoice " & _
	        "on tblinvoice_detail.invd_inv_no=tblinvoice.inv_no " & _
			"where month(tblinvoice.inv_date) = " & jobmonth & " and year(tblinvoice.inv_date)= " & jobyear & " " & _
			"and tblinvoice.inv_status='Posted' and tblinvoice_detail.invd_parttype='Labour'" 
	pnl_labour_amt = selectid(sql2)
	
	'pnl_transport_amt
    sql2 = "SELECT sum(tblinvoice_detail.invd_subtotal) as pnl_transport_amt FROM tblinvoice_detail inner join tblinvoice " & _
	        "on tblinvoice_detail.invd_inv_no=tblinvoice.inv_no " & _
			"where month(tblinvoice.inv_date) = " & jobmonth & " and year(tblinvoice.inv_date)= " & jobyear & " " & _
			"and tblinvoice.inv_status='Posted' and tblinvoice_detail.invd_parttype='Transport'" 
	pnl_transport_amt = selectid(sql2)

	'pnl_commission_amt
    sql2 = "SELECT sum(rpc_total) as pnl_commission_amt FROM tblrpr_techcommission " & _
			"where rpc_month = " & jobmonth & " and rpc_year = " & jobyear & " "
	pnl_commission_amt = selectid(sql2)	
	
			
	pnl_net_amt = pnl_totalinvoice_amt - pnl_sparepartcost_amt - pnl_commission_amt
	
	sql4 = "Update tblrpr_pnl set pnl_totalinvoice_qty=" & ChkNumber(pnl_totalinvoice_qty) & ", pnl_totalinvoice_amt=" & ChkNumber(pnl_totalinvoice_amt) & ", pnl_totalOutputGST=" & ChkNumber(pnl_totalOutputGST) & ", pnl_sparepartcost_amt=" & ChkNumber(pnl_sparepartcost_amt) & ", " & _
		   "pnl_labour_amt=" & ChkNumber(pnl_labour_amt) & ", pnl_transport_amt=" & ChkNumber(pnl_transport_amt) & ", pnl_commission_amt=" & ChkNumber(pnl_commission_amt) & ", " & _
		   "pnl_net_amt=" & ChkNumber(pnl_net_amt) & " where pnl_month=" & jobmonth & " and pnl_year=" & jobyear & " "
	CUD(sql4)
	
	url = "rm_rpt_pnL.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&orderby=" & orderby & "&loginerr=Report has been updated.#articletitle" 	


'----------------------------------------------------------------------------------------------------    
  Case "warehouselocation"
  
	orderby = request("orderby")
	ordertype = request("ordertype")

	if ordertype = "" then 
	   ordertype = "desc"
	end if
	
	if request("job_date_from") <> "" then
	   job_date_from = request("job_date_from")
	else
	   job_date_from = chkdate(DateAdd("d",-90,date()))
	end if
	
	if request("job_date_to") <> "" then
	   job_date_to = chkdate(DateAdd("d",1,request("job_date_to"))) 
	else
	   job_date_to = chkdate(date())
	end if
	
	'''Generate tblrpr_farmonth table.
	sql = "Delete from tblrpr_warehouselocation "
	CUD(sql)
	sql = "INSERT INTO tblrpr_warehouselocation (rpi_item_code, rpi_item_name) " & _
	      "Select md_code, md_desc from tblmodel where md_category='Parts'"
	CUD(sql)
	
	dim rpi_tech(30)
    ''' Loop tblrpr_techinventory
	sql1 = "SELECT rpi_id, rpi_item_code, rpi_item_name, rpi_tech01_qty, rpi_tech02_qty, rpi_tech03_qty, rpi_tech04_qty, " & _
		"rpi_tech05_qty, rpi_tech06_qty, rpi_tech07_qty, rpi_tech08_qty, rpi_tech09_qty, rpi_tech10_qty, rpi_tech11_qty, " & _ 
		"rpi_tech12_qty, rpi_tech13_qty, rpi_tech14_qty, rpi_tech15_qty, rpi_tech16_qty, rpi_tech17_qty, rpi_tech18_qty, " & _ 
		"rpi_tech19_qty, rpi_tech20_qty, rpi_tech21_qty, rpi_tech22_qty, rpi_tech23_qty, rpi_tech24_qty, rpi_tech25_qty, " & _ 
		"rpi_tech26_qty, rpi_tech27_qty, rpi_tech28_qty, rpi_tech29_qty, rpi_tech30_qty, rpi_total " & _
		"FROM tblrpr_warehouselocation order by rpi_id "
	set rs1 = server.CreateObject("adodb.recordset")
	set rs2 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 
           
		    i = 1
		    sql2 = "select wh_code from tblwarehouse order by wh_code"
			rs2.ActiveConnection = strconnect
			rs2.Source = sql2
			rs2.CursorLocation  = 3
			rs2.Open
			while not rs2.eof 
			        'rpi_tech_in
					sql2 = "SELECT sum(stk_qty) as stk_qty FROM tblstocktran " & _
							"where stk_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and stk_date < '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
							"and stk_reference is not null and stk_itm_code='" & rs1("rpi_item_code") & "' " & _
							"and stk_reference='" & rs2("wh_code") & "' " 
	                rpi_tech(i) = selectid(sql2)
					i = i + 1
			rs2.movenext
			wend
			rs2.close
		
			
	rpi_total = 0
	for i = 1 to 30 
    rpi_total = rpi_total + ChkNumber(rpi_tech(i))
	next
	
	sql4 = "UPDATE tblrpr_warehouselocation SET " & _
	"rpi_tech01_qty=" & ChkNumber(rpi_tech(1)) & ", " & _
	"rpi_tech02_qty=" & ChkNumber(rpi_tech(2)) & ", " & _
	"rpi_tech03_qty=" & ChkNumber(rpi_tech(3)) & ", " & _
	"rpi_tech04_qty=" & ChkNumber(rpi_tech(4)) & ", " & _
	"rpi_tech05_qty=" & ChkNumber(rpi_tech(5)) & ", " & _
	"rpi_tech06_qty=" & ChkNumber(rpi_tech(6)) & ", " & _
	"rpi_tech07_qty=" & ChkNumber(rpi_tech(7)) & ", " & _
	"rpi_tech08_qty=" & ChkNumber(rpi_tech(8)) & ", " & _
	"rpi_tech09_qty=" & ChkNumber(rpi_tech(9)) & ", " & _
	"rpi_tech10_qty=" & ChkNumber(rpi_tech(10)) & ", " & _
	"rpi_tech11_qty=" & ChkNumber(rpi_tech(11)) & ", " & _
	"rpi_tech12_qty=" & ChkNumber(rpi_tech(12)) & ", " & _
	"rpi_tech13_qty=" & ChkNumber(rpi_tech(13)) & ", " & _
	"rpi_tech14_qty=" & ChkNumber(rpi_tech(14)) & ", " & _
	"rpi_tech15_qty=" & ChkNumber(rpi_tech(15)) & ", " & _
	"rpi_tech16_qty=" & ChkNumber(rpi_tech(16)) & ", " & _
	"rpi_tech17_qty=" & ChkNumber(rpi_tech(17)) & ", " & _
	"rpi_tech18_qty=" & ChkNumber(rpi_tech(18)) & ", " & _
	"rpi_tech19_qty=" & ChkNumber(rpi_tech(19)) & ", " & _
	"rpi_tech20_qty=" & ChkNumber(rpi_tech(20)) & ", " & _
	"rpi_tech21_qty=" & ChkNumber(rpi_tech(21)) & ", " & _
	"rpi_tech22_qty=" & ChkNumber(rpi_tech(22)) & ", " & _
	"rpi_tech23_qty=" & ChkNumber(rpi_tech(23)) & ", " & _
	"rpi_tech24_qty=" & ChkNumber(rpi_tech(24)) & ", " & _
	"rpi_tech25_qty=" & ChkNumber(rpi_tech(25)) & ", " & _
	"rpi_tech26_qty=" & ChkNumber(rpi_tech(26)) & ", " & _
	"rpi_tech27_qty=" & ChkNumber(rpi_tech(27)) & ", " & _
	"rpi_tech28_qty=" & ChkNumber(rpi_tech(28)) & ", " & _
	"rpi_tech29_qty=" & ChkNumber(rpi_tech(29)) & ", " & _
	"rpi_tech30_qty=" & ChkNumber(rpi_tech(30)) & ", " & _
	"rpi_total=" & ChkNumber(rpi_total) & " WHERE rpi_id = " & rs1("rpi_id")
	CUD(sql4)

	rs1.movenext
	wend
	rs1.close
	
	url = "rm_rpt_stockcard_warehouse.asp?type=showresult&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&orderby=" & orderby & "&loginerr=Report has been updated.#articletitle" 	
	
'----------------------------------------------------------------------------------------------------    

		 
End Select
Response.Clear
Response.Redirect(url)
%>

