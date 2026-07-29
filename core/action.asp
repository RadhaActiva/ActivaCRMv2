<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<!-- #include file="ExcelADO.asp" -->
<%Dim act,url,i,var,img(2),Upload,file
act = Request("type")

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
   
  Case "addorder"   	

	if request.form("job_cust_cnty_id") = "" or request.form("job_cust_postcode") = "" then 'country or postcode cannot be empty
		job_cust_cnty_id=request.form("job_cust_cnty_id")
		job_cust_postcode=request("job_cust_postcode")
		job_cust_name=request("job_cust_name")
		job_cust_tel1=request("job_cust_tel1")
		job_cust_tel2=request("job_cust_tel2")
		job_cust_address=request("job_cust_address")
		Response.Redirect "rm_jobsheet.asp?job_code="&job_code&"&job_cust_postcode="&job_cust_postcode&"&job_cust_name="&job_cust_name&"&job_cust_address="&job_cust_address&"&job_cust_tel1="&job_cust_tel1&"&job_cust_tel2="&job_cust_tel2&"&job_cust_cnty_id="&job_cust_cnty_id& "&loginerr=Ensure Country and Postcode is entered.#articletitle" 
	end if

	'if request.form("job_cust_cnty_id") <> "" and request.form("job_cust_postcode") = "" then 'just return regardless country
	'	job_cust_cnty_id=request.form("job_cust_cnty_id")
	'	job_cust_postcode=request("job_cust_postcode")
	'	job_cust_name=request("job_cust_name")
	'	job_cust_tel1=request("job_cust_tel1")
	'	job_cust_tel2=request("job_cust_tel2")
	'	job_cust_address=request("job_cust_address")
	'	Response.Redirect "rm_jobsheet.asp?job_code="&job_code&"&job_cust_postcode="&job_cust_postcode&"&job_cust_name="&job_cust_name&"&job_cust_address="&job_cust_address&"&job_cust_tel1="&job_cust_tel1&"&job_cust_tel2="&job_cust_tel2&"&job_cust_cnty_id="&job_cust_cnty_id& "&loginerr=Enter Postcode.#articletitle" 
	'end if		

	if request.form("job_cust_postcode") <> ""  and request.form("job_cust_code") = ""  and request.form("job_cust_cnty_id") = "129" then
	'  if request.form("job_cust_city_id") = "" and request.form("job_cust_state_id") = "" then 'first time entering & creating auto-fill for state/city	
	  if request.form("job_cust_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			job_cust_postcode=request("job_cust_postcode")
			job_cust_name=request("job_cust_name")
			job_cust_tel1=request("job_cust_tel1")
			job_cust_tel2=request("job_cust_tel2")
			job_cust_address=request("job_cust_address")
			job_cust_cnty_id=request.form("job_cust_cnty_id")

			sql = "select state_id from tblpostcode where postcode =" & request("job_cust_postcode")	
			job_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("job_cust_postcode")	
			job_cust_state = selectid(sql)

			if Request.Cookies("GAPS")("slevel") = "technician2" then 
			   url = "tech2_jobsheet_new.asp?job_code=" & job_code & "&loginerr=Updated Address.#articletitle" 
			else
			   Response.Redirect "rm_jobsheet.asp?job_code="&job_code&"&job_cust_postcode="&job_cust_postcode&"&job_cust_name="&job_cust_name&"&job_cust_address="&job_cust_address&"&job_cust_tel1="&job_cust_tel1&"&job_cust_tel2="&job_cust_tel2&"&job_cust_cnty_id="&job_cust_cnty_id& "&loginerr=Updated Address.#articletitle" 
			end if
		end if
	end if

	if request.form("job_cust_cnty_id") = "129" then
		sql = "select state_code from tblstate where state_id =" & request("job_cust_state_id") 
		state_code = selectid(sql)

		if request.form("job_cust_code") <> "" then 'exiting customer select saved city
			sql ="select cust_city_id from tblcustomer where cust_code ='" & request.form("job_cust_code") & "'"
			job_cust_city_id = selectid(sql)		
		end if

		if request.form("job_cust_city_id") <> "" then
			sql = "select ct_name2 from tblcity where ct_id ='" &  request.form("job_cust_city_id") & "'" 
			job_cust_city = selectid(sql)	
			job_cust_city_id = request.form("job_cust_city_id")
		else	'this is for pre-selected customer from the customer selection
			job_cust_city=request.form("job_cust_city")
			job_cust_city_id= request.form("job_cust_city_code")
		end if
	end if
	
	if  request.form("job_cust_cnty_id") <> "129" then 
			job_cust_city = request.form("job_cust_city")
	   	    job_cust_city_id = "0"
	end if
		
	'if request.form("job_cust_cnty_id") = "129" and request.form("job_cust_city_id") = "" then '180125 to prevent from leaving city empty
	if request.form("job_cust_cnty_id") = "129" and job_cust_city_id = "" then '180125 to prevent from leaving city empty
		'response.write "<script>history.back();</script>"
		job_cust_cnty_id=request.form("job_cust_cnty_id")
		job_cust_postcode=request("job_cust_postcode")
		job_cust_name=request("job_cust_name")
		job_cust_tel1=request("job_cust_tel1")
		job_cust_tel2=request("job_cust_tel2")
		job_cust_address=request("job_cust_address")
		job_cust_code = request("job_cust_code")
		Response.Redirect "rm_jobsheet.asp?job_code="&job_code&"&job_cust_postcode="&job_cust_postcode&"&job_cust_name="&job_cust_name&"&job_cust_code="&job_cust_code&"&job_cust_address="&job_cust_address&"&job_cust_tel1="&job_cust_tel1&"&job_cust_tel2="&job_cust_tel2&"&job_cust_cnty_id="&job_cust_cnty_id& "&loginerr=Select City.#articletitle" 
	end if

       ''''Add Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email,job_cust_cnty_id, " & _
				"job_cust_tel1, job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_dealer, job_payee,job_dealer_inv " & _
				"FROM tbljob "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
	    'if ChkString(Request.Form("job_createdby")) <> "App" then
		rs.AddNew  
		'end if
	
		rs("job_status")  = "Open"   					
        rs("job_date") = ChkDateYYYYMMDD(date())
        rs("job_cust_code")  = ChkString(Request.Form("job_cust_code"))	
        rs("job_cust_name")  = ChkString(Request.Form("job_cust_name"))
		rs("job_cust_address") = ChkString(Request.Form("job_cust_address"))
		rs("job_cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
	
		if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
			rs("job_cust_state") = ChkString(Request.Form("job_cust_state"))	
			rs("job_cust_state_id") = ChkString(Request.Form("job_cust_state_id"))
		end if

		rs("job_cust_city") = job_cust_city			
		rs("job_cust_city_id") = job_cust_city_id	
		rs("job_cust_cnty_id") = request.form("job_cust_cnty_id")
		rs("job_cust_email") = ChkString(Request.Form("job_cust_email")) 
		rs("job_cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
		rs("job_cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
		rs("job_remark") = ChkString(Request.Form("job_remark")) 
		rs("job_createddate") = ChkDateTimeMySQL(now())
		rs("job_createdby") = Request.Cookies("GAPS")("sloginid")
		rs("job_JS_receiveddate") = ChkString(Request.Form("job_JS_receiveddate"))
		rs("job_JS_receivedby") = ChkString(Request.Form("job_JS_receivedby"))
		rs("job_onlineWrtyStatus") = ChkString(Request.Form("job_onlineWrtyStatus"))
		
		if ChkString(Request.Form("job_purchase_date")) <> "" then 
		rs("job_purchase_date") = ChkString(Request.Form("job_purchase_date"))
		end if 
			 
		rs("job_onlineWrtyNo") = ChkString(Request.Form("job_onlineWrtyNo")) 
		rs("job_type") = ChkString(Request.Form("job_type")) 
		rs("job_SN_no") = ChkString(Request.Form("job_SN_no")) 
		rs("job_Model") = ChkString(Request.Form("job_Model")) 
		rs("job_model_desc") = ChkString(Request.Form("job_model_desc")) 
		rs("job_faulty_reason_cs") = ChkString(Request.Form("job_faulty_reason_cs")) 
		rs("job_faulty_desc") = ChkString(Request.Form("job_faulty_desc")) 
		rs("job_reportedby") = ChkString(Request.Form("job_reportedby"))
		
		if ChkString(Request.Form("job_appointment_date")) <> "" then 
		rs("job_appointment_date") = ChkString(Request.Form("job_appointment_date")) 
		end if
		rs("job_appointment_time") = ChkString(Request.Form("job_appointment_time"))
		rs("job_tech_code") = ChkString(Request.Form("job_tech_code")) 
		rs("job_appointment_remark")  = ChkString(Request.Form("job_appointment_remark"))	
		rs("job_dealer") = ChkString(Request.Form("job_dealer")) 
		rs("job_payee") = ChkString(Request.Form("job_payee"))
		rs("job_dealer_inv")	=	ChkString(Request.Form("job_dealer_inv"))

		rs.Update 
		rs.Close      
		
        sql = "select top 1 job_id from tbljob order by job_id desc "
        job_id = selectid(sql)
		temp = ZeroPadLeft(job_id,6)
	
		if ChkString(Request.Form("job_faulty_reason_cs"))="Installation" then '041224 add I to job_no
			add_job_type = ChkString(Request.Form("job_type")) & "-I"
		else
			add_job_type = ChkString(Request.Form("job_type"))
		end if

		if request.form("job_cust_cnty_id") = "129" then 'for Malaysia
			job_code = state_code & "-" & temp & "-1-" & add_job_type 
		elseif request.form("job_cust_cnty_id") = "130" then 'for Singapore
			job_code = "SG" & "-" & temp & "-1-" & add_job_type 
		end if
  
        sql = "update tbljob set job_code = '" & job_code & "' where job_id = " & job_id
        CUD(sql)
		
		'--------------------------- Update Customer Profile --------------------------------------

        sql = "SELECT top 1 cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
			  "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
			  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic,cust_type,cust_debtor_code " & _
			  "FROM tblcustomer where cust_code = '" & ChkString(Request("job_cust_code")) & "' and cust_tel1 = '" & ChkString(Request.Form("job_cust_tel1"))  & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		   rs.addnew			
					if len(ChkString(Request.Form("job_cust_postcode"))) = "5" or len(ChkString(Request.Form("job_cust_postcode"))) = "6" then 
						rs("cust_name")  = ChkString(Request.Form("job_cust_name"))			
						rs("cust_address") = ChkString(Request.Form("job_cust_address"))
						rs("cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
			
						if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
							rs("cust_state") = ChkString(Request.Form("job_cust_state"))
							rs("cust_state_id") = ChkString(Request.Form("job_cust_state_id"))
						end if
		
						rs("cust_city") = job_cust_city
						rs("cust_city_id") = job_cust_city_id
						rs("cust_cnty_id") = request.form("job_cust_cnty_id")
						rs("cust_email") = ChkString(Request.Form("job_cust_email")) 
						rs("cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
						rs("cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
						rs("cust_status") = "Y"
						rs("cust_createddate") = ChkDateTimeMySQL(now())
						rs("cust_createdby") = Request.Cookies("GAPS")("sloginid")
						rs("cust_type") = "Customer" '080226 assumed entered data is customer
						rs("cust_debtor_code") = "300-CD10" '080226 assumed entered data is customer
						rs.Update 
				   end if
				   rs.Close 
				   
					sql = "select top 1 cust_id from tblcustomer order by cust_id desc "	
					cust_id = selectid(sql)
					temp = 100000+cust_id
					cust_code = "C" & temp
		
					sql = "Update tblcustomer set cust_code='" & cust_code & "' where cust_id=" & cust_id
					CUD(sql)
		
					sql = "Update tbljob set job_cust_code='" & cust_code & "' where job_id=" & job_id
					CUD(sql)
		end if

		if Request.Cookies("GAPS")("slevel") = "technician2" then 
		   url = "tech2_jobsheet_new.asp?job_code=" & job_code & "&loginerr=New Job has been created.#articletitle" 
		else
		   url = "rm_jobsheet.asp?job_code=" & job_code & "&loginerr=New Job has been created.#articletitle" 
		end if
      
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','addorder=" & ChkString(left(job_code,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editjob"   
	
	 if request.form("job_cust_postcode") <> "" and request.form("job_cust_cnty_id") = "129" then
			job_cust_postcode=request.form("job_cust_postcode")	

			sql = "select ct_name2 from tblcity where ct_id ='" &  request.form("job_cust_city_id") & "'" 
			job_cust_city = selectid(sql)	
	   	    job_cust_city_id=request.form("job_cust_city_id")

			sql = "select state_id from tblpostcode where postcode =" & job_cust_postcode
			job_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & job_cust_postcode 
			job_cust_state = selectid(sql)	
	 elseif request.form("job_cust_postcode") = "" and request.form("job_cust_cnty_id") = "129" then
			Response.Redirect "rm_jobsheet.asp?job_code="&request.form("job_code")&"&loginerr=Jobsheet Not Updated.#articletitle"
	 end if
	
	if  request.form("job_cust_cnty_id") = "129" then 
		sql = "select state_code from tblstate where state_id ='" & request("job_cust_state_id")  & "'" 
		state_code = selectid(sql)
	end if

	if  request.form("job_cust_cnty_id") <> "129" then 
			job_cust_city = request.form("job_cust_city")
	   	    job_cust_city_id = "0"
	end if
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id,job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_dealer,job_payee,job_dealer_inv " & _
				"FROM tbljob where job_id is not null "	
				
	    if request("job_code") <> "" then 
			sql = sql & " and job_code = '" & request("job_code") & "' "
		elseif request("job_id") <> "" then 
			sql = sql & " and job_id = " & request("job_id") & " "
		end if
	
	    set rs = server.CreateObject("adodb.recordset")
		rs.Open sql,strconnect,2,2,&H0001
        if not rs.eof then
			rs("job_cust_code")  = ChkString(Request.Form("job_cust_code"))	
			rs("job_cust_name")  = ChkString(Request.Form("job_cust_name"))
			rs("job_cust_address") = ChkString(Request.Form("job_cust_address"))
			rs("job_cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
			
			if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs("job_cust_state") = job_cust_state			
				rs("job_cust_state_id") = job_cust_state_id
			end if

			rs("job_cust_city") = job_cust_city		
			rs("job_cust_city_id") = job_cust_city_id
			rs("job_cust_cnty_id") = request.form("job_cust_cnty_id")
			rs("job_cust_email") = ChkString(Request.Form("job_cust_email")) 
			rs("job_cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
			rs("job_cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
			rs("job_remark") = ChkString(Request.Form("job_remark")) 
			rs("job_onlineWrtyStatus") = ChkString(Request.Form("job_onlineWrtyStatus"))
			
			if ChkString(Request.Form("job_purchase_date")) <> "" then 
		       rs("job_purchase_date") = ChkString(Request.Form("job_purchase_date"))
		    end if 

		  	rs("job_onlineWrtyNo") = ChkString(Request.Form("job_onlineWrtyNo")) 
			rs("job_type") = ChkString(Request.Form("job_type")) 
			rs("job_SN_no") = ChkString(Request.Form("job_SN_no")) 
			rs("job_Model") = ChkString(Request.Form("job_Model")) 
			rs("job_model_desc") = ChkString(Request.Form("job_model_desc")) 
			rs("job_faulty_reason_cs") = ChkString(Request.Form("job_faulty_reason_cs")) 
			rs("job_faulty_desc") = ChkString(Request.Form("job_faulty_desc")) 
			rs("job_reportedby") = ChkString(Request.Form("job_reportedby"))
			rs("job_createdby") = Request.Cookies("GAPS")("sloginid")

			if ChkString(Request.Form("job_appointment_date"))  <> "" then
			rs("job_appointment_date") = ChkString(Request.Form("job_appointment_date")) 
			end if
			
			rs("job_appointment_time") = ChkString(Request.Form("job_appointment_time"))
			rs("job_tech_code") = ChkString(Request.Form("job_tech_code")) 
			rs("job_appointment_remark")  = ChkString(Request.Form("job_appointment_remark"))	
			rs("job_dealer") = ChkString(Request.Form("job_dealer")) 
			rs("job_payee") = ChkString(Request.Form("job_payee"))
			rs("job_dealer_inv") = ChkString(Request.Form("job_dealer_inv"))

			job_id = rs("job_id")
			job_count = rs("job_count") 
			job_type = rs("job_type") 
		rs.Update 		 
		end if
		rs.Close		
	
		job_id_old = job_id
		temp = ZeroPadLeft(job_id_old,6)

		if ChkString(Request.Form("job_faulty_reason_cs"))="Installation" then '041224 add I to job_no
			add_job_type = job_type & "-I" '041224 for installation only job
		else
			add_job_type = job_type
		end if


		if request.form("job_cust_cnty_id") = "129" then 'for Malaysia
			job_code = state_code & "-" & temp & "-" & job_count & "-" & add_job_type
		elseif request.form("job_cust_cnty_id") = "130" then 'for Singapore
			job_code = "SG" & "-" & temp & "-" & job_count & "-" & add_job_type
		end if
		sql = "update tbljob set job_code = '" & job_code & "' where job_id = " & job_id
		CUD(sql)

		if ChkString(Request.Form("job_createdby")) = "App" then
			sql = "select top 1 cust_id from tblcustomer order by cust_id desc "	
			cust_id = selectid(sql)
			temp = 100000+cust_id
			cust_code = "C" & temp
	
			sql = "Update tbljob set job_cust_code='" & cust_code & "' where job_id=" & job_id
			CUD(sql)

		   sql = "SELECT top 1 cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
				  "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
				  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic,cust_type,cust_debtor_code " & _
				  "FROM tblcustomer"	
			set rs = server.CreateObject("adodb.recordset")
			rs.ActiveConnection = strconnect
			rs.Source = sql
			rs.CursorLocation  = 3
			rs.CursorType = 2
			rs.LockType = 2
			rs.Open	
			if ChkString(Request.Form("job_cust_code")) = "" then
				rs.addnew()
				if len(ChkString(Request.Form("job_cust_postcode"))) = "5" or len(ChkString(Request.Form("job_cust_postcode"))) = "6" then 
						rs("cust_code") = cust_code
						rs("cust_name")  = ChkString(Request.Form("job_cust_name"))
						rs("cust_address") = ChkString(Request.Form("job_cust_address"))
						rs("cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
						if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
							rs("cust_state") = job_cust_state
							rs("cust_state_id") = job_cust_state_id 
						end if
						rs("cust_city") = job_cust_city
						rs("cust_city_id") = job_cust_city_id 
						rs("cust_cnty_id") = request.form("job_cust_cnty_id")
						rs("cust_email") = ChkString(Request.Form("job_cust_email")) 
						rs("cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
						rs("cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
						rs("cust_status") = "Y"
						rs("cust_createddate") = ChkDateTimeMySQL(now())
						rs("cust_createdby") = Request.Cookies("GAPS")("sloginid")
						rs("cust_type") = "Customer" '080226 assumed entered data is customer
						rs("cust_debtor_code") = "300-CD10" '080226 assumed entered data is customer
						rs.Update 
					end if
				end if
			rs.Close 			
		else
			sql = "SELECT top 1 cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
				  "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
				  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic " & _
				  "FROM tblcustomer where cust_code = '" & ChkString(Request.Form("job_cust_code")) & "' "	
			set rs = server.CreateObject("adodb.recordset")
			rs.ActiveConnection = strconnect
			rs.Source = sql
			rs.CursorLocation  = 3
			rs.CursorType = 2
			rs.LockType = 2
			rs.Open		
			if not rs.eof then
			   if len(ChkString(Request.Form("job_cust_postcode"))) = "5" or len(ChkString(Request.Form("job_cust_postcode"))) = "6" then 
					rs("cust_code") = ChkString(Request.Form("job_cust_code"))
					rs("cust_name")  = ChkString(Request.Form("job_cust_name"))
					rs("cust_address") = ChkString(Request.Form("job_cust_address"))
					rs("cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
					if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
							rs("cust_state") = job_cust_state
							rs("cust_state_id") = job_cust_state_id 
					end if
					
					rs("cust_city") = job_cust_city
					rs("cust_city_id") = job_cust_city_id 
					rs("cust_cnty_id") = request.form("job_cust_cnty_id")
					rs("cust_email") = ChkString(Request.Form("job_cust_email")) 
					rs("cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
					rs("cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
					rs("cust_status") = "Y"
					rs("cust_createddate") = ChkDateTimeMySQL(now())
					rs("cust_createdby") = Request.Cookies("GAPS")("sloginid")
					rs.Update 
			    end if
			rs.Close 
			end if			
	    end if

		if Request.Cookies("GAPS")("slevel") = "technician2" then 
		   url = "tech2_jobsheet_new.asp?job_code=" & job_code & "&loginerr=New Job has been updated.#articletitle" 
		else
		    url = "rm_jobsheet.asp?job_code=" & job_code & "&loginerr=Job has been updated.#articletitle" 
		end if
				
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','editjob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	 

'----------------------------------------------------------------------------------------------------    
  Case "editjob_Technical"  
  
	set mySmartUpload = server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.Upload 
	
	tempid = mySmartUpload.form("job_wrty_photo")
	tempid2 = mySmartUpload.form("job_wrty_photo2")
	tempid3 = mySmartUpload.form("job_wrty_photo3")
	

		intCount = 1
		For each file In mySmartUpload.Files
		If not file.IsMissing Then   
			if file.name = "job_wrty_photo" then
				file.SaveAs(Server.MapPath(documentpath & tempid & file.FileName))  		
				job_wrty_photo = tempid & file.FileName
			end if
		
			if file.name = "job_wrty_photo2" then
				file.SaveAs(Server.MapPath(documentpath & tempid2 & file.FileName))  		
				job_wrty_photo2 = tempid2 & file.FileName
			end if
		
			if file.name = "job_wrty_photo3" then
				file.SaveAs(Server.MapPath(documentpath & tempid3 & file.FileName))  		
				job_wrty_photo3 = tempid3 & file.FileName
			end if		
		End if   
		intCount = intCount + 1     
		Next 
	
		sql = "select fr_description from tblfaultyreason where fr_code='" & ChkString(mySmartUpload.Form("job_tech_faulty_code")) & "' and fr_status='Y'" 
		job_tech_faulty_reason = selectid(sql)
		
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_code, job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2, job_wrty_photo3, job_tech_logby, job_tech_logdate, job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_doneby, job_donedate " & _
				"FROM tbljob where job_code = '" & mySmartUpload.Form("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
		
		    if ChkString(mySmartUpload.Form("job_appointment_date")) <> "" then  
			rs("job_appointment_date") = ChkString(mySmartUpload.Form("job_appointment_date"))
			end if
			
			if ChkString(mySmartUpload.Form("job_appointment_time")) <> "" then  
			rs("job_appointment_time") = ChkString(mySmartUpload.Form("job_appointment_time"))
			end if
			
			rs("job_tech_type") = ChkString(mySmartUpload.Form("job_tech_type"))
			rs("job_tech_model")  = ChkString(mySmartUpload.Form("job_tech_model"))	
			rs("job_tech_model_desc")  = ChkString(mySmartUpload.Form("job_tech_model_desc"))	
			rs("job_tech_tax_invoice")  = ChkString(mySmartUpload.Form("job_tech_tax_invoice"))
			rs("job_tech_SN") = ChkString(mySmartUpload.Form("job_tech_SN"))
			rs("job_tech_faulty_code") = ChkString(mySmartUpload.Form("job_tech_faulty_code"))
			rs("job_tech_faulty_reason") = ChkString(job_tech_faulty_reason)
			rs("job_tech_faulty_action") = ChkString(mySmartUpload.Form("job_tech_faulty_action")) 
			rs("job_tech_status") = ChkString(mySmartUpload.Form("job_tech_status")) 
			
			if ChkString(mySmartUpload.Form("job_tech_product_collectdate")) <> "" then 
			rs("job_tech_product_collectdate") = ChkString(mySmartUpload.Form("job_tech_product_collectdate")) 
			end if
			
			if ChkString(mySmartUpload.Form("job_tech_service_date")) <> "" then 
			rs("job_tech_service_date") = ChkString(mySmartUpload.Form("job_tech_service_date")) 
			end if
			
			if ChkString(mySmartUpload.Form("job_tech_returntoCustDate")) <> "" then 
			rs("job_tech_returntoCustDate") = ChkString(mySmartUpload.Form("job_tech_returntoCustDate")) 
			end if
			
			rs("job_actual_wrty_status") = ChkString(mySmartUpload.Form("job_actual_wrty_status"))
			
			if job_wrty_photo <> "" then  
				rs("job_wrty_photo") = ChkString(job_wrty_photo)
			end if

			if job_wrty_photo2 <> "" then  
				rs("job_wrty_photo2") = ChkString(job_wrty_photo2)
			end if

			if job_wrty_photo3 <> "" then  
				rs("job_wrty_photo3") = ChkString(job_wrty_photo3)
			end if

			rs("job_tech_logby") = Request.Cookies("GAPS")("sloginid")
			rs("job_return_tech_date")  = ChkDateTimeMySQL(now())	
			
		rs.Update 
		rs.Close 
		end if  
 
	 if request.Cookies("GAPS")("slevel") = "technician" then 
	 url = "rmtech_jobsheet.asp?job_code=" & mySmartUpload.Form("job_code") & "&loginerr=Job Technical Findings has been updated.#articletitle" 
	 elseif request.Cookies("GAPS")("slevel") = "technician2" then 
	 url = "tech2_jobsheet_new.asp?job_code=" & mySmartUpload.Form("job_code") & "&loginerr=Job Technical Findings has been updated.#articletitle" 
	 else
	 url = "rm_jobsheet.asp?job_code=" & mySmartUpload.Form("job_code") & "&loginerr=Job Technical Findings has been updated.#articletitle" 
	 end if
	
	 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','editjob=" & ChkString(left(mySmartUpload.Form("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	 CUD(sql)
 
 '----------------------------------------------------------------------------------------------------    
  Case "editjob_hq"  
  
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("job_hq_remark") = ChkString(Request.Form("job_hq_remark")) 

			if len(ChkString(Request.Form("job_hq_category_code"))) < 1  then 'added default code in case empty -  28/02/2024		
				rs("job_hq_category_code")  = "MD"
			else
				rs("job_hq_category_code")  = ChkString(Request.Form("job_hq_category_code"))	
			end if
			
			
			if ChkString(Request.Form("job_hq_received_date"))  <> "" then
			rs("job_hq_received_date")  = ChkString(Request.Form("job_hq_received_date"))
			end if
			
			if ChkString(Request.Form("job_repair_date"))  <> "" then
			rs("job_repair_date") = ChkString(Request.Form("job_repair_date"))
			end if
			
			if ChkString(Request.Form("job_return_tech_date"))  <> "" then
			rs("job_return_tech_date") = ChkString(Request.Form("job_return_tech_date"))
			end if
			
			if ChkString(Request.Form("job_tech_returntoCustDate"))  <> "" then
			rs("job_tech_returntoCustDate") = ChkString(Request.Form("job_tech_returntoCustDate"))
			end if
			
			rs("job_logbyhq") = Request.Cookies("GAPS")("sloginid")
			rs("job_logdatehq") = ChkDateTimeMySQL(now())			
			
		rs.Update 
		rs.Close 
		end if

		 if request.Cookies("GAPS")("slevel") = "technician2" then 
		 url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=Job HQ has been updated.#articletitle" 
		 else
		 url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=Job HQ has been updated.#articletitle" 
		 end if
	 
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','editjob_hq=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "editJob_officeuse"  
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_postedby, job_posteddate, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_faulty_reason_cs,job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			rs1("job_office_issueRemark") = ChkString(Request.Form("job_office_issueRemark"))
			rs1("job_office_supervisor")  = ChkString(Request.Cookies("GAPS")("sloginid"))	
			rs1("job_office_taxinvoice")  = ChkString(Request.Form("job_office_taxinvoice"))
			job_tech_code = rs1("job_tech_code")

			if ChkString(Request.Form("job_office_issueRemark")) = "Cancel" then 
				rs1("job_status") = "Cancel"
				rs1("job_cancelledby") = Request.Cookies("GAPS")("sloginid")
				rs1("job_cancelleddate") = ChkDateTimeMySQL(now())  
			elseif ChkString(Request.Form("job_office_issueRemark")) = "Posted" then 

			   '240624 - check each part for the job if the qty is suffcient at the warehouse
			
				sql="select tech_wh_code from tbltechnician where tech_code='" & job_tech_code & "'"
				wh_code = selectid(sql)
				
				sql2 = "SELECT tblmodel.md_category,  jobp_id, job_code, jobp_partcode, jobp_qty FROM tbljob_parts " & _ 
						"inner join tblmodel on tbljob_parts.jobp_partcode = tblmodel.md_code " & _
						"where job_code = '" & request("job_code") & "' order by jobp_id"

				set rs2 = server.CreateObject("adodb.recordset")
				rs2.Open sql2,strconnect,3,3,&H0001
				can_be_posted=true
				Service=False
				
				if not rs2.eof then '121024 this condition is needed in case no line items and jobs gets posted otherwise it throws an error
					if rs2("md_category") = "Service" then 
						Service=True
					end if
				end if

				if Service <> True then '021224 this is meant for installation service only, so dont check stock
						while Not rs2.EOF
							' checking stk qty in tech warehouse before allowing posting
							stk_qty=0
							sql = "select wst_itm_current_qty from tblwarehouse_stock where wst_wh_code='" & wh_code & "' and wst_itm_code = '" & rs2("jobp_partcode") & "'"				
							stk_qty = selectid(sql)
	
							if stk_qty < rs2("jobp_qty") then 'post job only if each time qty at tech warehouse is sufficient
									can_be_posted = false	
							end if										
						rs2.movenext
						wend
				end if
				rs2.close

				if can_be_posted = true then
					rs1("job_status") = "Posted"
					rs1("job_postedby") = Request.Cookies("GAPS")("sloginid")
					rs1("job_posteddate") = ChkDateTimeMySQL(now())  
				end if
			end if					
	
			rs1.Update 
			rs1.Close 
		end if
		
        if ChkString(Request.Form("job_office_issueRemark")) = "Posted"  and can_be_posted = true and Service <> True then 		
				'''Technician Warehouse
				'sql = "select wh_code from tblwarehouse where wh_contact_person='" & job_tech_code & "'"
				sql="select tech_wh_code from tbltechnician where tech_code='" & job_tech_code & "'"
				wh_code = selectid(sql)
				
				'''''Stock-Out Detail
				sql1 = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal " & _
							   "FROM tbljob_parts where job_code = '" & request("job_code") & "' order by jobp_id"	     
				'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				set rs2 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
				while Not rs1.EOF
						
					''''warehouse stock part qty deducted once job posted	   	  
					sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
						   "FROM tblwarehouse_stock where wst_wh_code = '" & wh_code & "' and wst_itm_code = '" & rs1("jobp_partcode") & "'"
					rs2.Open sql2,strconnect,2,2,&H0001
					if not rs2.eof then
						wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("jobp_qty") 
						rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("jobp_qty") 
						rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
						rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
						rs2.Update 

						sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
						Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostJob&DeductStk=" & rs1("job_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
						CUD(sql)
					end if
					rs2.Close    
				
					'Update Stocktrans - Stock Movement
					sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
						   "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
					rs2.Open sql2,strconnect,2,2,&H0001
					'the bug occurs when status change from posted-done-posted again. can be simulated
					'if rs2.eof then ' 19012024- originally no if condition, record appended each time users edits the job
						rs2.AddNew   
					'end if
					rs2("stk_voucherno") = request("job_code")
					rs2("stk_reference") = wh_code
					rs2("stk_date")  = ChkDateTimeMySQL(now())
					rs2("stk_type")  = "Job"
					rs2("stk_itm_code")  = ChkString(rs1("jobp_partcode"))
					rs2("stk_fromwarehouse")  = wh_code
					rs2("stk_towarehouse")  = so_towarehouse
					rs2("stk_desc")  = ChkString(rs1("jobp_desc"))
					rs2("stk_qty")  = ChkNumber(rs1("jobp_qty")*-1)
					rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
					rs2("stk_sales_price")  = ChkNumber(rs1("jobp_subtotal"))
					rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
					rs2("stk_logdate")  = ChkDateTimeMySQL(now())
					rs2.Update 
					rs2.Close  
					
				rs1.movenext
				wend
				rs1.close	
		end if	
	
		if ChkString(Request.Form("job_office_issueRemark")) = "Re-Assigned Job" then  
			
			'sql = "update tbljob set job_reassignedby='" & Request.Cookies("GAPS")("sloginid") & "', job_reassigndate='" & ChkDateTimeMySQL(now()) & "' where  job_code='" & request("job_code") & "'"
			'CUD(sql)
			
			'sql = "INSERT INTO tbljob_reassign  (SELECT  * FROM tbljob WHERE  job_code='" & request("job_code") & "') "
			'CUD(sql)
			
			'sql = "Update tbljob set job_appointment_date=null, job_appointment_time='', job_tech_code='', job_appointment_remark='', job_status='Open' where job_code = '" & request("job_code") & "'"
			'CUD(sql)

			
		elseif ChkString(Request.Form("job_office_issueRemark")) = "RCN" then  
		
		        ''''Generate RCN
		   sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, job_cust_tel1, " & _
				"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
				"job_type, job_SN_no, job_Model, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
				"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
				"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
				"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
				"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date " & _
				"FROM tbljob WHERE job_code = '" & request("job_code") & "' and (job_rcn_no = '' or job_rcn_no is null)"
			set rs = server.CreateObject("adodb.recordset")
			rs.Open sql,strconnect,0,1,&H0001
			If Not rs.EOF Then
				job_id = rs("job_id") 
				job_code = rs("job_code") 
				job_date = rs("job_date")
				job_cust_code = rs("job_cust_code")
				job_cust_name = rs("job_cust_name")
				job_cust_address = rs("job_cust_address") 
				job_cust_postcode = rs("job_cust_postcode") 
				job_cust_state = rs("job_cust_state") 
				job_cust_state_id = rs("job_cust_state_id") 
				job_cust_city = rs("job_cust_city") 
				job_cust_city_id = rs("job_cust_city_id") 
				job_cust_email = rs("job_cust_email") 
				job_cust_tel1 = rs("job_cust_tel1") 
				job_cust_tel2 = rs("job_cust_tel2")  
				job_remark = rs("job_remark")
				job_createddate = rs("job_createddate") 
				job_createdby = rs("job_createdby") 
				job_JS_receiveddate = rs("job_JS_receiveddate") 
				job_JS_receivedby = rs("job_JS_receivedby") 
				job_status = rs("job_status") 
				job_purchase_date = rs("job_purchase_date") 
				job_onlineWrtyNo = rs("job_onlineWrtyNo") 
				job_onlineWrtyStatus = rs("job_onlineWrtyStatus") 
				job_type = rs("job_type") 
				job_SN_no = rs("job_SN_no") 
				job_Model = rs("job_Model") 
				job_faulty_reason_cs = rs("job_faulty_reason_cs") 
				job_faulty_desc = rs("job_faulty_desc") 
				job_reportedby = rs("job_reportedby") 
				job_appointment_date = rs("job_appointment_date") 
				job_appointment_time = rs("job_appointment_time") 
				job_tech_code = rs("job_tech_code") 
				job_appointment_remark = rs("job_appointment_remark") 
				job_emailsentdate = rs("job_emailsentdate") 
				job_emailsent = rs("job_emailsent") 
				job_smssentdate = rs("job_smssentdate") 
				job_smssent = rs("job_smssent") 
				job_tech_type = rs("job_tech_type") 
				job_tech_model = rs("job_tech_model") 
				job_tech_tax_invoice = rs("job_tech_tax_invoice") 
				job_tech_SN = rs("job_tech_SN") 
				job_tech_faulty_reason = rs("job_tech_faulty_reason") 
				job_tech_faulty_action = rs("job_tech_faulty_action") 
				job_tech_status = rs("job_tech_status") 
				job_tech_product_collectdate = rs("job_tech_product_collectdate") 
				job_tech_returntoCustDate = rs("job_tech_returntoCustDate") 
				job_actual_wrty_status = rs("job_actual_wrty_status") 
				job_wrty_photo = rs("job_wrty_photo") 
				job_wrty_photo2 = rs("job_wrty_photo2") 
				job_wrty_photo3 = rs("job_wrty_photo3") 
				job_tech_logby  = rs("job_tech_logby")
				job_tech_logdate  = rs("job_tech_logdate")
				job_hq_remark = rs("job_hq_remark") 
				job_hq_category_code = rs("job_hq_category_code") 
				job_hq_received_date = rs("job_hq_received_date") 
				job_totalPartsAmt = rs("job_totalPartsAmt") 
				job_totallabourAmt = rs("job_totallabourAmt") 
				job_totaltransportAmt = rs("job_totaltransportAmt") 
				job_totalAmt = rs("job_totalAmt") 
				job_repair_date = rs("job_repair_date") 
				job_return_tech_date = rs("job_return_tech_date") 
				job_office_issueRemark = rs("job_office_issueRemark") 
				job_office_supervisor = rs("job_office_supervisor") 
				job_office_taxinvoice = rs("job_office_taxinvoice") 
				job_rcn_no = rs("job_rcn_no")
				job_rcn_Date = rs("job_rcn_Date")
				job_inv_no = rs("job_inv_no") 
				job_inv_date = rs("job_inv_date") 
				job_do_no = rs("job_do_no")
				job_do_date = rs("job_do_date") 
			End If
				rs.Close	
	    sql1="SELECT rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_cnty_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2,  " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_job_code = '" & request("job_code") & "'"
	    set rs1 = server.CreateObject("adodb.recordset")
		rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if rs1.eof then 
	    rs1.addnew
			rs1("rcn_status") = "Open"
			rs1("rcn_date") = ChkDateTimeMySQL(now())
			rs1("rcn_job_code") = job_code
			rs1("rcn_tech_code") = job_tech_code
			rs1("rcn_cust_code") = job_cust_code
			rs1("rcn_cust_name") = job_cust_name
			rs1("rcn_cust_address") = job_cust_address
			rs1("rcn_cust_postcode") = job_cust_postcode
			rs1("rcn_cust_state") = job_cust_state
			rs1("rcn_cust_state_id") = job_cust_state_id
			rs1("rcn_cust_city") = job_cust_city
			rs1("rcn_cust_city_id") = job_cust_city_id
			rs1("rcn_cust_cnty_id") = rcn_cust_cnty_id			
			rs1("rcn_cust_email") = job_cust_email
			rs1("rcn_cust_tel1") = job_cust_tel1
			rs1("rcn_cust_tel2") = job_cust_tel2
			rs1("rcn_createddate") = ChkDateTimeMySQL(now()) 
			rs1("rcn_createdby") = Request.Cookies("GAPS")("sloginid")
			rs1("rcn_totalqty") = 1
			rs1("rcn_totalPartsAmt") = 0
			rs1("rcn_labourAmt") = 0
			rs1("rcn_transportAmt") = 0
			rs1("rcn_gstAmt") = 0
			rs1("rcn_gstRate") = 6
			rs1("rcn_gstCode") = "SR"
			rs1("rcn_totalAmt") = 0
		rs1.update
		end if
		rs1.close
		
        sql = "select top 1 rcn_id from tblrcn order by rcn_id desc "
        rcn_id = selectid(sql)
		temp = 100000 + rcn_id
        rcn_no = "RCN" & temp 
        sql = "update tblrcn set rcn_no = '" & rcn_no & "' where rcn_id = " & rcn_id
        CUD(sql)
	    
		sql = "Update tbljob set job_rcn_no='" & rcn_no & "', job_rcn_Date='" & ChkDateTimeMySQL(now())  & "' where job_code='" & request("job_code") & "'" 
		CUD(sql)
		
		sql = "SELECT md_desc FROM tblmodel where md_code='" & job_tech_model & "'"
		rcnd_desc = selectid(sql)
		
				'''''RCN_detail
				sql="SELECT top 1 rcnd_id, rcnd_rcn_no, rcnd_job_code, rcnd_partcode, rcnd_desc, rcnd_unitcost, rcnd_qty, rcnd_discountamt, rcnd_discounttype, " & _
					"rcnd_netcost, rcnd_subtotal FROM tblrcn_detail where rcnd_job_code='" & job_code & "' "
				set rs = server.CreateObject("adodb.recordset")
				rs.ActiveConnection = strconnect
				rs.Source = sql
				rs.CursorLocation  = 3
				rs.CursorType = 2
				rs.LockType = 2
				rs.Open
				if rs.eof then 
				rs.addnew
				rs("rcnd_rcn_no") = rcn_no
				rs("rcnd_job_code") = job_code
				rs("rcnd_partcode") = job_tech_model
				rs("rcnd_desc") = rcnd_desc
				rs("rcnd_unitcost") = 0
				rs("rcnd_qty") = 1
				rs("rcnd_discountamt") = 0
				rs("rcnd_discounttype") = "%"
				rs("rcnd_netcost") = 0
				rs("rcnd_subtotal") = 0
				rs.update
				end if
				rs.close
		
		        response.redirect "rm_rcn_new.asp?rcn_no=" & rcn_no
		
		'elseif ChkString(Request.Form("job_office_issueRemark")) = "Posted" then 	
	     '      sql = "update tbljob set job_postedby='" & Request.Cookies("GAPS")("sloginid") & "', job_posteddate='" & ChkDateTimeMySQL(now()) & "', job_status='Posted' where  job_code='" & request("job_code") & "'"
			'   CUD(sql)			   
			 '  response.redirect "rm_jobsheet_view.asp?job_status=Posted&loginerr=Job HQ has been updated.#articletitle" 			
		end if 
	
         url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=Job HQ has been updated.#articletitle" 
        
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','editjob_hq=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

 '----------------------------------------------------------------------------------------------------    
  Case "duplicateJob"  

   sql2 = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id,job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_remark, job_createddate, job_createdby, job_submittedby, job_submitteddate, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo,job_wrty_photo2,job_wrty_photo3, job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs2 = server.CreateObject("adodb.recordset")
	    rs2.ActiveConnection = strconnect
		rs2.Source = sql2
		rs2.CursorLocation  = 3
		rs2.CursorType = 2
        rs2.LockType = 2
		rs2.Open
        if not rs2.eof then 
			
			   '''''' duplicate new job
				sql1 = "SELECT top 1 job_id, job_code, job_count, job_linkedcode, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id, job_cust_email, " & _
						"job_cust_tel1, job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
						"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
						"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
						"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
						"job_wrty_photo,job_wrty_photo2,job_wrty_photo3, job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
						"job_repair_date, job_return_tech_date, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no " & _
						"FROM tbljob "	
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,2,2,&H0001
				rs1.AddNew   
				rs1("job_status")  = "Open"   					
				rs1("job_date") = ChkDateYYYYMMDD(date())
				rs1("job_cust_code")  = ChkString(rs2("job_cust_code"))	
				rs1("job_count")  = ChkString(rs2("job_count")) + 1	
				rs1("job_linkedcode")  = ChkString(rs2("job_code"))	
				rs1("job_cust_name")  = ChkString(rs2("job_cust_name"))
				rs1("job_cust_address") = ChkString(rs2("job_cust_address"))
				rs1("job_cust_postcode") = ChkString(rs2("job_cust_postcode"))
	
				if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
					rs1("job_cust_state") = ChkString(rs2("job_cust_state"))
					rs1("job_cust_state_id") = ChkString(rs2("job_cust_state_id")) 
				end if
				
				rs1("job_cust_city") = ChkString(rs2("job_cust_city")) 
				rs1("job_cust_city_id") = ChkString(rs2("job_cust_city_id")) 
				rs1("job_cust_cnty_id") = ChkString(rs2("job_cust_cnty_id")) 
				rs1("job_cust_email") = ChkString(rs2("job_cust_email")) 
				rs1("job_cust_tel1") = ChkString(rs2("job_cust_tel1")) 
				rs1("job_cust_tel2") = ChkString(rs2("job_cust_tel2")) 
				rs1("job_remark") = ChkString(rs2("job_remark")) 
				rs1("job_createddate") = ChkDateTimeMySQL(now())
				rs1("job_createdby") = Request.Cookies("GAPS")("sloginid")
				rs1("job_JS_receiveddate") = ChkString(rs2("job_JS_receiveddate"))
				rs1("job_JS_receivedby") = ChkString(rs2("job_JS_receivedby"))
				rs1("job_purchase_date") = ChkString(rs2("job_purchase_date")) 
				rs1("job_onlineWrtyNo") = ChkString(rs2("job_onlineWrtyNo")) 
				rs1("job_onlineWrtyStatus") = ChkString(rs2("job_onlineWrtyStatus")) 
				rs1("job_type") = ChkString(rs2("job_type")) 
				rs1("job_SN_no") = ChkString(rs2("job_SN_no")) 
				rs1("job_Model") = ChkString(rs2("job_Model")) 
				rs1("job_model_desc") = ChkString(rs2("job_model_desc")) 
				rs1("job_faulty_reason_cs") = ChkString(rs2("job_faulty_reason_cs")) 
				rs1("job_faulty_desc") = ChkString(rs2("job_faulty_desc")) 
				rs1("job_reportedby") = ChkString(rs2("job_reportedby"))
				rs1("job_appointment_date") = ChkString(rs2("job_appointment_date")) 
				rs1("job_appointment_time") = ChkString(rs2("job_appointment_time"))
				rs1("job_tech_code") = ChkString(rs2("job_tech_code")) 
				rs1("job_appointment_remark")  = ChkString(rs2("job_appointment_remark"))	
				rs1("job_tech_product_collectdate")  = ChkString(rs2("job_tech_product_collectdate"))	
				rs1.Update 
				rs1.Close     
				
				'''
				sql3 = "select state_code from tblstate where state_name='" & rs2("job_cust_state") & "'"
				job_cust_state = selectid(sql3)
				
				sql3 = "select top 1 job_id from tbljob order by job_id desc "
				job_id = selectid(sql3) 
				
				job_id_old = rs2("job_id")
				temp = 10000 + job_id_old
				job_count = ChkString(rs2("job_count")) + 1
				job_cust_cnty_id = rs2("job_cust_cnty_id")  'determine country code

				temp = ZeroPadLeft(job_id,6)

				if job_cust_cnty_id = "129" then 'for Malaysia
					job_code = state_code & "-" & temp &  "-" & job_count & "-" & rs2("job_type")
				elseif job_cust_cnty_id = "130" then 'for Singapore
					job_code = "SG" & "-" & temp &  "-" & job_count & "-" & rs2("job_type")
				end if
			'job_code = job_cust_state & "-" & temp & "-" & job_count & "-" & rs2("job_type")
		  
				sql = "update tbljob set job_code = '" & job_code & "' where job_id = " & job_id
				CUD(sql)
		
		rs2.Close 
		end if

        url = "rm_jobsheet.asp?job_code=" & job_code & "&job_status=Submitted&loginerr=Job has been duplicated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','duplicateJob=" & ChkString(left(job_code,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		

 '----------------------------------------------------------------------------------------------------    
  Case "submitJob"  
	Dim submitJobModel, submitJobPurchaseDate, submitJobWarrantyMonth, submitJobWarrantyExpiryDate

	if request.form("job_cust_cnty_id") = "129" then
		sql = "select ct_name from tblcity where ct_id =" & request("job_cust_city_id") 
		job_cust_city = selectid(sql)
		
		sql = "select state_name from tblstate where state_id =" & request("job_cust_state_id") 
		job_cust_state = selectid(sql)
		job_cust_state_id = request("job_cust_state_id")

		sql = "select state_code from tblstate where state_id =" & request("job_cust_state_id") 
		state_code = selectid(sql)
	end if

	if  request.form("job_cust_cnty_id") <> "129" then 
			job_cust_city = request.form("job_cust_city")
	   	    job_cust_city_id = "0"
	end if
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id,job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_submittedby, job_submitteddate, job_pendingby, job_pendingdate, job_doneby, job_donedate, job_postedby, job_posteddate, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate, job_remark, job_model_desc, job_faulty_reason_cs " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
		    job_id = rs("job_id")
			rs("job_status") = "Submitted"
			rs("job_submittedby") = Request.Cookies("GAPS")("sloginid")
			rs("job_submitteddate") = ChkDateTimeMySQL(now())
			
			
			''''re-save job information.
			rs("job_cust_code")  = ChkString(Request.Form("job_cust_code"))	
			rs("job_cust_name")  = ChkString(Request.Form("job_cust_name"))
			rs("job_cust_address") = ChkString(Request.Form("job_cust_address"))
		    rs("job_cust_postcode") = ChkString(Request.Form("job_cust_postcode"))
			
			if request.form("job_cust_cnty_id") = "129" then
				rs("job_cust_state") = job_cust_state
				rs("job_cust_state_id") = job_cust_state_id
			end if
		    rs("job_cust_city") = job_cust_city
		    rs("job_cust_city_id") = ChkString(Request.Form("job_cust_city_id"))
			rs("job_cust_cnty_id") = ChkString(Request.Form("job_cust_cnty_id"))	
			rs("job_cust_email") = ChkString(Request.Form("job_cust_email")) 
			rs("job_cust_tel1") = ChkString(Request.Form("job_cust_tel1")) 
			rs("job_cust_tel2") = ChkString(Request.Form("job_cust_tel2")) 
			rs("job_remark") = ChkString(Request.Form("job_remark")) 

			submitJobModel = Trim(ChkString(Request.Form("job_Model")))
			submitJobPurchaseDate = Trim(ChkString(Request.Form("job_purchase_date")))
			rs("job_onlineWrtyStatus") = "Over"

			if submitJobPurchaseDate <> "" then
				rs("job_purchase_date") = submitJobPurchaseDate
			end if

			if submitJobModel <> "" and submitJobPurchaseDate <> "" then
				sql5 = "select [month] from tbl_warranty where md_code = '" & submitJobModel & "'"
				set rs5 = server.CreateObject("adodb.recordset")
				rs5.ActiveConnection = strconnect
				rs5.Source = sql5
				rs5.CursorLocation  = 3
				rs5.CursorType = 2
				rs5.LockType = 2
				rs5.Open
				if not rs5.eof then 					
					submitJobWarrantyMonth = rs5("month")
				end if

				'sql = "select [month] from tbl_warranty where md_code = '" & submitJobModel & "'"
				'submitJobWarrantyMonth = selectid(sql)

				if not isnull(submitJobWarrantyMonth) then
					if trim(CStr(submitJobWarrantyMonth & "")) <> "" and isnumeric(submitJobWarrantyMonth) and chkdate(submitJobPurchaseDate) <> "" then
						submitJobWarrantyExpiryDate = DateAdd("m", CLng(submitJobWarrantyMonth), CDate(chkdate(submitJobPurchaseDate)))
						if Date() <= submitJobWarrantyExpiryDate then
							rs("job_onlineWrtyStatus") = "Under"
						end if
					end if
				end if
			end if
			rs("job_onlineWrtyNo") = ChkString(Request.Form("job_onlineWrtyNo")) 
			rs("job_type") = ChkString(Request.Form("job_type")) 
			rs("job_SN_no") = ChkString(Request.Form("job_SN_no")) 
			rs("job_Model") = ChkString(Request.Form("job_Model")) 
			rs("job_model_desc") = ChkString(Request.Form("job_model_desc")) 
			rs("job_faulty_reason_cs") = ChkString(Request.Form("job_faulty_reason_cs")) 
			rs("job_faulty_desc") = ChkString(Request.Form("job_faulty_desc")) 
			rs("job_reportedby") = ChkString(Request.Form("job_reportedby"))
			if ChkString(Request.Form("job_appointment_date"))  <> "" then
			rs("job_appointment_date") = ChkString(Request.Form("job_appointment_date")) 
			end if
			rs("job_appointment_time") = ChkString(Request.Form("job_appointment_time"))
			rs("job_tech_code") = ChkString(Request.Form("job_tech_code")) 
			rs("job_appointment_remark")  = ChkString(Request.Form("job_appointment_remark"))	
			
			
			if Request.Cookies("GAPS")("slevel") = "technician2" then 
			   url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&job_status=Submitted&loginerr=Job has been updated.#articletitle" 
			else
			   url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Submitted&loginerr=Job has been updated.#articletitle" 
			end if
			
			if rs("job_tech_code") = "resolved_no_appt" then 
			        rs("job_status") = "Posted"
					rs("job_pendingby") = Request.Cookies("GAPS")("sloginid")
					rs("job_pendingdate") = ChkDateTimeMySQL(now())
					rs("job_doneby") = Request.Cookies("GAPS")("sloginid")
					rs("job_donedate") = ChkDateTimeMySQL(now())
					rs("job_postedby") = Request.Cookies("GAPS")("sloginid")
					rs("job_posteddate") = ChkDateTimeMySQL(now())
					url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Done&loginerr=Job has been Done.#articletitle" 
			end if
		rs.Update 
		rs.Close 
		end if
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','submitJob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "ClaimCheckOK" 'this option triggered by CS team
		if  Request.Form("ClaimCheck_chk") = "OK" then
			sql4 = "Update tblrpr_techcommission set rpc_checkedby='" & Request.Cookies("GAPS")("sloginid") & "', rpc_checked_date='" & ChkDateTimeMySQL(now()) & "' where rpc_id = " & Request("rpc_id") & ""
			CUD(sql4)			

			'011223 - this logic is to ensure once the claim has been submitted and appears in claim form (processed), it cannot be regenerated/considered for other months.
			sql5= "update tbljob set job_claim_approved = 'Yes' , job_claim_approved_date = '" & ChkDateTimeMySQL(now()) & "' where tbljob.job_id is not null and tbljob.job_status = 'Posted' and tbljob.job_actual_wrty_status in('Under','Over') " & _
		          "and tbljob.job_tech_code = '" & Request("tech_code") & "' and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NULL "
			CUD(sql5)

			sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrpr_techcommission','CheckClaim=" & Request("rpc_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
			CUD(sql)
	    end if

	if Request("techtype") ="IHT" then
		url = "rm_rpt_tech_monthcommisionIHT.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="IC" then
		url = "rm_rpt_tech_monthcommisionIC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="IHC" then
		url = "rm_rpt_tech_monthcommisionIHC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="TPC" then
		url = "rm_rpt_tech_monthcommisionTPC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	end if
'----------------------------------------------------------------------------------------------------    
 Case "ClaimVerifyOK" 'this option triggered by managers with Can_Verify rights

		if  Request.Form("ClaimVerify_chk") = "OK" then
			sql4 = "Update tblrpr_techcommission set rpc_verifiedby='" & Request.Cookies("GAPS")("sloginid") & "', rpc_verified_date='" & ChkDateTimeMySQL(now()) & "' where rpc_id = " & Request("rpc_id") & ""
			CUD(sql4)
		
			sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrpr_techcommission','VerifyClaim=" & Request("rpc_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
			CUD(sql)
	    end if
	if Request("techtype") ="IHT" then
		url = "rm_rpt_tech_monthcommisionIHT.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="IC" then
		url = "rm_rpt_tech_monthcommisionIC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="IHC" then
		url = "rm_rpt_tech_monthcommisionIHC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	elseif Request("techtype") ="TPC" then
		url = "rm_rpt_tech_monthcommisionTPC.asp?type=showresult&jobmonth=" & jobmonth & "&jobyear=" & jobyear & "&loginerr=Claim has been updated.#articletitle" 	
	end if
'----------------------------------------------------------------------------------------------------    

  Case "Acceptedjob"

  listjob_code = "'" & request("Accepted") & "'"
  listjob_code = replace(listjob_code, " ", "")
  listjob_code = replace(listjob_code, ",", "','")
  
  sql = "update tbljob set job_status='Accepted', job_JS_receiveddate='" & ChkDateTimeMySQL(now()) & "', job_JS_receivedby='" & Request.Cookies("GAPS")("sloginid") & "' where job_code in (" & listjob_code & ")"
  CUD(sql)
  
  if Request.Cookies("GAPS")("slevel") =  "technician" then
	url = "rmtech_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Accepted&loginerr=Job has been updated.#articletitle" 
  else
    url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Accepted&loginerr=Job has been updated.#articletitle" 
  end if	
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','submitJob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)
 
'---------------------------------------------------------------------------------------------------
  Case "JobActionClaimandAccept"

	'this logic accepts jobs to be accepted by tech and also claims submission by technicians
	
  if len(request("Accepted")) > 0 then
	listjob_code = "'" & request("Accepted") & "'"
	listjob_code = replace(listjob_code, " ", "")
	listjob_code = replace(listjob_code, ",", "','")
	
	  sql = "update tbljob set job_status='Accepted', job_JS_receiveddate='" & ChkDateTimeMySQL(now()) & "', job_JS_receivedby='" & Request.Cookies("GAPS")("sloginid") & "' where job_code in (" & listjob_code & ")"
	  CUD(sql)
  
	  if Request.Cookies("GAPS")("slevel") =  "technician" then
		url = "rmtech_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Accepted&loginerr=Job has been updated.#articletitle" 
	  else
		url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Accepted&loginerr=Job has been updated.#articletitle" 
	  end if	
	
		sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','submitJob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
		CUD(sql)
	else 	
		'tech submit claims for processing from Done jobs list 02/10/2023
		listjob_code = "'" & request("Claim") & "'"
		listjob_code = replace(listjob_code, " ", "")
		listjob_code = replace(listjob_code, ",", "','")
	    'sql = "update tbljob set job_submitforclaims='Yes', job_amount_received = '" & ChkNumber2(request("job_amount_received")) & "', job_submitforclaims_date='" & ChkDateTimeMySQL(now()) & "' where job_code in (" & listjob_code & ")"
		sql = "update tbljob set job_submitforclaims='Yes', job_submitforclaims_date='" & ChkDateTimeMySQL(now()) & "' where job_code in (" & listjob_code & ")"
	    CUD(sql)
	   
	  if Request.Cookies("GAPS")("slevel") =  "technician" then
		url = "rmtech_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Done&loginerr=Claims submitted.#articletitle" 
	  else
		url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Done&loginerr=Claims submitted.#articletitle" 
	  end if

	end if
 
  '----------------------------------------------------------------------------------------------------
	Case "addManualClaim"

	 sql = "select * from tbltech_claim_manual"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		

		if Request.Form("tech_code") <> ""  then
			rs.AddNew   
			rs("tech_code") = Request.Form("tech_code")
			rs("total_petrol") = ChkNumber2Decimal(Request.Form("total_petrol"))
			rs("total_toll") = ChkNumber2Decimal(Request.Form("total_toll"))
			rs("total_parking") = ChkNumber2Decimal(Request.Form("total_parking"))
			rs("total_incentive") = ChkNumber2Decimal(Request.Form("total_incentive"))
			rs("completed") = "No"
			rs("period") = Request.Form("period")
			rs("entry_date") = ChkDateTimeMySQL(now())
			rs.Update 
		end if
		rs.Close    				  
		   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_manual','addClaims=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

	url = "rmtech_claims_manual.asp?tech_code=" & request("tech_code") & "&loginerr=New manual claims has been updated.#claims" 
'-----------------------------------------------------------------------------------------------------------
	Case "editManualClaim"

	'this logic adds the line items for petrol/mileage claims by the tech
	    sql = "select * from tbltech_claim_manual where claims_id ='" & ChkString(Request.Form("claim_id")) &"'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
			rs("tech_code") = chkstring(Request.Form("tech_code"))
			rs("total_petrol") = ChkNumber2Decimal(Request.Form("total_petrol"))
			rs("total_toll") = ChkNumber2Decimal(Request.Form("total_toll"))
			rs("total_parking") = ChkNumber2Decimal(Request.Form("total_parking"))
			rs("total_incentive") = ChkNumber2Decimal(Request.Form("total_incentive"))
			rs("completed") = ChkString(Request.Form("completed"))
			rs("period") = ChkString(Request.Form("period"))
			'rs("entry_date") = ChkString(Request.Form("entry_date"))
			rs.Update 
			rs.Close 
				  
  	     url = "rmtech_claims_manual.asp?tech_code=" & request("claim_id") & "&loginerr=New claims has been updated.#claims" 
	     sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_manual','editManualClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"
         CUD(sql)
'---------------------------------------------------------------------------------------------------
  Case "delClaimsManual"
	sql = "delete from tbltech_claim_manual where claims_id='" & request("claim_id") &"'"
	CUD(sql)
	
    url = "rmtech_claims_manual.asp?tech_code=" & request("tech_code") & "&loginerr=Manual claims has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_manual','delClaimsManual=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'---------------------------------------------------------------------------------------------------
	Case "addPetrolClaim"
	
	    sql = "select * from tbltech_claim_petrol where tp_month ='" & ChkString(Request.Form("jobmonth")) & "' and tp_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tp_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		

		if Request.Form("trip_date") <> "" and Request.Form("mileage_start") <> "" and Request.Form("mileage_end") <> "" and Request.Form("claim_amount") <>"" and (Request.Form("mileage_end") > Request.Form("mileage_start")) then
			rs.AddNew   
			rs("tp_date") = Request.Form("submit_date")
			rs("tp_year") = Request.Form("jobyear")	
			rs("tp_month") = Request.Form("jobmonth")	
			rs("tp_tech_code") = Request.Form("tech_code")
			rs("tp_vehicle_no") = ChkString(Request.Form("vehicle_no"))
			rs("tp_trip_date") = Request.Form("trip_date")
			rs("tp_job_sheet") = Request.Form("job_sheet")
			rs("tp_mileage_start") = Request.Form("mileage_start")
			rs("tp_mileage_end") = Request.Form("mileage_end")
			rs("tp_distance") = Request.Form("distance")
			rs("tp_claim_amount") = Request.Form("claim_amount")
			'rs("tp_claim_amount") = ChkString(Request.Form("jobp_faultycode"))		
			rs.Update 
		end if
		rs.Close    				  
		   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_petrol','addPetrolClaims=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

		'adding record into main claim table

	if Request.Form("trip_date") <> "" and Request.Form("mileage_start") <> "" and Request.Form("mileage_end") <> "" and Request.Form("claim_amount") <> "" and (Request.Form("mileage_end") > Request.Form("mileage_start")) then	
	    sql = "select sum(tp_claim_amount) as total_petrol from tbltech_claim_petrol where tp_month ='" & ChkString(Request.Form("jobmonth")) & "' and tp_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
			" and tp_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    total_petrol_sum = selectid(sql) 'get the total mileage sum for the given tech, month and year

	    sql = "select * from tbltech_claim where tc_month ='" & ChkString(Request.Form("jobmonth")) & "' and tc_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tc_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if rs.eof then 
			rs.AddNew
	    end if
		'rs("tp_date") = Request.Form("submit_date")
		rs("tc_submit_date") = chkdate(date())
		rs("tc_year") = Request.Form("jobyear")	
		rs("tc_month") = Request.Form("jobmonth")	
		rs("tc_tech_code") = Request.Form("tech_code")
		rs("tc_total_petrol") = total_petrol_sum * 0.9 '11/11/24 90% of total petrol claims
		
		rs.update			
	    rs.close
	end if

	url = "rmtech_petrol_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New petrol claims has been updated.#claims" 
  '----------------------------------------------------------------------------------------------------
  Case "delPetrolClaim"
	sql = "delete from tbltech_claim_petrol where tp_tech_code='" & request("tech_code") &"' and tp_id ='" & request("claim_id") &"'"
	CUD(sql)
	
    sql = "select sum(tp_distance) as total_mileage from tbltech_claim_petrol where tp_month ='" & ChkString(Request("jobmonth")) & "' and tp_year ='" & ChkString(Request("jobyear")) & "' " & _
  		  " and tp_tech_code ='" & ChkString(Request("tech_code")) & "'"
	total_mileage_sum = selectid(sql) 'get the total mileage sum for the given tech, month and year
	sql = "update tbltech_claim set tc_total_petrol = '" & total_mileage_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
	CUD(sql)			
	
	url = "rmtech_petrol_claims.asp?tech_code=" & request("tech_code") & "&loginerr=Petrol claims has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_petrol','delPetrolClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

  '----------------------------------------------------------------------------------------------------
	Case "editPetrolClaim"

	'this logic adds the line items for petrol/mileage claims by the tech
	    sql = "select * from tbltech_claim_petrol where tp_tech_code ='" & ChkString(Request.Form("tech_code")) & "' and tp_id = '" & ChkString(Request.Form("claim_id")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if Request.Form("trip_date") <> "" and Request.Form("mileage_start") <> "" and Request.Form("mileage_end") <> ""  and Request.Form("claim_amount") <> "" and (Request.Form("mileage_end") > Request.Form("mileage_start")) then
			rs("tp_date") = Request.Form("submit_date")
			rs("tp_year") = Request.Form("jobyear")	
			rs("tp_month") = Request.Form("jobmonth")	
			rs("tp_tech_code") = Request.Form("tech_code")
			rs("tp_vehicle_no") = ChkString(Request.Form("vehicle_no"))
			rs("tp_trip_date") = Request.Form("trip_date")
			rs("tp_job_sheet") = Request.Form("job_sheet")
			rs("tp_mileage_start") = Request.Form("mileage_start")
			rs("tp_mileage_end") = Request.Form("mileage_end")
			rs("tp_distance") = Request.Form("distance")
			rs("tp_claim_amount") = Request.Form("claim_amount")
			'rs("tp_claim_amount") = ChkString(Request.Form("jobp_faultycode"))		
			rs.Update 
			rs.Close 

			sql = "select sum(tp_claim_amount) as total_petrol_sum from tbltech_claim_petrol where tp_month ='" & ChkString(Request.Form("jobmonth")) & "' and tp_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
				" and tp_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	
			total_petrol_sum = selectid(sql) 'get the total mileage sum for the given tech, month and year
			total_petrol_sum = total_petrol_sum * 0.9 '11/11/24 approve 90% of petrol claim
			sql = "update tbltech_claim set tc_total_petrol = '" & total_petrol_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
			CUD(sql)	
		end if
				  
  	    url = "rmtech_petrol_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New petrol claims has been updated.#claims" 
	   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_petro','editPetrolClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------
	Case "addParkingTollClaim"
	
	    sql = "select * from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		

		if (Request.Form("parking_amount") <> "" or Request.Form("toll_amount") <> "") and Request.Form("job_sheet") <> "" then
			rs.AddNew   
			rs("tpt_date") = Request.Form("submit_date")
			rs("tpt_trip_date") = Request.Form("trip_date")
			rs("tpt_year") = Request.Form("jobyear")	
			rs("tpt_month") = Request.Form("jobmonth")	
			rs("tpt_tech_code") = Request.Form("tech_code")
			rs("tpt_job_sheet") = Request.Form("job_sheet")	
			rs("tpt_parking_amount") = ChkNumber2Decimal(Request.Form("parking_amount"))
			rs("tpt_toll_amount") = ChkNumber2Decimal(Request.Form("toll_amount"))
			rs.Update 
		end if
		rs.Close    				  
		   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_parkingtoll','addParkingTollClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

		'adding record into main claim table

	if (Request.Form("parking_amount") <> "" or Request.Form("toll_amount") <> "") and Request.Form("job_sheet") <> "" then
	    sql = "select sum(tpt_parking_amount) as total_parking from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
			" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    total_parking_sum = selectid(sql) 'get the total parking sum for the given tech, month and year

	    sql = "select sum(tpt_toll_amount) as total_toll from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
			" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    total_toll_sum = selectid(sql) 'get the total parking sum for the given tech, month and year

	    sql = "select * from tbltech_claim where tc_month ='" & ChkString(Request.Form("jobmonth")) & "' and tc_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tc_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if rs.eof then 
			rs.AddNew
	    end if
		'rs("tp_date") = Request.Form("submit_date")
		rs("tc_submit_date") = chkdate(date())
		rs("tc_year") = Request.Form("jobyear")	
		rs("tc_month") = Request.Form("jobmonth")	
		rs("tc_tech_code") = Request.Form("tech_code")
		rs("tc_total_parking") = total_parking_sum
		rs("tc_total_toll") = total_toll_sum		
		rs.update			
	    rs.close
	end if

	url = "rmtech_parkingtoll_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New parking toll claims has been updated.#claims" 
 '----------------------------------------------------------------------------------------------------
	Case "delParkingTollClaim"
  
	sql = "delete from tbltech_claim_parkingtoll where tpt_tech_code='" & request("tech_code") &"' and tpt_id ='" & request("claim_id") &"'"
	CUD(sql)
	
	sql = "select sum(tpt_parking_amount) as total_parking from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	total_parking_sum = selectid(sql) 'get the total parking sum for the given tech, month and year

	sql = "select sum(tpt_toll_amount) as total_toll from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	total_toll_sum = selectid(sql) 'get the total toll sum for the given tech, month and year

	sql = "update tbltech_claim set tc_total_parking = '" & total_parking_sum &"',tc_total_toll = '" & total_toll_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
	CUD(sql)			
	
	url = "rmtech_parkingtoll_claims.asp?tech_code=" & request("tech_code") & "&loginerr=ParkingToll claims has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_parkingtoll','delParkingtoll=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'------------------------------------------------------------------------------------------------------------
	
	Case "editParkingTollClaim"

	'this logic adds the line items for parking/toll claims by the tech
	   ' sql = "select * from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		'" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "' and tpt_id = '" & ChkString(Request.Form("claim_id")) & "'"
	     sql = "select * from tbltech_claim_parkingtoll where tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "' and tpt_id = '" & ChkString(Request.Form("claim_id")) & "'"
	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if (Request.Form("parking_amount") <> "" or Request.Form("toll_amount") <> "") and Request.Form("job_sheet") <> "" then
			rs("tpt_date") = Request.Form("submit_date")
			rs("tpt_trip_date") = Request.Form("trip_date")
			rs("tpt_year") = Request.Form("jobyear")	
			rs("tpt_month") = Request.Form("jobmonth")	
			rs("tpt_job_sheet") = Request.Form("job_sheet")	
			rs("tpt_tech_code") = Request.Form("tech_code")
			rs("tpt_parking_amount") = Request.Form("parking_amount")
			rs("tpt_toll_amount") = Request.Form("toll_amount")	
			rs.Update 
			rs.Close 

			sql = "select sum(tpt_parking_amount) as total_parking from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
				" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
			total_parking_sum = selectid(sql) 'get the total parking sum for the given tech, month and year

			sql = "select sum(tpt_toll_amount) as total_toll from tbltech_claim_parkingtoll where tpt_month ='" & ChkString(Request.Form("jobmonth")) & "' and tpt_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
				" and tpt_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
			total_toll_sum = selectid(sql) 'get the total parking sum for the given tech, month and year

			sql = "update tbltech_claim set tc_total_parking = '" & total_parking_sum &"',tc_total_toll = '" & total_toll_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
			CUD(sql)
		end if
				  
  	    url = "rmtech_parkingtoll_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New ParkingToll claims has been updated.#claims" 
	   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_petro','editPetrolClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'-----------------------------------------------------------------------------------------------------------

	Case "addHotelClaim"
	
	    sql = "select * from tbltech_claim_hotel where th_month ='" & ChkString(Request.Form("jobmonth")) & "' and th_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and th_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		

		if Request.Form("date_from") <> "" and Request.Form("date_to") <> "" and Request.Form("daily_rate") <> "" and Request.Form("claim_amount") <> "" and CDate(Request.Form("date_to")) >= CDate(Request.Form("date_from")) and Request.Form("hotel_name") <> ""  then
			rs.AddNew   
			rs("th_date") = Request.Form("submit_date")
			rs("th_year") = Request.Form("jobyear")	
			rs("th_month") = Request.Form("jobmonth")	
			rs("th_job_sheet") = Request.Form("job_sheet")
			rs("th_tech_code") = Request.Form("tech_code")
			rs("th_hotel_name") =  Request.form("hotel_name")
			rs("th_date_from") = Request.Form("date_from")
			rs("th_date_to") = Request.Form("date_to")
			rs("th_daily_rate") = Request.Form("daily_rate")
			rs("th_claim_amount") = Request.Form("claim_amount")
			rs.Update 
		end if
		rs.Close    				  
		   
        sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	       Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_hotel','addHotelDetail=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
        CUD(sql)

		'adding record into main claim table

	if Request.Form("date_from") <> "" and Request.Form("date_to") <> "" and Request.Form("daily_rate") <> "" and Request.Form("claim_amount") <> "" and (Request.Form("date_to") >= Request.Form("date_start")) then
	    sql = "select sum(th_claim_amount) as total_hotele from tbltech_claim_hotel where th_month ='" & ChkString(Request.Form("jobmonth")) & "' and th_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and th_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    total_hotel_sum = selectid(sql) 'get the total hotel sum for the given tech, month and year

	    sql = "select * from tbltech_claim where tc_month ='" & ChkString(Request.Form("jobmonth")) & "' and tc_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tc_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if rs.eof then 
			rs.AddNew
	    end if
		'rs("tp_date") = Request.Form("submit_date")
		rs("tc_submit_date") = chkdate(date())
		rs("tc_year") = Request.Form("jobyear")	
		rs("tc_month") = Request.Form("jobmonth")	
		rs("tc_tech_code") = Request.Form("tech_code")
		rs("tc_total_hotel") = total_hotel_sum
		rs.update			
	    rs.close
	end if

	url = "rmtech_hotel_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New hotel claims has been updated.#claims" 	 
  '--------------------------------------------------------------------------------------------------------------------------------------------------------------
	Case "editHotelClaim"

	   sql = "select * from tbltech_claim_hotel where th_month ='" & ChkString(Request.Form("jobmonth")) & "' and th_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and th_tech_code ='" & ChkString(Request.Form("tech_code")) & "' and th_id = '" & ChkString(Request.Form("claim_id")) & "'"

	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if Request.Form("date_from") <> "" and Request.Form("date_to") <> "" and Request.Form("daily_rate") <> "" and Request.Form("claim_amount") <> "" and CDate(Request.Form("date_to")) >= CDate(Request.Form("date_from")) and trim(Request.Form("hotel_name")) <> ""then
			rs("th_date") = Request.Form("submit_date")
			rs("th_year") = Request.Form("jobyear")	
			rs("th_month") = Request.Form("jobmonth")	
			rs("th_job_sheet") = Request.Form("job_sheet")
			rs("th_tech_code") = Request.Form("tech_code")
			rs("th_hotel_name") =  Request.form("hotel_name")
			rs("th_date_from") = Request.Form("date_from")
			rs("th_date_to") = Request.Form("date_to")
			rs("th_daily_rate") = Request.Form("daily_rate")
			rs("th_claim_amount") = Request.Form("claim_amount")
			rs.Update 
			rs.Close 

			sql = "select sum(th_claim_amount) as total_hotel from tbltech_claim_hotel where th_month ='" & ChkString(Request.Form("jobmonth")) & "' and th_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
				  " and th_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
			total_hotel_sum = selectid(sql) 'get the total hotel sum for the given tech, month and year

			sql = "update tbltech_claim set tc_total_hotel = '" & total_hotel_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
			CUD(sql)	
		end if
				  
  	    url = "rmtech_hotel_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New hotel claims has been updated.#claims" 
	   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_hotel','editHotelDetail=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 '-----------------------------------------------------------------------------------------------------------------------------------
	
	Case "delHotelClaim"
  
	sql = "delete from tbltech_claim_hotel where th_tech_code='" & request("tech_code") &"' and th_id ='" & request("claim_id") &"'"
	CUD(sql)
	
	sql = "select sum(th_claim_amount) as total_hotel from tbltech_claim_hotel where th_month ='" & ChkString(Request("jobmonth")) & "' and th_year ='" & ChkString(Request("jobyear")) & "' " & _
	" and th_tech_code ='" & ChkString(Request("tech_code")) & "'"
	
	total_hotel_sum = selectid(sql) 'get the total hotel sum for the given tech, month and year
	sql = "update tbltech_claim set tc_total_hotel = '" & total_hotel_sum &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
	CUD(sql)			
	url = "rmtech_hotel_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New hotel claims has been updated.#claims" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_hotel','delHotelDetail=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

  '----------------------------------------------------------------------------------------------------
	Case "addexMileageClaim"
	
	    sql = "select * from tbltech_claim_exmileage where te_month ='" & ChkString(Request.Form("jobmonth")) & "' and te_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and te_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		

		if Request.Form("trip_date") <> "" and Request.Form("loc_start") <> "" and Request.Form("loc_end") <> "" and Request.Form("job_sheet") <> "" and (Request.Form("mileage")) <> "" then
			rs.AddNew   
			rs("te_date") = Request.Form("submit_date")
			rs("te_year") = Request.Form("jobyear")	
			rs("te_month") = Request.Form("jobmonth")	
			rs("te_tech_code") = Request.Form("tech_code")
			rs("te_job_sheet") = Request.Form("job_sheet")
			rs("te_trip_date") = Request.Form("trip_date")
			rs("te_loc_start") = Request.Form("loc_start")
			rs("te_loc_end") = Request.Form("loc_end")
			rs("te_mileage") = ChkNumberInt(Request.Form("mileage"))
			rs("te_offset_mileage") = ChkNumberInt(Request.Form("offset_mileage"))
			rs("te_net_mileage") = ChkNumberInt(Request.Form("net_mileage"))
			rs("te_claim_amount") = Request.Form("claim_amount")
			rs.Update 
		end if
		rs.Close    				  
		   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_exmileage','addexMileageClaims=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"
         CUD(sql)

		'adding record into main claim table

	if Request.Form("trip_date") <> "" and Request.Form("loc_start") <> "" and Request.Form("loc_end") <> "" and (Request.Form("mileage")) <> "" then
	    sql = "select sum(te_claim_amount) as total_claim from tbltech_claim_exmileage where te_month ='" & ChkString(Request.Form("jobmonth")) & "' and te_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
			" and te_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    total_claim = selectid(sql) 'get the total mileage sum for the given tech, month and year

	    sql = "select * from tbltech_claim where tc_month ='" & ChkString(Request.Form("jobmonth")) & "' and tc_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
		" and tc_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if rs.eof then 
			rs.AddNew
	    end if
		'rs("tp_date") = Request.Form("submit_date")
		rs("tc_submit_date") = chkdate(date())
		rs("tc_year") = Request.Form("jobyear")	
		rs("tc_month") = Request.Form("jobmonth")	
		rs("tc_tech_code") = Request.Form("tech_code")
		rs("tc_total_extramileage") = total_claim
		rs.update			
	    rs.close
	end if

	url = "rmtech_exmileage_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New extra mileage claims has been updated.#claims" 
  '----------------------------------------------------------------------------------------------------
	Case "editexMileageClaim"

	'this logic adds the line items for petrol/mileage claims by the tech
	    sql = "select * from tbltech_claim_exmileage where te_tech_code ='" & ChkString(Request.Form("tech_code")) & "' and te_id = '" & ChkString(Request.Form("claim_id")) & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open		
		if Request.Form("trip_date") <> "" and Request.Form("loc_start") <> "" and Request.Form("loc_end") <> "" and Request.Form("job_sheet") <> "" and (Request.Form("mileage")) <> "" then
			rs("te_date") = Request.Form("submit_date")
			rs("te_year") = Request.Form("jobyear")	
			rs("te_month") = Request.Form("jobmonth")	
			rs("te_tech_code") = Request.Form("tech_code")
			rs("te_job_sheet") = Request.Form("job_sheet")
			rs("te_trip_date") = Request.Form("trip_date")
			rs("te_loc_start") = Request.Form("loc_start")
			rs("te_loc_end") = Request.Form("loc_end")
			rs("te_mileage") = ChkNumberInt(Request.Form("mileage"))
			rs("te_offset_mileage") = ChkNumberInt(Request.Form("offset_mileage"))
			rs("te_net_mileage") = ChkNumberInt(Request.Form("net_mileage"))
			rs("te_claim_amount") = Request.Form("claim_amount")
			rs.Update 
			rs.Close 

			sql = "select sum(te_claim_amount) as total_claim from tbltech_claim_exmileage where te_month ='" & ChkString(Request.Form("jobmonth")) & "' and te_year ='" & ChkString(Request.Form("jobyear")) & "' " & _
			" and te_tech_code ='" & ChkString(Request.Form("tech_code")) & "'"

			total_claim = selectid(sql) 'get the total mileage sum for the given tech, month and year

			sql = "update tbltech_claim set tc_total_extramileage = '" & total_claim &"' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
			CUD(sql)	
		end if
				  
  	    url = "rmtech_exmileage_claims.asp?tech_code=" & request("tech_code") & "&loginerr=New petrol claims has been updated.#claims" 
	   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_exmileage','editextramileageClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'--------------------------------------------------------------------------------------------------------------------------------------------
	Case "delexMileageClaim"
  
	sql = "delete from tbltech_claim_exmileage where te_tech_code='" & request("tech_code") &"' and te_id ='" & request("claim_id") &"'"
	CUD(sql)
	
   sql = "select sum(te_claim_amount) as total_claim from tbltech_claim_exmileage where te_month ='" & ChkString(Request("jobmonth")) & "' and te_year ='" & ChkString(Request("jobyear")) & "' " & _
			" and te_tech_code ='" & ChkString(Request("tech_code")) & "'"
	total_claim = selectid(sql) 'get the total mileage sum for the given tech, month and year
	sql = "update tbltech_claim set tc_total_extramileage = '" & total_claim & "' where tc_tech_code='" & request("tech_code") &"' and tc_month ='" & ChkString(Request("jobmonth")) & "' and tc_year ='" & ChkString(Request("jobyear")) & "'"
	CUD(sql)			
	
	url = "rmtech_exmileage_claims.asp?tech_code=" & request("tech_code") & "&loginerr=Petrol claims has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim_exmileage','delextramileageClaim=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'--------------------------------------------------------------------------------------------------------------------------------------------

  Case "PendingJob"  
  
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_submittedby, job_submitteddate, job_pendingby, job_pendingdate, job_doneby, job_donedate, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("job_status") = "Pending"
			rs("job_pendingby") = Request.Cookies("GAPS")("sloginid")
			rs("job_pendingdate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Pending&loginerr=Job has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','doneJob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql) 
 
  '----------------------------------------------------------------------------------------------------    
  Case "resetJobOpen"  
  
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_submittedby, job_submitteddate, job_pendingby, job_pendingdate, job_doneby, job_donedate, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("job_status") = "Open"
		rs.Update 
		rs.Close 
		end if

        url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=Job status has been updated to Open .#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','resetJobOpen=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql) 
		 
 '----------------------------------------------------------------------------------------------------    
  Case "DoneJob"  
 
 	set mySmartUpload = server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.Upload 
	
	tempid = mySmartUpload.form("job_wrty_photo")
	tempid2 = mySmartUpload.form("job_wrty_photo2")
	tempid3 = mySmartUpload.form("job_wrty_photo3")

	intCount = 1
	For each file In mySmartUpload.Files
	 IF not file.IsMissing Then   
		file.SaveAs(Server.MapPath(documentpath & tempid & file.FileName))  		
			job_wrty_photo = tempid & file.FileName
	 End if   
	intCount = intCount + 1     
	Next 
	
	intCount = 1
	For each file In mySmartUpload.Files
	 IF not file.IsMissing Then   
		file.SaveAs(Server.MapPath(documentpath & tempid2 & file.FileName))  		
			job_wrty_photo2 = tempid2 & file.FileName
	 End if   
	intCount = intCount + 1     
	Next 

	intCount = 1
	For each file In mySmartUpload.Files
	 IF not file.IsMissing Then   
		file.SaveAs(Server.MapPath(documentpath & tempid3 & file.FileName))  		
			job_wrty_photo3 = tempid3 & file.FileName
	 End if   
	intCount = intCount + 1     
	Next 


	
	sql = "select fr_description from tblfaultyreason where fr_code='" & ChkString(mySmartUpload.Form("job_tech_faulty_code")) & "' and fr_status='Y'" 
		job_tech_faulty_reason = selectid(sql)
		
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_code, job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_tech_logby, job_tech_logdate, job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_doneby, job_donedate " & _
				"FROM tbljob where job_code = '" & mySmartUpload.Form("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
		
		    if ChkString(mySmartUpload.Form("job_appointment_date")) <> "" then  
			rs("job_appointment_date") = ChkString(mySmartUpload.Form("job_appointment_date"))
			end if
			
			if ChkString(mySmartUpload.Form("job_appointment_time")) <> "" then  
			rs("job_appointment_time") = ChkString(mySmartUpload.Form("job_appointment_time"))
			end if
			
			rs("job_tech_type") = ChkString(mySmartUpload.Form("job_tech_type"))
			rs("job_tech_model")  = ChkString(mySmartUpload.Form("job_tech_model"))	
			rs("job_tech_model_desc")  = ChkString(mySmartUpload.Form("job_tech_model_desc"))	
			rs("job_tech_tax_invoice")  = ChkString(mySmartUpload.Form("job_tech_tax_invoice"))
			rs("job_tech_SN") = ChkString(mySmartUpload.Form("job_tech_SN"))
			rs("job_tech_faulty_code") = ChkString(mySmartUpload.Form("job_tech_faulty_code"))
			rs("job_tech_faulty_reason") = ChkString(job_tech_faulty_reason)
			rs("job_tech_faulty_action") = ChkString(mySmartUpload.Form("job_tech_faulty_action")) 
			rs("job_tech_status") = ChkString(mySmartUpload.Form("job_tech_status")) 
			
			if ChkString(mySmartUpload.Form("job_tech_product_collectdate")) <> "" then 
			rs("job_tech_product_collectdate") = ChkString(mySmartUpload.Form("job_tech_product_collectdate")) 
			end if
			
			if ChkString(mySmartUpload.Form("job_tech_service_date")) <> "" then 
			rs("job_tech_service_date") = ChkString(mySmartUpload.Form("job_tech_service_date")) 
			end if
			
			if ChkString(mySmartUpload.Form("job_tech_returntoCustDate")) <> "" then 
			rs("job_tech_returntoCustDate") = ChkString(mySmartUpload.Form("job_tech_returntoCustDate")) 
			end if
			
			rs("job_actual_wrty_status") = ChkString(mySmartUpload.Form("job_actual_wrty_status"))
			
			if job_wrty_photo <> "" then  
			rs("job_wrty_photo") = ChkString(job_wrty_photo)
			end if

			if job_wrty_photo2 <> "" then  
			rs("job_wrty_photo2") = ChkString(job_wrty_photo2)
			end if

			if job_wrty_photo3 <> "" then  
			rs("job_wrty_photo3") = ChkString(job_wrty_photo3)
			end if
			
			rs("job_tech_logby") = Request.Cookies("GAPS")("sloginid")
			rs("job_return_tech_date")  = ChkDateTimeMySQL(now())	
			
			rs("job_status") = "Done"
			
			if isnull(rs("job_doneby")) then 
			rs("job_doneby") = Request.Cookies("GAPS")("sloginid")
			rs("job_donedate") = ChkDateTimeMySQL(now())
			end if
			rs("job_tech_status") = "Done"
			
		rs.Update 
		rs.Close 
		end if  
		
        if request.Cookies("GAPS")("slevel") = "technician" then 
	       url = "rmtech_jobsheet_view.asp?job_code=" & mySmartUpload.Form("job_code") & "&job_status=Done&loginerr=Job has been updated.#articletitle" 
        elseif request.Cookies("GAPS")("slevel") = "technician2" then 
	       url = "tech2_jobsheet_view.asp?job_code=" & mySmartUpload.Form("job_code") & "&job_status=Done&loginerr=Job has been updated.#articletitle" 	    
		else
           url = "rm_jobsheet_view.asp?job_code=" & mySmartUpload.Form("job_code") & "&job_status=Done&loginerr=Job has been updated.#articletitle" 
		end if   
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','doneJob=" & ChkString(left(mySmartUpload.Form("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		  		 		 		 
 '----------------------------------------------------------------------------------------------------    
  Case "cancelJob"  
  
        ''''Edit Job Order	   	  
        sql = "SELECT top 1 job_id, job_code, job_count, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_email, " & _
				"job_cust_tel1, job_cust_tel2, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, " & _
				"job_onlineWrtyStatus, job_type, job_SN_no, job_Model, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, " & _ 
				"job_appointment_remark, job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, " & _
				"job_tech_faulty_reason, job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, " & _
				"job_wrty_photo,job_wrty_photo2,job_wrty_photo3, job_hq_remark, job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, " & _
				"job_repair_date, job_return_tech_date, job_logbyhq, job_logdatehq, job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_do_no, job_cancelledby, job_cancelleddate " & _
				"FROM tbljob where job_code = '" & request("job_code") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("job_status") = "Cancel"
			rs("job_cancelledby") = Request.Cookies("GAPS")("sloginid")
			rs("job_cancelleddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_jobsheet_view.asp?job_code=" & request("job_code") & "&job_status=Cancel&loginerr=Job has been cancel.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','cancelJob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)


'----------------------------------------------------------------------------------------------------    
  Case "addJobDetail"   

	if isNotExceed(ChkString(Request.Form("jobp_partcode")), ChkString(Request.Form("jobp_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
	    sql = "select job_actual_wrty_status from tbljob where job_code='" & ChkString(Request.Form("job_code")) & "'"
		job_actual_wrty_status = selectid(sql)
	
        ''''Add Job parts	   	  
        sql = "SELECT top 1 jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal, jobp_faultycode " & _
	          "FROM tbljob_parts "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        rs.AddNew   
        rs("job_code") = ChkString(Request.Form("job_code"))
        rs("jobp_partcode")  = ChkString(Request.Form("jobp_partcode"))	
		rs("jobp_desc")  = ChkString(Request.Form("jobp_desc"))	
        rs("jobp_unitcost")  = ChkString(Request.Form("jobp_unitcost"))
		rs("jobp_discountamt")  = ChkString(Request.Form("jobp_discountamt"))
		rs("jobp_discounttype")  = ChkString(Request.Form("jobp_discounttype"))
		rs("jobp_qty") = ChkString(Request.Form("jobp_qty"))
		rs("jobp_faultycode") = ChkString(Request.Form("jobp_faultycode"))
		
		if ChkString(Request.Form("jobp_discounttype")) = "%" then 
			jobp_netcost  = ChkString(Request.Form("jobp_unitcost")) * (ChkString(Request.Form("jobp_discountamt")/100))
			jobp_netcost = ChkString(Request.Form("jobp_unitcost")) - jobp_netcost
			rs("jobp_netcost")  = jobp_netcost
			rs("jobp_subtotal")  = jobp_netcost * chknumber(Request("jobp_qty"))
		else
			jobp_netcost  = ChkString(Request.Form("jobp_unitcost")) -  ChkString(Request.Form("jobp_discountamt"))
			rs("jobp_netcost")  = jobp_netcost
			rs("jobp_subtotal")  = jobp_netcost * chknumber(Request("jobp_qty"))
		end if

		if job_actual_wrty_status="Under" then
		   rs("jobp_unitcost")  = ChkString(Request.Form("jobp_unitcost"))
		   rs("jobp_discountamt")  = ChkString(Request.Form("jobp_unitcost"))
		   rs("jobp_netcost")  = 0
		   rs("jobp_subtotal") = 0
		end if
	
		rs.Update 
		rs.Close      

		sql = "SELECT top 1 jobp_faultycode " & _
	          "FROM tbljob_parts where job_code = '" & ChkString(Request.Form("job_code")) & "' and  (jobp_faultycode <> '' or jobp_faultycode is not null) order by jobp_id " 	
		jobp_faultycode = selectid(sql)  
					  
        sql = "select sum(jobp_subtotal) as job_totalPartsAmt from tbljob_parts where job_code = '" & request("job_code") & "'"
        job_totalPartsAmt = selectid(sql)
		if isnull(job_totalPartsAmt) or job_totalPartsAmt="" then 
		   job_totalPartsAmt = 0
		end if
		
        sql = "update tbljob set job_totalPartsAmt = " & (job_totalPartsAmt) & ", job_tech_faulty_code='" & jobp_faultycode & "' where job_code = '" & request("job_code") & "'"
        CUD(sql)
		
		sql = "update tbljob set job_totalAmt = job_totalPartsAmt+job_totallabourAmt+job_totaltransportAmt where job_code = '" & request("job_code") & "'"
        CUD(sql)
      
		 if request.Cookies("GAPS")("slevel") = "technician2" then 
		  url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
		 else
		  url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
		 end if
	   
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_parts','addJobDetail=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	else
		 if request.Cookies("GAPS")("slevel") = "technician2" then 
			  url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=New Parts has not been updated.#spareparts" 
		 else
			  url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Parts has not been updated.#spareparts" 
		 end if
	end if
'----------------------------------------------------------------------------------------------------    
  Case "editJobDetail"   
	if isNotExceed(ChkString(Request.Form("jobp_partcode")), ChkString(Request.Form("jobp_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
	    sql = "delete from tbljob_parts where jobp_id=" & request("jobp_id")	
	    CUD(sql)
	
        sql = "select job_actual_wrty_status from tbljob where job_code='" & ChkString(Request.Form("job_code")) & "'"
		job_actual_wrty_status = selectid(sql)
	
        ''''Add Job parts	   	  
        sql = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal, jobp_faultycode " & _
	          "FROM tbljob_parts " 	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
		rs.AddNew 
        rs("job_code") = ChkString(Request.Form("job_code"))
        rs("jobp_partcode")  = ChkString(Request.Form("jobp_partcode"))	
		rs("jobp_desc")  = ChkString(Request.Form("jobp_desc"))	
        rs("jobp_unitcost")  = ChkString(Request.Form("jobp_unitcost"))
		rs("jobp_discountamt")  = ChkString(Request.Form("jobp_discountamt"))
		rs("jobp_discounttype")  = ChkString(Request.Form("jobp_discounttype"))
		rs("jobp_qty") = ChkString(Request.Form("jobp_qty"))
		rs("jobp_faultycode") = ChkString(Request.Form("jobp_faultycode"))
		
		if ChkString(Request.Form("jobp_discounttype")) = "%" then 
			jobp_netcost  = ChkString(Request.Form("jobp_unitcost")) * (ChkString(Request.Form("jobp_discountamt")/100))
			jobp_netcost = ChkString(Request.Form("jobp_unitcost")) - jobp_netcost
			rs("jobp_netcost")  = jobp_netcost
			rs("jobp_subtotal")  = jobp_netcost * ChkString(Request.Form("jobp_qty"))
		else
			jobp_netcost  = ChkString(Request.Form("jobp_unitcost")) -  ChkString(Request.Form("jobp_discountamt"))
			rs("jobp_netcost")  = jobp_netcost
			rs("jobp_subtotal")  = jobp_netcost * ChkString(Request.Form("jobp_qty"))
		end if
		
		if job_actual_wrty_status="Under" then
		   rs("jobp_unitcost")  = ChkString(Request.Form("jobp_unitcost"))
		   rs("jobp_discountamt")  = ChkString(Request.Form("jobp_unitcost"))
		   rs("jobp_netcost")  = 0
		   rs("jobp_subtotal") = 0
		end if
		
		rs.Update 
		rs.Close      
		
		sql = "SELECT top 1 jobp_faultycode " & _
	          "FROM tbljob_parts where job_code = '" & ChkString(Request.Form("job_code")) & "' and  (jobp_faultycode <> '' or jobp_faultycode is not null) order by jobp_id " 	
		jobp_faultycode = selectid(sql)  
		
        sql = "select sum(jobp_subtotal) as job_totalPartsAmt from tbljob_parts where job_code = '" & request("job_code") & "'"
        job_totalPartsAmt = selectid(sql)
		if isnull(job_totalPartsAmt) or job_totalPartsAmt="" then 
		   job_totalPartsAmt = 0
		end if
		
        sql = "update tbljob set job_totalPartsAmt = " & (job_totalPartsAmt) & ", job_tech_faulty_code='" & jobp_faultycode & "' where job_code = '" & request("job_code") & "'"
        CUD(sql)
		
		sql = "update tbljob set job_totalAmt = job_totalPartsAmt+job_totallabourAmt+job_totaltransportAmt where job_code = '" & request("job_code") & "'"
        CUD(sql)

		if request.Cookies("GAPS")("slevel") = "technician2" then 		
		url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
		else
		url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
		end if
		   
        sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_parts','addJobDetail=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
        CUD(sql)
	else
		 if request.Cookies("GAPS")("slevel") = "technician2" then 
			  url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=New Parts has not been updated.#spareparts" 
		 else
			  url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Parts has not been updated.#spareparts" 
		 end if
	end if

'----------------------------------------------------------------------------------------------------    
  Case "delJobDetail"
  
	sql = "delete from tbljob_parts where jobp_id=" & request("jobp_id")	
	CUD(sql)
	
	sql = "select sum(jobp_subtotal) as job_totalPartsAmt from tbljob_parts where job_code = '" & request("job_code") & "'"
	job_totalPartsAmt = selectid(sql)
	if job_totalPartsAmt = "" or isnull(job_totalPartsAmt) then 
	   job_totalPartsAmt = "0"
	end if
	
	sql = "update tbljob set job_totalPartsAmt = " & (job_totalPartsAmt) & " where job_code = '" & request("job_code") & "'"
	CUD(sql)
	
	sql = "update tbljob set job_totalAmt = job_totalPartsAmt+job_totallabourAmt+job_totaltransportAmt where job_code = '" & request("job_code") & "'"
	CUD(sql)
	
	url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_parts','delJobDetail=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "editJobDetailTotal"   
		'26/06/2025 overwty allowance allows tech to be paid even there's no labour charged for spec cases.

		sql = "Update tbljob set job_overwty_allowance = '" & request("job_overwty_allowance") & "', job_totallabourAmt=" & chknumber2(request("job_totallabourAmt")) & ", " & _
		      "job_totaltransportAmt=" & chknumber2(request("job_totaltransportAmt")) & " where job_code = '" & request("job_code") & "'"
		CUD(sql) 

		sql = "select sum(jobp_subtotal) as jobp_subtotal from tbljob_parts where job_code = '" & request("job_code") & "'"
		jobp_subtotal = selectid(sql)
		if jobp_subtotal = "" or isnull(jobp_subtotal) then 
		   jobp_subtotal = "0"
		end if
		
		sql = "update tbljob set job_totalPartsAmt = " & (jobp_subtotal) & " where job_code = '" & request("job_code") & "'"
		CUD(sql)
		
		sql = "update tbljob set job_totalAmt = job_totalPartsAmt+job_totallabourAmt+job_totaltransportAmt where job_code = '" & request("job_code") & "'"
		CUD(sql)
		
		url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=New Job has been updated.#spareparts" 
		
		sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','editJobDetailTotal=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
		CUD(sql)
		 					  		   		 	  	 
'----------------------------------------------------------------------------------------------------    
  Case "generateinvoicejob"  
  	
generateinvoice="N" 
set rs = server.CreateObject("adodb.recordset")
  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_state_id, job_cust_city, job_cust_city_id, job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo,job_wrty_photo2,job_wrty_photo3, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date " & _
	    "FROM tbljob WHERE job_code = '" & request("job_code") & "' and (job_inv_no = '' or job_inv_no is null)"
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			job_id = rs("job_id") 
			job_code = rs("job_code") 
			job_date = rs("job_date")
			job_cust_code = rs("job_cust_code")
			job_cust_name = rs("job_cust_name")
			job_cust_address = rs("job_cust_address") 
			job_cust_postcode = rs("job_cust_postcode") 
			job_cust_state = rs("job_cust_state") 
			job_cust_state_id = rs("job_cust_state_id") 
			job_cust_city = rs("job_cust_city") 
			job_cust_city_id = rs("job_cust_city_id") 
			job_cust_cnty_id = rs("job_cust_cnty_id") 
			job_cust_email = rs("job_cust_email") 
			job_cust_tel1 = rs("job_cust_tel1") 
			job_cust_tel2 = rs("job_cust_tel2")  
			job_remark = rs("job_remark")
			job_createddate = rs("job_createddate") 
			job_createdby = rs("job_createdby") 
			job_JS_receiveddate = rs("job_JS_receiveddate") 
			job_JS_receivedby = rs("job_JS_receivedby") 
			job_status = rs("job_status") 
			
			if isdate(rs("job_purchase_date")) then 
			job_purchase_date = rs("job_purchase_date") 
			end if
			
			job_onlineWrtyNo = rs("job_onlineWrtyNo") 
			job_onlineWrtyStatus = rs("job_onlineWrtyStatus") 
			job_type = rs("job_type") 
			job_SN_no = rs("job_SN_no") 
			job_Model = rs("job_Model") 
			job_faulty_reason_cs = rs("job_faulty_reason_cs") 
			job_faulty_desc = rs("job_faulty_desc") 
			job_reportedby = rs("job_reportedby") 
			
			if isdate(rs("job_appointment_date")) then 
			job_appointment_date = rs("job_appointment_date") 
			end if
			
			job_appointment_time = rs("job_appointment_time") 
			job_tech_code = rs("job_tech_code") 
			job_appointment_remark = rs("job_appointment_remark") 
			if isdate(rs("job_emailsentdate")) then 
			job_emailsentdate = rs("job_emailsentdate")
			end if 
			job_emailsent = rs("job_emailsent") 
			
			if isdate(rs("job_smssentdate")) then 
			job_smssentdate = rs("job_smssentdate") 
			end if
			
			job_smssent = rs("job_smssent") 
			job_tech_type = rs("job_tech_type") 
			job_tech_model = rs("job_tech_model") 
			job_tech_tax_invoice = rs("job_tech_tax_invoice") 
			job_tech_SN = rs("job_tech_SN") 
			job_tech_faulty_reason = rs("job_tech_faulty_reason") 
			job_tech_faulty_action = rs("job_tech_faulty_action") 
			job_tech_status = rs("job_tech_status") 
			if isdate(rs("job_tech_product_collectdate")) then 
			job_tech_product_collectdate = rs("job_tech_product_collectdate") 
			end if
			
			if isdate(rs("job_tech_returntoCustDate")) then
			job_tech_returntoCustDate = rs("job_tech_returntoCustDate") 
			end if
			
			job_actual_wrty_status = rs("job_actual_wrty_status") 
			job_wrty_photo = rs("job_wrty_photo") 
			job_wrty_photo2 = rs("job_wrty_photo2") 
			job_wrty_photo3 = rs("job_wrty_photo3") 

			job_tech_logby  = rs("job_tech_logby")
			
			if isdate(rs("job_tech_logdate")) then 
			job_tech_logdate  = rs("job_tech_logdate")
			end if
			
			job_hq_remark = rs("job_hq_remark") 
			job_hq_category_code = rs("job_hq_category_code") 
			
			if isdate(rs("job_hq_received_date")) then 
			job_hq_received_date = rs("job_hq_received_date") 
			end if
			
			job_totalPartsAmt = rs("job_totalPartsAmt") 
			job_totallabourAmt = rs("job_totallabourAmt") 
			job_totaltransportAmt = rs("job_totaltransportAmt") 
			job_totalAmt = rs("job_totalAmt") 
			if isdate(rs("job_repair_date")) then 
			job_repair_date = rs("job_repair_date") 
			end if
			
			if isdate(rs("job_return_tech_date")) then 
			job_return_tech_date = rs("job_return_tech_date") 
			end if
			job_office_issueRemark = rs("job_office_issueRemark") 
			job_office_supervisor = rs("job_office_supervisor") 
			job_office_taxinvoice = rs("job_office_taxinvoice") 
			job_rcn_no = rs("job_rcn_no")
			
			if isdate(rs("job_rcn_Date")) then 
			job_rcn_Date = rs("job_rcn_Date")
			end if
			job_inv_no = rs("job_inv_no") 
			if isdate(rs("job_inv_date")) then 
			job_inv_date = rs("job_inv_date") 
			end if
			job_do_no = rs("job_do_no")
			if isdate(rs("job_do_date")) then 
			job_do_date = rs("job_do_date") 
			end if
			generateinvoice="Y"
		Else
			generateinvoice="N" 
		End If
		rs.Close
		
	  
	   if generateinvoice="Y" then 
	    sql="SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, " & _
			  "inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_job_code, inv_tech_code, inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt,  " & _
			  "inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent, inv_emailsentdate, inv_status " & _
			  "FROM tblinvoice where inv_job_code = '" & request("job_code") & "'"
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then 
	    rs.addnew
			rs("inv_status") = "Open"
			rs("inv_date") = ChkDateYYYYMMDD(date())
			rs("inv_cust_code") = job_cust_code
			rs("inv_cust_name") = job_cust_name
			rs("inv_cust_address") = job_cust_address
			rs("inv_cust_postcode") = job_cust_postcode
			rs("inv_cust_state") = job_cust_state
			rs("inv_cust_state_id") = job_cust_state_id
			rs("inv_cust_city") = job_cust_city
			rs("inv_cust_city_id") = job_cust_city_id
			rs("inv_cust_cnty_id") = job_cust_cnty_id
			rs("inv_cust_email") = job_cust_email
			rs("inv_cust_tel1") = job_cust_tel1
			rs("inv_cust_tel2") = job_cust_tel2
			rs("inv_createddate") = ChkDateTimeMySQL(now()) 
			rs("inv_createdby") = Request.Cookies("GAPS")("sloginid")
			rs("inv_job_code") = job_code
			rs("inv_tech_code") = job_tech_code
			rs("inv_totalqty") = job_totalqty
			rs("inv_totalPartsAmt") = job_totalPartsAmt
			rs("inv_labourAmt") = job_totallabourAmt
			rs("inv_transportAmt") = job_totaltransportAmt
			rs("inv_gstAmt") = job_totalAmt * GSTRateBack
			rs("inv_gstRate") = GSTRate
			rs("inv_gstCode") = GSTCode
			rs("inv_totalAmt") = job_totalAmt + (job_totalAmt * GSTRateBack)
			rs("inv_no") = "temp-" & ChkDateTimeMySQL(now()) ' 280724 temp number as inv_no cannot be null
		rs.update
		end if
		rs.close
		
        sql = "select top 1 inv_id from tblinvoice order by inv_id desc"
        inv_id = selectid(sql)
		'temp = 1 + inv_id
        'inv_no = "RS" & temp 
		inv_no = "RS" & InvoiceNumbering(inv_id)
        sql = "update tblinvoice set inv_no = '" & inv_no & "' where inv_id = " & inv_id
        CUD(sql)
	    
		sql = "Update tbljob set job_inv_no='" & inv_no & "', job_inv_date='" & ChkDateTimeMySQL(now())  & "' where job_code='" & request("job_code") & "'" 
		CUD(sql)
		
		'''''job_detail
		sql1 = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal " & _
		       "FROM tbljob_parts where job_code = '" & request("job_code") & "' order by jobp_id"	   
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
			   sql="SELECT top 1 invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_parttype, invd_desc, invd_unitcost, invd_qty, invd_discountamt, invd_discounttype, " & _ 
			        "invd_netcost, invd_subtotal FROM tblinvoice_detail "
				rs.ActiveConnection = strconnect
				rs.Source = sql
				rs.CursorLocation  = 3
				rs.CursorType = 2
				rs.LockType = 2
				rs.Open
				rs.addnew
				rs("invd_inv_no") = inv_no
				rs("invd_job_code") = job_code
				rs("invd_partcode") = rs1("jobp_partcode")
				rs("invd_parttype") = "Parts"
				rs("invd_desc") = rs1("jobp_desc")
				rs("invd_unitcost") = rs1("jobp_unitcost")
				rs("invd_qty") = rs1("jobp_qty")
				rs("invd_discountamt") = rs1("jobp_discountamt")
				rs("invd_discounttype") = rs1("jobp_discounttype")
				rs("invd_netcost") = rs1("jobp_netcost")
				rs("invd_subtotal") = rs1("jobp_subtotal")
				rs.update
				rs.close
		rs1.movenext
		wend
		rs1.close
	
		'''add Labour Charges
		if ccur(job_totallabourAmt) > 0 then 
			   sql="SELECT top 1 invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_parttype, invd_desc, invd_unitcost, invd_qty, invd_discountamt, invd_discounttype, " & _ 
			        "invd_netcost, invd_subtotal FROM tblinvoice_detail "
				rs.ActiveConnection = strconnect
				rs.Source = sql
				rs.CursorLocation  = 3
				rs.CursorType = 2
				rs.LockType = 2
				rs.Open
				rs.addnew
				rs("invd_inv_no") = inv_no
				rs("invd_job_code") = job_code
				rs("invd_partcode") = "Labour"
				rs("invd_parttype") = "Labour"
				rs("invd_desc") = "Labour Charges"
				rs("invd_unitcost") = job_totallabourAmt
				rs("invd_qty") = 1
				rs("invd_discountamt") = 0
				rs("invd_discounttype") = "%"
				rs("invd_netcost") = job_totallabourAmt
				rs("invd_subtotal") = job_totallabourAmt
				rs.update
				rs.close
		end if
		
		'''add Transport Charges  
		if ccur(job_totaltransportAmt) > 0 then 
			   sql="SELECT top 1 invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_parttype, invd_desc, invd_unitcost, invd_qty, invd_discountamt, invd_discounttype, " & _ 
			        "invd_netcost, invd_subtotal FROM tblinvoice_detail "
				rs.ActiveConnection = strconnect
				rs.Source = sql
				rs.CursorLocation  = 3
				rs.CursorType = 2
				rs.LockType = 2
				rs.Open
				rs.addnew
				rs("invd_inv_no") = inv_no
				rs("invd_job_code") = job_code
				rs("invd_partcode") = "Transport"
				rs("invd_parttype") = "Transport"
				rs("invd_desc") = "Transport Charges"
				rs("invd_unitcost") = job_totaltransportAmt
				rs("invd_qty") = 1
				rs("invd_discountamt") = 0
				rs("invd_discounttype") = "%"
				rs("invd_netcost") = job_totaltransportAmt
				rs("invd_subtotal") = job_totaltransportAmt
				rs.update
				rs.close
		end if
		
        sql = "select sum(invd_subtotal) as invd_subtotal from tblinvoice_detail where invd_inv_no = '" & inv_no & "' and invd_parttype='Parts'"
        inv_totalPartsAmt = selectid(sql)
		if isnull(inv_totalPartsAmt) then 
		   inv_totalPartsAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as invd_subtotal from tblinvoice_detail where invd_inv_no = '" & inv_no & "' and invd_parttype='Labour'"
        inv_labourAmt = selectid(sql)
		if isnull(inv_labourAmt) then 
		   inv_labourAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as invd_subtotal from tblinvoice_detail where invd_inv_no = '" & inv_no & "' and invd_parttype='Transport'"
        inv_transportAmt = selectid(sql)
		if isnull(inv_transportAmt) then 
		   inv_transportAmt = 0
		end if
		
		sql = "select sum(invd_subtotalcost) as invd_subtotalcost from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype = 'Parts'"
        invd_subtotalcost = selectid(sql)
		if isnull(invd_subtotalcost) then 
		   invd_subtotalcost = 0
		end if
		
		invd_subtotal = inv_totalPartsAmt + inv_labourAmt + inv_transportAmt
        inv_gstAmt = invd_subtotal * GSTRateBack
		
		sql = "update tblinvoice set inv_totalPartsAmt=" & inv_totalPartsAmt & ", inv_labourAmt=" & inv_labourAmt & ", " & _
		      "inv_transportAmt=" & inv_transportAmt & ", inv_gstAmt = " & chknumber2(inv_gstAmt) & ", " & _
              "inv_totalAmt=" & invd_subtotal & ", " & _
			  "inv_totalPartCost=" & invd_subtotalcost & " " & _
			  "where inv_no = '" & inv_no & "'"
        CUD(sql)
		
		sql = "update tblinvoice set inv_balance=inv_totalAmt-inv_payment-inv_cnamount where inv_no = '" & inv_no & "'"
		CUD(sql)
		
		 if request.Cookies("GAPS")("slevel") = "technician2" then 
		 url = "tech2_jobsheet_new.asp?job_code=" & request("job_code") & "&loginerr=Job's invoice has been generated.#articletitle" 
		 else
		 url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=Job's invoice has been generated.#articletitle" 
		 end if
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob','generateinvoicejob=" & ChkString(left(request("job_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
    else    
         
		 url = "rm_jobsheet.asp?job_code=" & request("job_code") & "&loginerr=Job's invoice already generated.#articletitle" 
  
    end if

'----------------------------------------------------------------------------------------------------     
    Case "addinvoice"   
	
	if request.form("inv_cust_cnty_id") <> "" and request.form("inv_cust_postcode") = "" then 'just return regardless country
		inv_cust_cnty_id = request("inv_cust_cnty_id")
		inv_cust_name	=  request("inv_cust_name")
		Response.Redirect "rm_invoice_new.asp?inv_no="&inv_no&"&inv_cust_postcode="&inv_cust_postcode&"&inv_cust_name="&inv_cust_name&"&inv_cust_address="&inv_cust_address&"&inv_cust_tel1="&inv_cust_tel1&"&inv_cust_tel2="&inv_cust_tel2&"&inv_cust_cnty_id="&inv_cust_cnty_id& "&loginerr=Updated Address.#articletitle" 
	end if

	if request.form("inv_cust_postcode") <> ""  and request.form("inv_cust_code") = "" and request.form("inv_cust_cnty_id") = "129"  then
	  if request.form("inv_cust_city_id") = "" and request.form("inv_cust_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			inv_cust_postcode=request("inv_cust_postcode")
			inv_cust_name=request("inv_cust_name")
			inv_cust_tel1=request("inv_cust_tel1")
			inv_cust_tel2=request("inv_cust_tel2")
			inv_cust_address=request("inv_cust_address")
			inv_cust_cnty_id = request.form("inv_cust_cnty_id")

			sql = "select city_id from tblpostcode where postcode =" & request("inv_cust_postcode")	
			inv_cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & inv_cust_city_id	
			inv_cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("inv_cust_postcode")	
			inv_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("inv_cust_postcode")	
			inv_cust_state = selectid(sql)

 		    Response.Redirect "rm_invoice_new.asp?inv_no="&inv_no&"&inv_cust_postcode="&inv_cust_postcode&"&inv_cust_name="&inv_cust_name&"&inv_cust_address="&inv_cust_address&"&inv_cust_tel1="&inv_cust_tel1&"&inv_cust_tel2="&inv_cust_tel2&"&inv_cust_cnty_id="&inv_cust_cnty_id& "&loginerr=Updated Address.#articletitle" 
	   end if
	end if
	
	if  request.form("inv_cust_cnty_id") <> "129" then 
			inv_cust_city = request("inv_cust_city")
	   	    inv_cust_city_id = "0"
	end if
	
        ''''Add Invoice 	   	  
        sql="SELECT top 1 inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
			"inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_job_code, inv_tech_code,  " & _
			"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
			"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_postedby, inv_posteddate " & _
			"FROM tblinvoice "	
	    set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2
		rs.Open
        rs.AddNew   
        rs("inv_status") = "Open"
		rs("inv_date") = ChkDateYYYYMMDD(date())
        rs("inv_cust_code")  = ChkString(Request.Form("inv_cust_code"))	
        rs("inv_cust_name")  = ChkString(Request.Form("inv_cust_name"))
		rs("inv_cust_address") = ChkString(Request.Form("inv_cust_address"))
		rs("inv_cust_postcode") = ChkString(Request.Form("inv_cust_postcode"))
		
		if request.form("inv_cust_cnty_id") = "129" then 'state applies to Malaysia only
			rs("inv_cust_state") =  ChkString(Request.Form("inv_cust_state"))
			rs("inv_cust_state_id") = ChkString(Request.Form("inv_cust_state_id")) 
		end if
		
		rs("inv_cust_city") =  ChkString(Request.Form("inv_cust_city"))
		rs("inv_cust_city_id") = ChkString(Request.Form("inv_cust_city_id"))
		
		rs("inv_cust_cnty_id") = ChkString(Request.Form("inv_cust_cnty_id")) 
		rs("inv_cust_email") = ChkString(Request.Form("inv_cust_email")) 
		rs("inv_cust_tel1") = ChkString(Request.Form("inv_cust_tel1")) 
		rs("inv_cust_tel2") = ChkString(Request.Form("inv_cust_tel2")) 
		rs("inv_remark") = ChkString(Request.Form("inv_remark")) 
		rs("inv_createddate") = ChkDateTimeMySQL(now())
		rs("inv_createdby") = Request.Cookies("GAPS")("sloginid")
		rs("inv_tech_code")  = ChkString(Request.Form("inv_tech_code"))	
		rs("inv_no") = "temp-" & ChkDateTimeMySQL(now()) ' 280724 temp number as inv_no cannot be null
		rs.Update 
		rs.Close   
		
        sql = "select top 1 inv_id from tblinvoice order by inv_id desc "
        inv_id = selectid(sql)		
        inv_no = "RS" & InvoiceNumbering(inv_id)
  
        sql = "update tblinvoice set inv_no = '" & inv_no & "' where inv_id = " & inv_id
        CUD(sql)
		
		''' Update Customer Profile
        sql1 = "SELECT top 1 cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
			  "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
			  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic,cust_type,cust_debtor_code " & _
			  "FROM tblcustomer where cust_code = '" & ChkString(Request.Form("inv_cust_code")) & "' "	

	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if rs1.eof then
			rs1.addnew
			rs1("cust_createdby") = Request.Cookies("GAPS")("sloginid")
			rs1("cust_createddate") = ChkDateTimeMySQL(now())
			rs1("cust_name")  = ChkString(Request.Form("inv_cust_name"))
			rs1("cust_address") = ChkString(Request.Form("inv_cust_address"))
			rs1("cust_postcode") = ChkString(Request.Form("inv_cust_postcode"))
			
			if request.form("inv_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs1("cust_state") = ChkString(Request.Form("inv_cust_state"))
				rs1("cust_state_id") = ChkString(Request.Form("inv_cust_state_id")) 
			end if
			
			rs1("cust_city") = ChkString(Request.Form("inv_cust_city"))
			rs1("cust_city_id") = ChkString(Request.Form("inv_cust_city_id"))
			rs1("cust_cnty_id") = ChkString(Request.Form("inv_cust_cnty_id")) 
			rs1("cust_email") = ChkString(Request.Form("inv_cust_email")) 
			rs1("cust_tel1") = ChkString(Request.Form("inv_cust_tel1")) 
			rs1("cust_tel2") = ChkString(Request.Form("inv_cust_tel2")) 
			rs1("cust_type") = "Customer" '080226 assumed entered data is customer
			rs1("cust_debtor_code") = "300-CD10" '080226 assumed entered data is customer
		    'rs1("inv_no") = "temp-" & ChkDateTimeMySQL(now())
			rs1.Update 
		    rs1.Close 
			
			sql = "select top 1 cust_id from tblcustomer order by cust_id desc "
			cust_id = selectid(sql)
			if cust_id <> "" then
				temp = 100000 + cust_id
			else
				temp = 100000 'first time
			end if
			cust_code = "C" & temp 
			sql = "update tblcustomer set cust_code = '" & cust_code & "' where cust_id = " & cust_id
			CUD(sql)
			
			sql = "update tblinvoice set inv_cust_code = '" & cust_code & "' where inv_id = " & inv_id
            CUD(sql)		
		else
		if len(ChkString(Request.Form("inv_cust_postcode"))) = "5" or len(ChkString(Request.Form("inv_cust_postcode"))) = "6"  then 
			rs1("cust_createdby") = Request.Cookies("GAPS")("sloginid")
			rs1("cust_createddate") = ChkDateTimeMySQL(now())	
			rs1("cust_name")  = ChkString(Request.Form("inv_cust_name"))
			rs1("cust_address") = ChkString(Request.Form("inv_cust_address"))
			rs1("cust_postcode") = ChkString(Request.Form("inv_cust_postcode"))

			if request.form("job_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs1("cust_state") = ChkString(Request.Form("inv_cust_state"))
				rs1("cust_state_id") = ChkString(Request.Form("inv_cust_state_id")) 
			end if
			
			rs1("cust_city") = ChkString(Request.Form("inv_cust_city"))
			rs1("cust_city_id") = ChkString(Request.Form("inv_cust_city_id")) 
			rs1("cust_cnty_id") = ChkString(Request.Form("inv_cust_cnty_id")) 
			rs1("cust_email") = ChkString(Request.Form("inv_cust_email")) 
			rs1("cust_tel1") = ChkString(Request.Form("inv_cust_tel1")) 
			rs1("cust_tel2") = ChkString(Request.Form("inv_cust_tel2")) 
			rs1.Update 
		end if
		    rs1.Close  
		end if
		
      
        url = "rm_invoice_new.asp?inv_no=" & inv_no & "&loginerr=New Invoice has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','addInvoice=" & ChkString(left(inv_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editinvoice"   
	
	    'sql = "select ct_name from tblcity where ct_id =" & request("inv_cust_city") 
		'inv_cust_city = selectid(sql)
		
		'sql = "select state_name from tblstate where state_id =" & request("inv_cust_state") 
		'inv_cust_state = selectid(sql)
		
		'sql = "select state_code from tblstate where state_id =" & request("inv_cust_state") 
		'state_code = selectid(sql)
	

		if  request.form("inv_cust_cnty_id") <> "129" then 
			inv_cust_city = request.form("inv_cust_city")
	   	    inv_cust_city_id = "0"
		end if

        ''''Edit Invoice   	  
        sql="SELECT top 1 inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
			"inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_job_code, inv_tech_code,  " & _
			"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
			"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_postedby, inv_posteddate " & _
			"FROM tblinvoice where inv_no = '" & request("inv_no") & "' "	

	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("inv_cust_code")  = ChkString(Request.Form("inv_cust_code"))	
			rs("inv_cust_name")  = ChkString(Request.Form("inv_cust_name"))
			rs("inv_cust_address") = ChkString(Request.Form("inv_cust_address"))
			rs("inv_cust_postcode") = ChkString(Request.Form("inv_cust_postcode"))

			if request.form("inv_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs("inv_cust_state") =  ChkString(Request.Form("inv_cust_state"))
				rs("inv_cust_state_id") = ChkString(Request.Form("inv_cust_state_id")) 
			end if
			
			rs("inv_cust_city") =  ChkString(Request.Form("inv_cust_city"))
			rs("inv_cust_city_id") = ChkString(Request.Form("inv_cust_city_id")) 
			rs("inv_cust_cnty_id") = ChkString(Request.Form("inv_cust_cnty_id")) 
			rs("inv_cust_email") = ChkString(Request.Form("inv_cust_email")) 
			rs("inv_cust_tel1") = ChkString(Request.Form("inv_cust_tel1")) 
			rs("inv_cust_tel2") = ChkString(Request.Form("inv_cust_tel2")) 
			rs("inv_remark") = ChkString(Request.Form("inv_remark")) 
			rs("inv_createddate") = ChkDateTimeMySQL(now())
			rs("inv_createdby") = Request.Cookies("GAPS")("sloginid")
			rs("inv_tech_code") = ChkString(Request.Form("inv_tech_code")) 
		rs.Update 
		rs.Close 
		end if

        url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','editInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	 
		 
'----------------------------------------------------------------------------------------------------    
  Case "addInvoiceDetail"   

if isNotExceed(ChkString(Request.Form("invd_partcode")), ChkString(Request.Form("invd_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
  sql = "select md_averageecost from tblmodel where md_code='" & ChkString(Request.Form("invd_partcode")) & "'"
  md_averageecost = selectid(sql)
  if isnull(md_averageecost) or md_averageecost="" then
     md_averageecost="0"
  end if	 
  
  sql = "select md_costprice from tblmodel where md_code='" & ChkString(Request.Form("invd_partcode")) & "'"
  md_costprice = selectid(sql)
  
  if isnull(md_costprice) or md_costprice="" then
     md_costprice="0"
  end if
  
        ''''Add Inv parts	   	  
        sql = "SELECT top 1 invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_parttype, invd_desc, " & _
		      "invd_unitcost, invd_discountamt, invd_discounttype, invd_netcost, invd_qty, invd_subtotal, invd_itemcost, invd_subtotalcost, invd_averagecost " & _
	          "FROM tblinvoice_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        rs.AddNew   
        rs("invd_inv_no") = ChkString(Request.Form("inv_no"))
        rs("invd_partcode")  = ChkString(Request.Form("invd_partcode"))	
		rs("invd_averagecost") = ChkString(md_averageecost)
		
		if lcase(Request.Form("invd_partcode"))	= "labour" then 
		rs("invd_parttype")  = "Labour"
		elseif lcase(Request.Form("invd_partcode"))	= "transport" then 
		rs("invd_parttype")  = "Transport"
		else
		rs("invd_parttype")  = "Parts"
		end if
		
		rs("invd_desc")  = ChkString(Request.Form("invd_desc"))	
        rs("invd_unitcost")  = ChkString(Request.Form("invd_unitcost"))
		rs("invd_qty") = ChkString(Request.Form("invd_qty"))
		
		if ChkString(Request.Form("invd_discounttype")) = "%" then 
			invd_netcost  = ChkString(Request.Form("invd_unitcost")) * (ChkString(Request.Form("invd_discountamt")/100))
			rs("invd_discountamt")  = ChkString(invd_netcost)
			rs("invd_discounttype")  = "RM"
			invd_netcost = ChkString(Request.Form("invd_unitcost")) - invd_netcost
			rs("invd_netcost")  = invd_netcost
			rs("invd_subtotal")  = invd_netcost * ChkString(Request.Form("invd_qty"))
		else	
			rs("invd_discountamt")  = ChkString(Request.Form("invd_discountamt"))
			rs("invd_discounttype")  = "RM"
			invd_netcost  = ChkString(Request.Form("invd_unitcost")) -  ChkString(Request.Form("invd_discountamt"))
			rs("invd_netcost")  = invd_netcost
			rs("invd_subtotal")  = invd_netcost * ChkString(Request.Form("invd_qty"))
		end if
		
		rs("invd_itemcost") = ChkString(md_costprice)
		rs("invd_subtotalcost") = ChkString(md_costprice) * ChkString(Request.Form("invd_qty"))
		
		rs.Update 
		rs.Close      
		
        sql = "select sum(invd_subtotal) as inv_totalPartsAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Parts')"
        inv_totalPartsAmt = selectid(sql)
		if isnull(inv_totalPartsAmt) then 
		   inv_totalPartsAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as inv_labourAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Labour')"
        inv_labourAmt = selectid(sql)
		if isnull(inv_labourAmt) then 
		   inv_labourAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as inv_transportAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Service', 'Transport')"
        inv_transportAmt = selectid(sql)
		if isnull(inv_transportAmt) then 
		   inv_transportAmt = 0
		end if
		
		invd_subtotal = inv_totalPartsAmt + inv_labourAmt + inv_transportAmt
        inv_gstAmt = invd_subtotal * GSTRateBack
		
		sql = "select sum(invd_subtotalcost) as invd_subtotalcost from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype = 'Parts'"
        invd_subtotalcost = selectid(sql)
		if isnull(invd_subtotalcost) then 
		   invd_subtotalcost = 0
		end if
		
		sql = "update tblinvoice set inv_totalPartsAmt=" & inv_totalPartsAmt & ", inv_labourAmt=" & inv_labourAmt & ", " & _
		      "inv_transportAmt=" & inv_transportAmt & ", inv_gstAmt = " & chknumber2(inv_gstAmt) & ", " & _
              "inv_totalAmt=" & invd_subtotal & ", " & _
			  "inv_totalPartCost=" & invd_subtotalcost & " " & _
			  "where inv_no = '" & request("inv_no") & "'"
        CUD(sql)
		
		sql = "update tblinvoice set inv_balance=inv_totalAmt-inv_payment-inv_cnamount where inv_no = '" & request("inv_no") & "'"
		CUD(sql)
		
        url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinv_parts','addInvoiceDetail=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 else
	   url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has not updated.#spareparts" 
 end if
'----------------------------------------------------------------------------------------------------    
  Case "editInvoiceDetail"   
	if isNotExceed(ChkString(Request.Form("invd_partcode")), ChkString(Request.Form("invd_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
    sql = "delete from tblinvoice_detail where invd_id=" & request("invd_id")	
	    CUD(sql)

	    sql = "select md_averageecost from tblmodel where md_code='" & ChkString(Request.Form("invd_partcode")) & "'"
		  md_averageecost = selectid(sql)
		  if isnull(md_averageecost) or md_averageecost="" then
			 md_averageecost="0"
		  end if
  
		sql = "select md_costprice from tblmodel where md_code='" & ChkString(Request.Form("invd_partcode")) & "'"
	    md_costprice = selectid(sql)
	  
	    if isnull(md_costprice) or md_costprice="" then
		 md_costprice="0"
	    end if
		  		
         ''''Add Inv parts	   	  
        sql = "SELECT top 1 invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_parttype, invd_desc, invd_unitcost, invd_discountamt, " & _
		      "invd_discounttype, invd_netcost, invd_qty, invd_subtotal, invd_itemcost, invd_subtotalcost, invd_averagecost " & _
	          "FROM tblinvoice_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("invd_inv_no") = ChkString(Request.Form("inv_no"))
        rs("invd_partcode")  = ChkString(Request.Form("invd_partcode"))	
		rs("invd_averagecost") = ChkString(md_averageecost)
		
		if lcase(Request.Form("invd_partcode"))	= "labour" then 
		rs("invd_parttype")  = "Labour"
		elseif lcase(Request.Form("invd_partcode"))	= "transport" then 
		rs("invd_parttype")  = "Transport"
		else
		rs("invd_parttype")  = "Parts"
		end if
		
		rs("invd_desc")  = ChkString(Request.Form("invd_desc"))	
        rs("invd_unitcost")  = ChkString(Request.Form("invd_unitcost"))
		rs("invd_qty") = ChkString(Request.Form("invd_qty"))
		
		if ChkString(Request.Form("invd_discounttype")) = "%" then 
			invd_netcost  = ChkString(Request.Form("invd_unitcost")) * (ChkString(Request.Form("invd_discountamt")/100))
			rs("invd_discountamt")  = ChkString(invd_netcost)
			rs("invd_discounttype")  = "RM"
			invd_netcost = ChkString(Request.Form("invd_unitcost")) - invd_netcost
			rs("invd_netcost")  = invd_netcost
			rs("invd_subtotal")  = invd_netcost * ChkString(Request.Form("invd_qty"))
		else	
			rs("invd_discountamt")  = ChkString(Request.Form("invd_discountamt"))
			rs("invd_discounttype")  = "RM"
			invd_netcost  = ChkString(Request.Form("invd_unitcost")) -  ChkString(Request.Form("invd_discountamt"))
			rs("invd_netcost")  = invd_netcost
			rs("invd_subtotal")  = invd_netcost * ChkString(Request.Form("invd_qty"))
		end if
		
		rs("invd_itemcost") = ChkString(md_costprice)
		rs("invd_subtotalcost") = ChkString(md_costprice) * ChkString(Request.Form("invd_qty"))
		
		rs.Update 
		rs.Close      
		
       sql = "select sum(invd_subtotal) as inv_totalPartsAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Parts')"
        inv_totalPartsAmt = selectid(sql)
		if isnull(inv_totalPartsAmt) then 
		   inv_totalPartsAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as inv_labourAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Labour')"
        inv_labourAmt = selectid(sql)
		if isnull(inv_labourAmt) then 
		   inv_labourAmt = 0
		end if
		
		sql = "select sum(invd_subtotal) as inv_transportAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Service', 'Transport')"
        inv_transportAmt = selectid(sql)
		if isnull(inv_transportAmt) then 
		   inv_transportAmt = 0
		end if
		
		invd_subtotal = inv_totalPartsAmt + inv_labourAmt + inv_transportAmt
        inv_gstAmt = invd_subtotal * GSTRateBack		
		
		sql = "select sum(invd_subtotalcost) as invd_subtotalcost from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype = 'Parts'"
        invd_subtotalcost = selectid(sql)
		if isnull(invd_subtotalcost) then 
		   invd_subtotalcost = 0
		end if
		
		sql = "update tblinvoice set inv_totalPartsAmt=" & inv_totalPartsAmt & ", inv_labourAmt=" & inv_labourAmt & ", " & _
		      "inv_transportAmt=" & inv_transportAmt & ", inv_gstAmt = " & chknumber2(inv_gstAmt) & ", " & _
              "inv_totalAmt=" & invd_subtotal & ", " & _
			  "inv_totalPartCost=" & invd_subtotalcost & " " & _
			  "where inv_no = '" & request("inv_no") & "'"
        CUD(sql)
		
		sql = "update tblinvoice set inv_balance=inv_totalAmt-inv_payment-inv_cnamount where inv_no = '" & request("inv_no") & "'"
		CUD(sql)
				
        url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has been updated.#spareparts" 
  
        sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice_detail','editInvoiceDetail=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
        CUD(sql)
else
	        url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has not been updated.#spareparts" 
end if
'----------------------------------------------------------------------------------------------------    
  Case "delInvoiceDetail"
  
	sql = "delete from tblinvoice_detail where invd_id=" & request("invd_id")	
	CUD(sql)
	
    sql = "select sum(invd_subtotal) as inv_totalPartsAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Parts')"
	inv_totalPartsAmt = selectid(sql)
	if isnull(inv_totalPartsAmt) then 
	   inv_totalPartsAmt = 0
	end if
	
	sql = "select sum(invd_subtotal) as inv_labourAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Labour')"
	inv_labourAmt = selectid(sql)
	if isnull(inv_labourAmt) then 
	   inv_labourAmt = 0
	end if
	
	sql = "select sum(invd_subtotal) as inv_transportAmt from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype in ('Service', 'Transport')"
	inv_transportAmt = selectid(sql)
	if isnull(inv_transportAmt) then 
	   inv_transportAmt = 0
	end if
	
	invd_subtotal = inv_totalPartsAmt + inv_labourAmt + inv_transportAmt
	inv_gstAmt = invd_subtotal * GSTRateBack
	
	sql = "select sum(invd_subtotalcost) as invd_subtotalcost from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "' and invd_parttype = 'Parts'"
	invd_subtotalcost = selectid(sql)
	if isnull(invd_subtotalcost) then 
	   invd_subtotalcost = 0
	end if
	
	sql = "update tblinvoice set inv_totalPartsAmt=" & inv_totalPartsAmt & ", inv_labourAmt=" & inv_labourAmt & ", " & _
		  "inv_transportAmt=" & inv_transportAmt & ", inv_gstAmt = " & chknumber2(inv_gstAmt) & ", " & _
		  "inv_totalAmt=" & invd_subtotal & ", " & _
		  "inv_totalPartCost=" & invd_subtotalcost & " " & _
		  "where inv_no = '" & request("inv_no") & "'"
	CUD(sql)
	
	sql = "update tblinvoice set inv_balance=inv_totalAmt-inv_payment-inv_cnamount where inv_no = '" & request("inv_no") & "'"
    CUD(sql)
	
	url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Invoice has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinv_parts','delInvoiceDetail=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "submitInvoice"   
  
sql1 = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_job_code  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
		
		can_be_submit=true

	    'user can only select stk if quantity i smore than zero. Hence no need to check here again.

        if not rs1.eof then 
				sql2 = "select invd_partcode,invd_qty from tblinvoice_detail where invd_inv_no = '" & request("inv_no") & "'"
				set rs2 = server.CreateObject("adodb.recordset")
				rs2.Open sql2,strconnect,3,3,&H0001				
				if not rs2.eof then  '20624 -  - dont allow submit if there are no parts or inv details
					rs1("inv_status") = "Submitted"
					rs1("inv_approvedby") = Request.Cookies("GAPS")("sloginid")
					rs1("inv_approveddate") = ChkDateTimeMySQL(now())
				else
					can_be_submit=false
				end if	
				rs2.close
				rs1.Update 
				rs1.Close 
		end if 

	if can_be_submit = true then
        url = "rm_invoice_view.asp?inv_status=Submitted&inv_no=" & request("inv_no") & "&loginerr=Invoice has been Submitted.#articletitle" 
	else
		url = "rm_invoice_view.asp?inv_status=Submitted&inv_no=" & request("inv_no") & "&loginerr=Invoice has NOT been Submitted.#articletitle"
	end if

    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','submitInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "PostedInvoice"   
  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id,  inv_cust_cnty_id,inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_job_code, inv_postedby, inv_posteddate  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("inv_status") = "Posted"
			   rs("inv_postedby") = Request.Cookies("GAPS")("sloginid")
			   rs("inv_posteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if
		
		
		'''Generate DO
        sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
				"inv_cust_city, inv_cust_city_id,  inv_cust_cnty_id,inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
				"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
				"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code, inv_posteddate, inv_postedby,  " & _
				"inv_payment, inv_balance, inv_payment_type, inv_chequeno, inv_payment_remark  " & _
				"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
		'response.write sql
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			inv_id = rs("inv_id") 
			inv_no = rs("inv_no") 
			inv_date = rs("inv_date") 
			inv_cust_code = rs("inv_cust_code")
			inv_cust_name = rs("inv_cust_name")
			inv_cust_address = rs("inv_cust_address") 
			inv_cust_postcode = rs("inv_cust_postcode") 
			inv_cust_state = rs("inv_cust_state") 
			inv_cust_state_id = rs("inv_cust_state_id") 
			inv_cust_city = rs("inv_cust_city") 
			inv_cust_city_id = rs("inv_cust_city_id") 
			inv_cust_cnty_id = rs("inv_cust_cnty_id")
			inv_cust_email = rs("inv_cust_email") 
			inv_cust_tel1 = rs("inv_cust_tel1") 
			inv_cust_tel2 = rs("inv_cust_tel2")  
			inv_createddate = rs("inv_createddate") 
			inv_createdby = rs("inv_createdby") 
			inv_approveddate = rs("inv_approveddate") 
			inv_approvedby = rs("inv_approvedby") 
			inv_status = rs("inv_status") 
			inv_job_code = rs("inv_job_code")
			
			inv_gstAmt = rs("inv_gstAmt")
			inv_totalPartsAmt = rs("inv_totalPartsAmt")
			inv_totalAmt = rs("inv_totalAmt")
			inv_tech_code = rs("inv_tech_code")
			inv_posteddate = rs("inv_posteddate") 
			inv_postedby = rs("inv_postedby") 
			inv_remark = rs("inv_remark")
			
			inv_payment = rs("inv_payment") 
			inv_balance = rs("inv_balance") 
			inv_payment_type = rs("inv_payment_type") 
			inv_chequeno = rs("inv_chequeno") 
			inv_payment_remark = rs("inv_payment_remark")
		End If
		rs.close
		
		'''Job Detail
        sql = "SELECT top 1 job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city,job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_code, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_wrty_photo2,job_wrty_photo3,job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date, job_submittedby, job_submitteddate, " & _
		"job_doneby, job_donedate, job_postedby, job_posteddate, job_cancelledby, job_cancelleddate  " & _
	    "FROM tbljob WHERE job_code = '" & inv_job_code & "' "
		rs.Open sql,strconnect,0,1,&H0001
		
		If Not rs.EOF Then
			job_code = rs("job_code")
			job_purchase_date = rs("job_purchase_date") 
			job_onlineWrtyNo = rs("job_onlineWrtyNo") 
			job_onlineWrtyStatus = chkstring(rs("job_onlineWrtyStatus"))
			job_type = rs("job_type") 
			job_SN_no = rs("job_SN_no") 
			job_Model = rs("job_Model") 
			job_model_desc = rs("job_model_desc") 
			job_faulty_reason_cs = rs("job_faulty_reason_cs") 
			job_faulty_desc = rs("job_faulty_desc") 
			job_reportedby = rs("job_reportedby") 
			
			if isdate(rs("job_appointment_date")) then 
			job_appointment_date = rs("job_appointment_date") 
			end if
			
			job_appointment_time = rs("job_appointment_time") 
			job_tech_code = rs("job_tech_code") 
			job_appointment_remark = rs("job_appointment_remark") 
		End If
		rs.close
		
		service = false
		if job_faulty_reason_cs = "Installation" then '02122024 do not create any DO and exit as this is a service job only
				service= true 
		end if

        ''''Add DO Order	  
		addDO = "N" 	  
        
		if service = false then '02122024 execute if its not service (DO meant for parts)
		sql = "SELECT top 1 do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
			"do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_cnty_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
			"do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
			"do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
			"do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
			"do_appointment_remark, do_remark FROM tbldo where do_inv_no='" & inv_no & "' "		

		
		set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
		if rs.eof then 
        rs.AddNew  
		do_no = replace(inv_no, "RS", "DOS") '050924 new inv format  
		rs("do_no")  = do_no
		rs("do_status")  = "Posted"   					
        rs("do_date") = ChkDateYYYYMMDD(date())
        rs("do_cust_code")  = ChkString(inv_cust_code)	
        rs("do_cust_name")  = ChkString(inv_cust_name)
		rs("do_cust_address") = ChkString(inv_cust_address)
		rs("do_cust_postcode") = ChkString(inv_cust_postcode)
		rs("do_cust_state") = inv_cust_state
		rs("do_cust_state_id") = ChkString(inv_cust_state_id) 
		rs("do_cust_city") = inv_cust_city
		rs("do_cust_city_id") = ChkString(inv_cust_city_id) 
		rs("do_cust_cnty_id") = ChkString(inv_cust_cnty_id) 
		rs("do_cust_email") = ChkString(inv_cust_email) 
		rs("do_cust_tel1") = ChkString(inv_cust_tel1) 
		rs("do_cust_tel2") = ChkString(inv_cust_tel2) 
		'' rs("do_remark") = "Auto Generate DO based on Invoice: " & request("inv_no")
		rs("do_remark") = ChkString(inv_remark)
		rs("do_createddate") = ChkDateTimeMySQL(now())
		rs("do_createdby") = Request.Cookies("GAPS")("sloginid")
		rs("do_tech_code") = ChkString(inv_tech_code) 
		
		if ChkString(job_purchase_date)  <> "" then 
		rs("do_purchase_date") = ChkString(job_purchase_date) 
		end if
		
		rs("do_onlineWrtyNo") = ChkString(job_onlineWrtyNo) 
		rs("do_onlineWrtyStatus") = ChkString(job_onlineWrtyStatus) 
		rs("do_SN_no") = ChkString(job_SN_no) 
		rs("do_type") = ChkString(job_type) 
		rs("do_Model") = ChkString(job_Model) 
		rs("do_model_desc") = ChkString(job_model_desc) 
		
		if ChkString(job_appointment_date)  <> "" then
		rs("do_appointment_date") = ChkString(job_appointment_date) 
		end if
		rs("do_appointment_time") = ChkString(job_appointment_time) 
		rs("do_appointment_remark") = ChkString(job_appointment_remark) 
		rs("do_inv_no") = ChkString(inv_no) 
		
		if ChkString(inv_date)  <> "" then
		rs("do_inv_date") = ChkString(inv_date) 
		end if
		addDO = "Y" 
		rs.Update 
		end if
		rs.Close     
		
		''Generate DO Number
		'sql = "select top 1 do_id from tbldo order by do_id desc "
        'do_id = selectid(sql)
		'temp = 100000 + do_id
        'do_no = "DO" & temp 
		'sql = "update tbldo set do_no='" & do_no & "' where do_id = '" & do_id & "'"
        'CUD(sql)
		
		''DO Detail
		if addDO = "Y" then 
		sql1 = "SELECT  tblmodel.md_category,invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_desc, invd_unitcost, invd_qty, invd_discountamt, " & _
			"invd_discounttype, invd_netcost, invd_subtotal	FROM tblinvoice_detail " & _
			"inner join tblmodel on tblinvoice_detail.invd_partcode = tblmodel.md_code " & _
			"where invd_inv_no = '" & inv_no & "' order by invd_id"	

		set rs1 = server.CreateObject("adodb.recordset")
		set rs4 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
	
		while Not rs1.EOF
		
				''''Add DO Detail	   	  
				sql2 = "SELECT top 1 dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, dod_discounttype, " & _
					   "dod_netcost, dod_subtotal FROM tbldo_detail "	
				'set rs = server.CreateObject("adodb.recordset")
				rs4.ActiveConnection = strconnect
				rs4.Source = sql2
				rs4.CursorLocation  = 3
				rs4.CursorType = 2
				rs4.LockType = 2
				rs4.Open
				rs4.AddNew   
				rs4("dod_do_no") = ChkString(do_no)
				rs4("dod_inv_no") = ChkString(rs1("invd_inv_no"))
				rs4("dod_job_code")  = ChkString(rs1("invd_job_code"))	
				rs4("dod_partcode")  = ChkString(rs1("invd_partcode"))	
				rs4("dod_desc")  = ChkString(rs1("invd_desc"))	
				rs4("dod_unitcost")  = ChkString(rs1("invd_unitcost"))
				rs4("dod_qty") = ChkString(rs1("invd_qty"))
				rs4("dod_discountamt")  = ChkString(rs1("invd_discountamt"))
				rs4("dod_discounttype")  = ChkString(rs1("invd_discounttype"))
				rs4("dod_netcost")  = ChkString(rs1("invd_netcost"))
				rs4("dod_subtotal")  = ChkString(rs1("invd_subtotal"))
				rs4.Update 
				rs4.Close  
				
				
				if isnull(inv_job_code) then 
					if rs1("md_category") <> "Service" then '12/03/2025 ignore if its installation or postage items
						'Update Stocktrans - Stock Movement
						sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
							   "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
						rs2.ActiveConnection = strconnect
						rs2.Source = sql2
						rs2.CursorLocation  = 3
						rs2.CursorType = 2
						rs2.LockType = 2
						rs2.Open
						rs2.AddNew   
						rs2("stk_voucherno") = ChkString(do_no)
						rs2("stk_reference") = "W1"
						rs2("stk_date")  = ChkDateTimeMySQL(now())
						rs2("stk_type")  = "DO"
						rs2("stk_itm_code")  = ChkString(rs1("invd_partcode"))
						rs2("stk_fromwarehouse")  = "W1"
						rs2("stk_towarehouse")  = ""
						rs2("stk_desc")  = ChkString(rs1("invd_desc"))	
						rs2("stk_qty")  = ChkNumber(rs1("invd_qty")*-1)
						rs2("stk_balanceqty")  = 0
						rs2("stk_sales_price")  = ChkNumber(rs1("invd_subtotal"))
						rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
						rs2("stk_logdate")  = ChkDateTimeMySQL(now())
						rs2.Update 
						rs2.Close   

						'for walk-in cases no job_sheet, hence deduct qty from tblstock for W1 --09122024
						'this is regardless even if the salesman buys the part to sell
						sql = "select wst_itm_current_qty from tblwarehouse_stock where wst_wh_code='W1' and wst_itm_code=  '" & ChkString(rs1("invd_partcode")) & "'" 
						stkcurrbalance= selectid(sql)
						stkcurrbalance = stkcurrbalance + ChkNumber(rs1("invd_qty")*-1)
						sql = "update tblwarehouse_stock set wst_itm_current_qty='" & stkcurrbalance & "',  wst_lastupdatedate = '" & ChkDateTimeMySQL(now()) & "' where wst_wh_code='W1' and wst_itm_code= '" & ChkString(rs1("invd_partcode")) & "'"
						CUD(sql)
						sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
						Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedInvoice_WalkIn=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
						CUD(sql)
					end if
				end if				
		rs1.movenext
		wend
		rs1.close
		
		
		sql = "select sum(dod_qty) as dod_qty from tbldo_detail where dod_do_no = '" & do_no & "'"
		dod_qty = selectid(sql)		
		
		sql = "update tbldo set do_totalqty=" & chknumber2(dod_qty) & " where do_no = '" & do_no & "'"
        CUD(sql)
		
		sql = "update tblinvoice set inv_dono='" & do_no & "', inv_dodate='" & ChkDateYYYYMMDD(date()) & "' WHERE inv_no = '" & request("inv_no") & "'"	
		CUD(sql)
		
		end if
		end if

        url = "rm_invoice_view.asp?inv_status=Posted&inv_no=" & request("inv_no") & "&loginerr=Invoice has been Posted.#articletitle" 
  
        sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','PostedInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
        CUD(sql)	

'----------------------------------------------------------------------------------------------------    
  Case "CancelInvoice"   
  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_job_code, inv_postedby, inv_posteddate, inv_cancelledby,  inv_cancelleddate  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("inv_status") = "Cancel"
			   rs("inv_cancelledby") = Request.Cookies("GAPS")("sloginid")
			   rs("inv_cancelleddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_invoice_view.asp?inv_status=Cancel&inv_no=" & request("inv_no") & "&loginerr=Invoice has been Cancelled.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','CancelInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	

'----------------------------------------------------------------------------------------------------    
  Case "RevertInvoice"   
  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_job_code, inv_postedby, inv_posteddate, inv_cancelledby,  inv_cancelleddate, inv_lastupdateby, inv_lastupdatedate " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("inv_status") = "Open"
			   rs("inv_lastupdateby") = Request.Cookies("GAPS")("sloginid")
			   rs("inv_lastupdatedate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_invoice_view.asp?inv_status=Open&inv_no=" & request("inv_no") & "&loginerr=Invoice has been Reverted to Open.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','RevertInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	

'----------------------------------------------------------------------------------------------------    
  Case "paymentInvoice"   '270624 -  - this function used to generate single receipt upon clicking the inv from view invoice
  
		
		sql = "Update tblinvoice set inv_payment=inv_payment+" & request("inv_payment") & " where inv_no = '" & request("inv_no") & "' "
		CUD(sql)
		
		sql = "Update tblinvoice set inv_balance=inv_totalAmt-(inv_cnamount+inv_payment)," & _
		      "inv_payment_type='" & request("inv_payment_type") & "', inv_chequeno='" & request("inv_chequeno") & "', " & _
			  "inv_payment_remark='" & request("inv_payment_remark") & "' where inv_no = '" & request("inv_no") & "' "
		CUD(sql)	
		
		'''Generate Receipt
	if request("inv_no") <> "" then	  
		sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
				"inv_cust_city, inv_cust_city_id,inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
				"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
				"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code, inv_posteddate, inv_postedby, inv_cnamount, " & _
				"inv_payment, inv_balance, inv_payment_type, inv_chequeno, inv_payment_remark, inv_dono, inv_dodate  " & _
				"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
				'response.write sql
				set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,0,1,&H0001
				If Not rs.EOF Then
					inv_id = rs("inv_id") 
					inv_no = rs("inv_no") 
					inv_date = rs("inv_date") 
					inv_cust_code = rs("inv_cust_code")
					inv_cust_name = rs("inv_cust_name")
					inv_cust_address = rs("inv_cust_address") 
					inv_cust_postcode = rs("inv_cust_postcode") 
					inv_cust_state = rs("inv_cust_state") 
					inv_cust_state_id = rs("inv_cust_state_id") 
					inv_cust_city = rs("inv_cust_city") 
					inv_cust_city_id = rs("inv_cust_city_id") 
					inv_cust_cnty_id = rs("inv_cust_cnty_id") 
					inv_cust_email = rs("inv_cust_email") 
					inv_cust_tel1 = rs("inv_cust_tel1") 
					inv_cust_tel2 = rs("inv_cust_tel2")  
					inv_createddate = rs("inv_createddate") 
					inv_createdby = rs("inv_createdby") 
					inv_approveddate = rs("inv_approveddate") 
					inv_approvedby = rs("inv_approvedby") 
					inv_status = rs("inv_status") 
					inv_job_code = rs("inv_job_code")
					
					inv_gstAmt = rs("inv_gstAmt")
					inv_totalPartsAmt = rs("inv_totalPartsAmt")
					inv_totalAmt = rs("inv_totalAmt")
					inv_tech_code = rs("inv_tech_code")
					inv_posteddate = rs("inv_posteddate") 
					inv_postedby = rs("inv_postedby") 
					inv_remark = rs("inv_remark")
					
					inv_payment = rs("inv_payment") 
					inv_cnamount = rs("inv_cnamount") 
					inv_balance = rs("inv_balance") 
					inv_payment_type = rs("inv_payment_type") 
					inv_chequeno = rs("inv_chequeno") 
					inv_payment_remark = rs("inv_payment_remark")
					inv_dono = rs("inv_dono")
					inv_dodate = rs("inv_dodate")
				End If
				rs.Close
				

			sql = "select top 1 receipt_no from tblreceipt order by receipt_no desc"
			receipt_no = selectid(sql)
			if receipt_no <> "" then 
				receipt_no = Replace(receipt_no, "RN", "") 
				temp = 1 + receipt_no
			else
				temp = 100000 + receipt_no '300724 required for 1st time receipt generation
			end if
			receipt_no = "RN" & temp 			
			

			sql = "SELECT top 1 receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
					"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,receipt_cust_cnty_id,  " & _
					"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
					"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
					"FROM tblreceipt "
			set rs = server.CreateObject("adodb.recordset")
			rs.ActiveConnection = strconnect
			rs.Source = sql
			rs.CursorLocation  = 3
			rs.CursorType = 2
			rs.LockType = 2
			rs.Open
			rs.AddNew    
			rs("receipt_status")  = "Posted"   					
			rs("receipt_date") = ChkDateYYYYMMDD(date())
			rs("receipt_cust_code")  = ChkString(inv_cust_code)	
			rs("receipt_cust_name")  = ChkString(inv_cust_name)
			rs("receipt_cust_address") = ChkString(inv_cust_address)
			rs("receipt_cust_postcode") = ChkString(inv_cust_postcode)
			rs("receipt_cust_state") = inv_cust_state
			rs("receipt_cust_state_id") = ChkString(inv_cust_state_id) 
			rs("receipt_cust_city") = inv_cust_city
			rs("receipt_cust_city_id") = ChkString(inv_cust_city_id) 
			rs("receipt_cust_cnty_id") = ChkString(inv_cust_cnty_id) 
			rs("receipt_cust_email") = ChkString(inv_cust_email) 
			rs("receipt_cust_tel1") = ChkString(inv_cust_tel1) 
			rs("receipt_cust_tel2") = ChkString(inv_cust_tel2) 
			rs("receipt_createddate") = ChkDateTimeMySQL(now())
			rs("receipt_createdby") = Request.Cookies("GAPS")("sloginid")
			rs("receipt_remark") = request("inv_payment_remark")
			rs("receipt_paymenttype") = request("inv_payment_type")
			rs("receipt_totalpayment") = request("inv_payment")
			rs("receipt_inv_no") = inv_no 
			rs("receipt_inv_date") = inv_date 
			rs("receipt_job_code") = inv_job_code 
			rs("receipt_no") = receipt_no 
			rs.Update 
			rs.Close      		
	end if 

		 url = "rm_invoice_new.asp?inv_no=" & request("inv_no") & "&loginerr=Payment has been updated.#articletitle" 
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblreceipt','paymentInvoice=" & ChkString(left(request("inv_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 		 		 
'----------------------------------------------------------------------------------------------------    
  Case "GenReceiptAll"

	if request("GenReceipt") <> "" then
	   inv_list = request("GenReceipt")
	   inv_list = replace(inv_list, " ", "")
	   arrinv_list = split(inv_list,",")
	else
	   inv_list = ""
	   arrinv_list = split("0,0",",") 'occurs when user clicks gen without ticking any box
	end if

	if request("GenReceipt") <> "" then  
		'Generate a single receipt for multiple invoice
		sql = "select top 1 receipt_no from tblreceipt order by receipt_no desc"
		receipt_no = selectid(sql)
		if receipt_no <> "" then 
				receipt_no = Replace(receipt_no, "RN", "") 
				temp = 1 + receipt_no
		else
				temp = 100000 + receipt_no '300724 required for 1st time receipt generation
		end if
		receipt_no = "RN" & temp 

		for k = 0 to ubound(arrinv_list)
			inv_no = arrinv_list(k)
		
		if inv_no <> "" then
				sql = "Update tblinvoice set inv_payment=inv_payment+inv_balance where inv_no = '" & inv_no & "' "
				CUD(sql)
				sql = "Update tblinvoice set inv_balance=inv_totalAmt-(inv_cnamount+inv_payment),inv_payment_type='BankIn' where inv_no = '" & inv_no & "' "
				CUD(sql)

				sql1 = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
				"inv_cust_city, inv_cust_city_id,inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
				"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
				"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code, inv_posteddate, inv_postedby, inv_cnamount, " & _
				"inv_payment, inv_balance, inv_payment_type, inv_chequeno, inv_payment_remark, inv_dono, inv_dodate  " & _
				"FROM tblinvoice WHERE inv_no = '" & inv_no & "' "
				
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,0,1,&H0001
				if Not rs1.EOF Then
					inv_id = rs1("inv_id") 
					inv_no = rs1("inv_no") 
					inv_date = rs1("inv_date") 
					inv_cust_code = rs1("inv_cust_code")
					inv_cust_name = rs1("inv_cust_name")
					inv_cust_address = rs1("inv_cust_address") 
					inv_cust_postcode = rs1("inv_cust_postcode") 
					inv_cust_state = rs1("inv_cust_state") 
					inv_cust_state_id = rs1("inv_cust_state_id") 
					inv_cust_city = rs1("inv_cust_city") 
					inv_cust_city_id = rs1("inv_cust_city_id") 
					inv_cust_cnty_id = rs1("inv_cust_cnty_id") 
					inv_cust_email = rs1("inv_cust_email") 
					inv_cust_tel1 = rs1("inv_cust_tel1") 
					inv_cust_tel2 = rs1("inv_cust_tel2")  
					inv_createddate = rs1("inv_createddate") 
					inv_createdby = rs1("inv_createdby") 
					inv_approveddate = rs1("inv_approveddate") 
					inv_approvedby = rs1("inv_approvedby") 
					inv_status = rs1("inv_status") 
					inv_job_code = rs1("inv_job_code")
					
					inv_gstAmt = rs1("inv_gstAmt")
					inv_totalPartsAmt = rs1("inv_totalPartsAmt")
					inv_totalAmt = rs1("inv_totalAmt")
					inv_tech_code = rs1("inv_tech_code")
					inv_posteddate = rs1("inv_posteddate") 
					inv_postedby = rs1("inv_postedby") 
					inv_remark = rs1("inv_remark")
					
					inv_payment = rs1("inv_payment") 
					inv_cnamount = rs1("inv_cnamount") 
					inv_balance = rs1("inv_balance") 
					inv_payment_type = rs1("inv_payment_type") 
					inv_chequeno = rs1("inv_chequeno") 
					inv_payment_remark = rs1("inv_payment_remark")
					inv_dono = rs1("inv_dono")
					inv_dodate = rs1("inv_dodate")		

					sql = "SELECT top 1 receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
					"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
					"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
					"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
					"FROM tblreceipt "

					set rs = server.CreateObject("adodb.recordset")
					rs.ActiveConnection = strconnect
					rs.Source = sql
					rs.CursorLocation  = 3
					rs.CursorType = 2
					rs.LockType = 2
					rs.Open
					rs.AddNew    
						rs("receipt_status")  = "Posted"   					
						rs("receipt_date") = ChkDateYYYYMMDD(date())
						rs("receipt_cust_code")  = ChkString(inv_cust_code)	
						rs("receipt_cust_name")  = ChkString(inv_cust_name)
						rs("receipt_cust_address") = ChkString(inv_cust_address)
						rs("receipt_cust_postcode") = ChkString(inv_cust_postcode)
						rs("receipt_cust_state") = inv_cust_state
						rs("receipt_cust_state_id") = ChkString(inv_cust_state_id) 
						rs("receipt_cust_city") = inv_cust_city
						rs("receipt_cust_city_id") = ChkString(inv_cust_city_id) 
						rs("receipt_cust_cnty_id") = ChkString(inv_cust_cnty_id) 
						rs("receipt_cust_email") = ChkString(inv_cust_email) 
						rs("receipt_cust_tel1") = ChkString(inv_cust_tel1) 
						rs("receipt_cust_tel2") = ChkString(inv_cust_tel2) 
						rs("receipt_createddate") = ChkDateTimeMySQL(now())
						rs("receipt_createdby") = Request.Cookies("GAPS")("sloginid")
						'rs("receipt_remark") = request("inv_payment_remark")
						rs("receipt_paymenttype") = "BankIn" 'auto update as cash payment
						rs("receipt_totalpayment") = inv_payment 'from inv table
						rs("receipt_inv_no") = inv_no 
						rs("receipt_inv_date") = inv_date 
						rs("receipt_job_code") = inv_job_code 
						rs("receipt_no") = receipt_no
					rs.Update 
					rs.Close      
				end if
				rs1.close
		end if 
		next
		 url = "rm_invoice_view.asp?inv_status=" & "Posted" & "&loginerr=Receipt has been updated.#articletitle" 
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	     Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','Multipleinv=" & receipt_no & "','" & ChkDateTimeMySQL(now()) & "')"       
         CUD(sql)
	else
		 url = "rm_invoice_view.asp?inv_status=" & "Posted" & "&loginerr=Receipt has NOT been updated.#articletitle" 
	end if
	

  Case "GenReceipt" '270624 - - this function used to generate bulk receipt upon selecting tick box from view invoice

	if request("GenReceipt") <> "" then
	   inv_list = request("GenReceipt")
	   inv_list = replace(inv_list, " ", "")
	   arrinv_list = split(inv_list,",")
	else
	   inv_list = ""
	   arrinv_list = split("0,0",",") 'occurs when user clicks gen without ticking any box
	end if

	if request("GenReceipt") <> "" then  
		for k = 0 to ubound(arrinv_list)
			inv_no = arrinv_list(k)
	
			sql = "Update tblinvoice set inv_payment=inv_payment+inv_balance where inv_no = '" & inv_no & "' "
			CUD(sql)
			sql = "Update tblinvoice set inv_balance=inv_totalAmt-(inv_cnamount+inv_payment),inv_payment_type='Cash' where inv_no = '" & inv_no & "' "
			CUD(sql)	
		
		'Generate Receipt
	 
		if inv_no <> "" then
				sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
				"inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
				"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
				"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code, inv_posteddate, inv_postedby, inv_cnamount, " & _
				"inv_payment, inv_balance, inv_payment_type, inv_chequeno, inv_payment_remark, inv_dono, inv_dodate  " & _
				"FROM tblinvoice WHERE inv_no = '" & inv_no & "' "
				
				set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,0,1,&H0001
				If Not rs.EOF Then
					inv_id = rs("inv_id") 
					inv_no = rs("inv_no") 
					inv_date = rs("inv_date") 
					inv_cust_code = rs("inv_cust_code")
					inv_cust_name = rs("inv_cust_name")
					inv_cust_address = rs("inv_cust_address") 
					inv_cust_postcode = rs("inv_cust_postcode") 
					inv_cust_state = rs("inv_cust_state") 
					inv_cust_state_id = rs("inv_cust_state_id") 
					inv_cust_city = rs("inv_cust_city") 
					inv_cust_city_id = rs("inv_cust_city_id") 
					inv_cust_cnty_id = rs("inv_cust_cnty_id") 
					inv_cust_email = rs("inv_cust_email") 
					inv_cust_tel1 = rs("inv_cust_tel1") 
					inv_cust_tel2 = rs("inv_cust_tel2")  
					inv_createddate = rs("inv_createddate") 
					inv_createdby = rs("inv_createdby") 
					inv_approveddate = rs("inv_approveddate") 
					inv_approvedby = rs("inv_approvedby") 
					inv_status = rs("inv_status") 
					inv_job_code = rs("inv_job_code")
					
					inv_gstAmt = rs("inv_gstAmt")
					inv_totalPartsAmt = rs("inv_totalPartsAmt")
					inv_totalAmt = rs("inv_totalAmt")
					inv_tech_code = rs("inv_tech_code")
					inv_posteddate = rs("inv_posteddate") 
					inv_postedby = rs("inv_postedby") 
					inv_remark = rs("inv_remark")
					
					inv_payment = rs("inv_payment") 
					inv_cnamount = rs("inv_cnamount") 
					inv_balance = rs("inv_balance") 
					inv_payment_type = rs("inv_payment_type") 
					inv_chequeno = rs("inv_chequeno") 
					inv_payment_remark = rs("inv_payment_remark")
					inv_dono = rs("inv_dono")
					inv_dodate = rs("inv_dodate")			
				End If
				rs.Close
				
			sql = "select top 1 receipt_no from tblreceipt order by receipt_no desc" 'gen new receipt no
			receipt_no = selectid(sql)
			if receipt_no <> "" then 
				receipt_no = Replace(receipt_no, "RN", "") 
				temp = 1 + receipt_no
			else
				temp = 100000 + receipt_no '300724 required for 1st time receipt generation
			end if
			receipt_no = "RN" & temp 

			sql1 = "SELECT top 1 receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
					"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
					"receipt_cust_cnty_id, receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
					"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
					"FROM tblreceipt "
			set rs1 = server.CreateObject("adodb.recordset")
			rs1.ActiveConnection = strconnect
			rs1.Source = sql1
			rs1.CursorLocation  = 3
			rs1.CursorType = 2
			rs1.LockType = 2
			rs1.Open
			rs1.AddNew    

			rs1("receipt_status")  = "Posted"   					
			rs1("receipt_date") = ChkDateYYYYMMDD(date())
			rs1("receipt_cust_code")  = ChkString(inv_cust_code)	
			rs1("receipt_cust_name")  = ChkString(inv_cust_name)
			rs1("receipt_cust_address") = ChkString(inv_cust_address)
			rs1("receipt_cust_postcode") = ChkString(inv_cust_postcode)
			rs1("receipt_cust_state") = inv_cust_state
			rs1("receipt_cust_state_id") = ChkString(inv_cust_state_id) 
			rs1("receipt_cust_city") = inv_cust_city
			rs1("receipt_cust_city_id") = ChkString(inv_cust_city_id) 
			rs1("receipt_cust_cnty_id") = ChkString(inv_cust_cnty_id) 
			rs1("receipt_cust_email") = ChkString(inv_cust_email) 
			rs1("receipt_cust_tel1") = ChkString(inv_cust_tel1) 
			rs1("receipt_cust_tel2") = ChkString(inv_cust_tel2) 
			rs1("receipt_createddate") = ChkDateTimeMySQL(now())
			rs1("receipt_createdby") = Request.Cookies("GAPS")("sloginid")
			'rs("receipt_remark") = request("inv_payment_remark")
			rs1("receipt_paymenttype") = "BankIn" 'auto update as cash payment
			rs1("receipt_totalpayment") = inv_payment 'from inv table
			rs1("receipt_inv_no") = inv_no 
			rs1("receipt_inv_date") = inv_date 
			rs1("receipt_job_code") = inv_job_code 
			rs1("receipt_no") = receipt_no 
			rs1.Update 
			rs1.Close      
	
		end if 
		next
	end if
		 url = "rm_invoice_view.asp?inv_status=" & "Posted" & "&loginerr=Receipt has been updated.#articletitle" 
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	     Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblinvoice','SingleInv=" & left(inv_list,200) & "','" & ChkDateTimeMySQL(now()) & "')"       
         CUD(sql)
'--------------------------------------------------------------------------------------------------

  Case "addCustomer"   
	
	if request.form("cust_cnty_id") = "" or request.form("cust_postcode") = "" or request.form("cust_tel1") = "" or  request("cust_name") = "" then 'just return regardless country
		cust_name=request("cust_name")
		cust_icno=request("cust_icno")
		cust_cnty_id=request("cust_cnty_id")
		Response.Redirect "rm_customer_new.asp?cust_name="&cust_name&"&cust_icno="&cust_icno&"&cust_cnty_id="&cust_cnty_id&"&loginerr=Enter Country/Postcode/Tel No/Name.#articletitle" 
	end if

	if request.form("cust_cnty_id") = "" or request.form("cust_postcode") = "" then
		cust_name=request("cust_name")
		cust_icno=request("cust_icno")		
		Response.Redirect "rm_customer_new.asp?cust_name="&cust_name&"&cust_icno="&cust_icno&"&cust_cnty_id="&cust_cnty_id&"&loginerr=Enter Country.#articletitle" 
	end if

	'if request("cust_postcode")= "" and  request("cust_name") = "" then '300724 occurs when user saves a blank form
	'    response.redirect("rm_customer_new.asp?type=reset")
	'end if

	if request.form("cust_postcode") <> "" and request.form("cust_cnty_id") = "129" then
			cust_postcode=request("job_cust_postcode")

			sql = "select city_id from tblpostcode where postcode =" & request("cust_postcode")	
			cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & cust_city_id	
			cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("cust_postcode")	
			cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("cust_postcode")	
			cust_state = selectid(sql)
	end if

	if  request.form("cust_cnty_id") <> "129" then 
			cust_city = request.form("cust_city")
	   	    cust_city_id = "0"
	end if

sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_code, cust_name, cust_type, cust_status, " & _
      "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
	  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic, cust_icno " & _
	  "FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof or request("cust_code") ="" then
		    rs.addnew
			if len(ChkString(Request("cust_postcode"))) = "5" or len(ChkString(Request("cust_postcode"))) = "6" then 	
				rs("cust_createddate") = ChkDateTimeMySQL(now())
				rs("cust_createdby") = Request.Cookies("GAPS")("sloginid")
				rs("cust_status") = request("cust_status")
				rs("cust_type") = request("cust_type")
				rs("cust_name") = request("cust_name")
				rs("cust_reg_no") = request("cust_reg_no")
				rs("cust_company") = request("cust_company")
				rs("cust_address") = request("cust_address") 
				rs("cust_postcode") = request("cust_postcode") 

				if  request.form("cust_cnty_id") = "129" then 
						rs("cust_state") = cust_state
						rs("cust_state_id") = cust_state_id  
				end if

				rs("cust_city")  = cust_city
				rs("cust_city_id") = cust_city_id  
				rs("cust_cnty_id")  = request("cust_cnty_id")
				rs("cust_email")  = request("cust_email")
				rs("cust_tel1")  = request("cust_tel1")
				rs("cust_tel2")  = request("cust_tel2")
				rs("cust_fax")   = request("cust_fax")
				rs("cust_website")  = request("cust_website")
				rs("cust_password") =  ""
				rs("cust_gstregno")  = request("cust_gstregno")
				rs("cust_source")  = request("cust_source")
				rs("cust_attention") = request("cust_attention")
				rs("cust_icno") = request("cust_icno")
				rs("cust_pic") = request("cust_pic")
				rs.Update 
			end if
		rs.Close 
		end if
		
		sql = "select top 1 cust_id from tblcustomer order by cust_id desc "
        cust_id = selectid(sql)
		temp = 10000 + cust_id
        cust_code = "C" & temp 
        sql = "update tblcustomer set cust_code = '" & cust_code & "' where cust_id = " & cust_id
        CUD(sql)

        url = "rm_customer_new.asp?cust_code=" & cust_code & "&loginerr=Customer has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcustomer','addCustomer=" & ChkString(left(request("cust_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 		 
'----------------------------------------------------------------------------------------------------    
  Case "editCustomer"   

	if request.form("cust_postcode") <> "" and request.form("cust_cnty_id") = "129" then 
			cust_postcode=request("job_cust_postcode")

			sql = "select city_id from tblpostcode where postcode =" & request("cust_postcode")	
			cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & cust_city_id	
			cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("cust_postcode")	
			cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("cust_postcode")	
			cust_state = selectid(sql)
	end if

	if  request.form("job_cust_cnty_id") <> "129" then 
			cust_city = request.form("cust_city")
	   	    cust_city_id = "0"
	end if

sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_code, cust_name, cust_type, cust_status, " & _
      "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
	  "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic, cust_icno " & _
	  "FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then
			 if len(request("cust_postcode")) = "5" or len(ChkString(Request("cust_postcode"))) = "6" then  
					rs("cust_status") = request("cust_status")
					rs("cust_type") = request("cust_type")
					rs("cust_name") = request("cust_name")
					rs("cust_reg_no") = request("cust_reg_no")
					rs("cust_company") = request("cust_company")
					rs("cust_address") = request("cust_address") 
					rs("cust_postcode") = request("cust_postcode") 
					
					if  request.form("cust_cnty_id") = "129" then 
							rs("cust_state") = cust_state  
							rs("cust_state_id") = request("cust_state_id")  
					end if
		
					rs("cust_city")  = cust_city
					rs("cust_city_id") = cust_city_id  
					rs("cust_cnty_id")  = request("cust_cnty_id")
					rs("cust_email")  = request("cust_email")
					rs("cust_tel1")  = request("cust_tel1")
					rs("cust_tel2")  = request("cust_tel2")
					rs("cust_fax")   = request("cust_fax")
					rs("cust_website")  = request("cust_website")
					rs("cust_password") =  ""
					rs("cust_gstregno")  = request("cust_gstregno")
					rs("cust_source")  = request("cust_source")
					rs("cust_attention") = request("cust_attention")
					rs("cust_icno") = request("cust_icno")
					rs("cust_pic") = request("cust_pic")
			end if
		rs.Update 
		rs.Close 
		end if

        url = "rm_customer_new.asp?cust_code=" & request("cust_code") & "&loginerr=Customer has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcustomer','editCustomer=" & ChkString(left(request("cust_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "addOnlineWrty"  
  
if request("cust_code") <> "" then	  
sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
      "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email, cust_tel1,  " & _
      "cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic, cust_icno " & _
	  "FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "' "
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			cust_id = rs("cust_id") 
			cust_createdby = rs("cust_createdby") 
			cust_createddate = rs("cust_createddate") 
			cust_code = rs("cust_code") 
			cust_name = rs("cust_name") 
			cust_type = rs("cust_type") 
			cust_status = rs("cust_status")
			cust_reg_no = rs("cust_reg_no")
			cust_company = rs("cust_company")
			cust_address = rs("cust_address")
			cust_postcode = rs("cust_postcode") 
			cust_state = rs("cust_state") 
			cust_state_id = rs("cust_state_id") 
			cust_city = rs("cust_city") 
			cust_city_id = rs("cust_city_id") 
			cust_cnty_id = rs("cust_cnty_id") 
			cust_email = rs("cust_email") 
			cust_tel1 = rs("cust_tel1") 
			cust_tel2 = rs("cust_tel2") 
			cust_fax = rs("cust_fax") 
			cust_website = rs("cust_website") 
			cust_password = rs("cust_password")
			cust_gstregno = rs("cust_gstregno")  
			cust_lastjob_code = rs("cust_lastjob_code")  
			cust_source = rs("cust_source") 
			cust_attention = rs("cust_attention") 
			cust_pic = rs("cust_pic")  
			cust_icno = rs("cust_icno")  
		End If
		rs.Close
end if
	 
sql = "SELECT refer_id, warrantyno, productmodel, othermodel, serialno, dealername, purchase_date, invoiceno, deliveryno, customername, " & _ 
	  "customericno, customeremail, customeraddress, customerpostcode, customerstate, customercity, customertel1, customertel2, customerfax, cust_code " & _
	  "FROM tblonlinewarranty WHERE refer_id = '" & request("refer_id") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		   rs.addnew
		   rs("warrantyno") = request("warrantyno")
		   rs("productmodel") = request("productmodel")
		   rs("othermodel") = request("othermodel")
		   rs("serialno") = request("serialno")
		   rs("dealername") = request("dealername")
		   
		   if request("purchase_date") <> "" then 
		   rs("purchase_date") = request("purchase_date")
		   end if
		   
		   rs("invoiceno") = request("invoiceno")
		   rs("deliveryno") = request("deliveryno")
		   rs("customername") = cust_name
		   rs("customericno") = cust_icno
		   rs("customeremail") = cust_email
		   rs("customeraddress") = cust_address
		   rs("customerpostcode") = cust_postcode
		   rs("customerstate") = cust_state
		   rs("customercity") = cust_city
		   rs("customertel1") = cust_tel1
		   rs("customertel2") = cust_tel2
		   rs("customerfax") = cust_fax
		   rs("cust_code") = cust_code
		   rs.Update 
		rs.Close 
		end if
		
	    url = "rm_customer_new.asp?cust_code=" & cust_code & "&loginerr=Warranty has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblonlinewarranty','addOnlineWrty=" & ChkString(left(request("dod_warrantynodo_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "editOnlineWrty"  
  
if request("cust_code") <> "" then	  
sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
      "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email, cust_tel1,  " & _
      "cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic, cust_icno " & _
	  "FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "' "
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			cust_id = rs("cust_id") 
			cust_createdby = rs("cust_createdby") 
			cust_createddate = rs("cust_createddate") 
			cust_code = rs("cust_code") 
			cust_name = rs("cust_name") 
			cust_type = rs("cust_type") 
			cust_status = rs("cust_status")
			cust_reg_no = rs("cust_reg_no")
			cust_company = rs("cust_company")
			cust_address = rs("cust_address")
			cust_postcode = rs("cust_postcode") 
			cust_state = rs("cust_state") 
			cust_state_id = rs("cust_state_id") 
			cust_city = rs("cust_city") 
			cust_city_id = rs("cust_city_id") 
			cust_cnty_id = rs("cust_cnty_id") 
			cust_email = rs("cust_email") 
			cust_tel1 = rs("cust_tel1") 
			cust_tel2 = rs("cust_tel2") 
			cust_fax = rs("cust_fax") 
			cust_website = rs("cust_website") 
			cust_password = rs("cust_password")
			cust_gstregno = rs("cust_gstregno")  
			cust_lastjob_code = rs("cust_lastjob_code")  
			cust_source = rs("cust_source") 
			cust_attention = rs("cust_attention") 
			cust_pic = rs("cust_pic")  
			cust_icno = rs("cust_icno")  
		End If
		rs.Close
end if
	 
sql = "SELECT refer_id, warrantyno, productmodel, othermodel, serialno, dealername, purchase_date, invoiceno, deliveryno, customername, " & _ 
	  "customericno, customeremail, customeraddress, customerpostcode, customerstate, customercity, customertel1, customertel2, customerfax, cust_code " & _
	  "FROM tblonlinewarranty WHERE refer_id = '" & request("refer_id") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then
		   rs("warrantyno") = request("warrantyno")
		   rs("productmodel") = request("productmodel")
		   rs("othermodel") = request("othermodel")
		   rs("serialno") = request("serialno")
		   rs("dealername") = request("dealername")
		   
		   if request("purchase_date") <> "" then 
		   rs("purchase_date") = request("purchase_date")
		   end if
		   
		   rs("invoiceno") = request("invoiceno")
		   rs("deliveryno") = request("deliveryno")
		   rs("customername") = cust_name
		   rs("customericno") = cust_icno
		   rs("customeremail") = cust_email
		   rs("customeraddress") = cust_address
		   rs("customerpostcode") = cust_postcode
		   rs("customerstate") = cust_state
		   rs("customercity") = cust_city
		   rs("customertel1") = cust_tel1
		   rs("customertel2") = cust_tel2
		   rs("customerfax") = cust_fax
		   rs("cust_code") = cust_code
		   rs.Update 
		rs.Close 
		end if
		
	    url = "rm_customer_new.asp?cust_code=" & cust_code & "&loginerr=Warranty has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblonlinewarranty','editOnlineWrty=" & ChkString(left(request("dod_warrantynodo_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "delOnlineWrty"
  
	sql = "delete from tblonlinewarranty where refer_id=" & request("refer_id")	
	CUD(sql)

	    url = "rm_customer_new.asp?cust_code=" & request("cust_code") & "&loginerr=Warranty has been deleted.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblonlinewarranty','delOnlineWrty=" & ChkString(left(request("dod_warrantynodo_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	 		 		 		 
'----------------------------------------------------------------------------------------------------    
 Case "addPOR"	
	'this logic handles both add/edit

	if request.form("orderqty1") <> "" and request.form("eta1") <> "" and request.form("partcode") <> "" then 'if user presses Save wthout entering at least 1 request then skip
		sql="SELECT por_id,por_docno,por_date,por_remark,por_part_code,por_eta1,por_order_qty1,por_eta2,por_order_qty2,por_last_stockin,por_total_incoming" & _
		",por_total_last,por_avg_3,por_avg_6,por_mth1,por_mth2,por_mth3,por_mth4,por_mth5,por_mth6,por_mth1_qty,por_mth2_qty,por_mth3_qty,por_mth4_qty" & _
		",por_mth5_qty,por_mth6_qty,por_createdby,por_createddate,por_sw1_qty,por_ex_qty FROM tblpor where por_id = '" & request("por_id") & "' and por_docno = '" & request("por_docno") & "' "
	    set rs = server.CreateObject("adodb.recordset")

	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		    rs.addnew
         end if
	        rs("por_docno") = request.form("por_docno")
			rs("por_date") = request.form("por_date")
			rs("por_remark") = request.form("por_remark")			
			rs("por_part_code") = request.form("partcode")			
			if ChkString(Request.Form("eta1")) <> "" then 
				rs("por_eta1") = ChkString(Request.Form("eta1"))
			end if
			rs("por_order_qty1") = ChkNumber2(Request.Form("orderqty1"))
			if ChkString(Request.Form("eta2")) <> "" then 
				rs("por_eta2") = ChkString(Request.Form("eta2"))
			end if		
			rs("por_order_qty2") = ChkNumber2(Request.Form("orderqty2"))
			rs("por_last_stockin")  = request.form("laststockin")
			rs("por_total_incoming")  = request.form("totalincomingstk")
			rs("por_total_last")  = request.form("stklastformonths")
			rs("por_avg_3")  = request.form("avgfor3mth")
			rs("por_avg_6")  = request.form("avgfor6mth")
			rs("por_mth1")   = request.form("mth1")
			rs("por_mth2")  = request.form("mth2")
			rs("por_mth3") =  request.form("mth3")
			rs("por_mth4")  = request.form("mth4")
			rs("por_mth5")  = request.form("mth5")
			rs("por_mth6") = request.form("mth6")
			rs("por_mth1_qty") = ChkNumber2(request.form("avgmth1"))
			rs("por_mth2_qty") = ChkNumber2(request.form("avgmth2"))
			rs("por_mth3_qty") = ChkNumber2(request.form("avgmth3"))
			rs("por_mth4_qty") = ChkNumber2(request.form("avgmth4"))
			rs("por_mth5_qty") = ChkNumber2(request.form("avgmth5"))
			rs("por_mth6_qty") = ChkNumber2(request.form("avgmth6"))
			rs("por_createdby") = Request.Cookies("GAPS")("sloginid")
			rs("por_createddate") =  ChkDateTimeMySQL(now())
			rs("por_sw1_qty") = ChkNumber2(request.form("sw1_qty"))
			rs("por_ex_qty") = ChkNumber2(request.form("ex_qty"))
		    rs.Update 
			rs.Close 		

		 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
         Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblpor','addPOR=" & ChkString(left(request("por_docno"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	end if

    url = "rm_por_edit.asp?por_docno=" & request("por_docno") & "&loginerr=POR Document has been created.#articletitle" 
  
'---------------------------------------------------------------------------------------------------
 Case "delPOR"
	  
	sql = "delete FROM tblpor where por_id = '" & request("por_id") & "' and por_docno = '" & request("por_docno") & "' "
	CUD(sql)

	url = "rm_por_edit.asp?por_docno=" & request("por_docno") & "&loginerr=POR Document has been created.#articletitle" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblpor','delPOR=" & ChkString(left(request("por_docno"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'----------------------------------------------------------------------------------------------------
 Case "addTechnClaim"
	if Request.Cookies("GAPS")("slevel") = "sc" then 'login by cs
			tech_code = request("tech_code")
	else
			tech_code = request.cookies("GAPS")("sloginid") 'login by tech
	end if
	
	set mySmartUpload = server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.Upload 	
	tempid1 = mySmartUpload.form("toll_receipt")
	tempid2 = mySmartUpload.form("hotel_receipt")
	tempid3 = mySmartUpload.form("parking_receipt")
	tempid4 = mySmartUpload.form("fuel_receipt")

	intCount1 = 1

	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	For each file In mySmartUpload.Files
	If not file.IsMissing Then  
		fullfilepath=fulldocpath_claims & file.FileName
		Select Case file.name
		Case "toll_receipt"  
				if fs.FileExists(fullfilepath) then
					realfilename = filenameChange(file.FileName,file.FileExt)
					file.SaveAs(Server.MapPath(documentpath_claims & tempid1 & filenameChange(file.FileName,file.FileExt)))			
					toll_receipt = tempid1 & realfilename
				else
					file.SaveAs(Server.MapPath(documentpath_claims & tempid1 & file.FileName))  		
					toll_receipt = tempid1 & file.FileName
				end if
	    Case "hotel_receipt" 
				if fs.FileExists(fullfilepath) then
					realfilename = filenameChange(file.FileName,file.FileExt)
					file.SaveAs(Server.MapPath(documentpath_claims & tempid2 & realfilename))			
					hotel_receipt = tempid2 & realfilename
				else
					file.SaveAs(Server.MapPath(documentpath_claims & tempid2 & file.FileName))  		
					hotel_receipt = tempid2 & file.FileName
				end if
		Case "parking_receipt" 
				if fs.FileExists(fullfilepath) then
					realfilename = filenameChange(file.FileName,file.FileExt)
					file.SaveAs(Server.MapPath(documentpath_claims & tempid3 & realfilename))			
					parking_receipt = tempid3 & realfilename
				else
					file.SaveAs(Server.MapPath(documentpath_claims & tempid3 & file.FileName))  		
					parking_receipt = tempid3 & file.FileName
				end if 
		Case "fuel_receipt" 
				if fs.FileExists(fullfilepath) then
					realfilename = filenameChange(file.FileName,file.FileExt)
					file.SaveAs(Server.MapPath(documentpath_claims & tempid4 & realfilename))			
					fuel_receipt = tempid4 & realfilename
				else
					file.SaveAs(Server.MapPath(documentpath_claims & tempid4 & file.FileName))  		
					fuel_receipt = tempid4 & file.FileName
				end if
		End Select
	 
	 end if   
  	 Next 
	 set fs=nothing
	
sql = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_overwrty_amt,tc_total_toll,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2,tc_toll_receipt,tc_hotel_receipt,tc_parking_receipt,tc_fuel_receipt,tc_year_process,tc_month_process " & _
 	 "FROM tbltech_claim where tc_tech_code = '" & tech_code & "' and tc_month ='" & ChkString(mySmartUpload.form("jobmonth")) & "' and tc_year ='" & ChkString(mySmartUpload.form("jobyear")) & "'"
	set rs = server.CreateObject("adodb.recordset")

	'if mySmartUpload.form("total_mileage") <> "" or mySmartUpload.form("total_toll") <> ""  or mySmartUpload.form("total_hotel") <> "" or mySmartUpload.form("descamt1") <> "" or mySmartUpload.form("descamt2") <> "" or mySmartUpload.form("deducdescamt1") <> "" or mySmartUpload.form("deducdescamt2") <> "" then
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
	    if rs.eof then
		    rs.addnew
	    end if
		rs("tc_year_process") = mySmartUpload.form("jobyear") '1st time use default submisson mth entered by technician
		rs("tc_month_process") = mySmartUpload.form("jobmonth")
		'end if
	    rs("tc_tech_code") = tech_code
		rs("tc_submit_date") = ChkString(mySmartUpload.form("submit_date"))
		rs("tc_year") = mySmartUpload.form("jobyear")
		rs("tc_month") = mySmartUpload.form("jobmonth")
		rs("tc_total_petrol") = ChkNumberInt(mySmartUpload.form("total_petrol"))
		rs("tc_total_extramileage") = ChkNumberInt(mySmartUpload.form("total_extramileage"))
		rs("tc_total_parking") = ChkNumber2(mySmartUpload.form("total_parking"))
		rs("tc_total_toll") = ChkNumber2(mySmartUpload.form("total_toll"))
		rs("tc_total_hotel") = ChkNumber2(mySmartUpload.form("total_hotel"))
		rs("tc_otherdesc1") = mySmartUpload.form("desc1")
		rs("tc_otheramt1") = ChkNumber2(mySmartUpload.form("descamt1"))
		rs("tc_otherdesc2") = mySmartUpload.form("desc2")
		rs("tc_otheramt2") = ChkNumber2(mySmartUpload.form("descamt2"))
		rs("tc_deduc1") = mySmartUpload.form("deducdesc1")
		rs("tc_deducamt1") = ChkNumber2(mySmartUpload.form("deducdescamt1"))
		rs("tc_deduc2") = mySmartUpload.form("deducdesc2")
		rs("tc_deducamt2") = ChkNumber2(mySmartUpload.form("deducdescamt2"))
		rs("tc_overwrty_amt") = ChkNumber2(mySmartUpload.form("overwrty_amt"))

		    if Request.Cookies("GAPS")("slevel") = "sc" then  'subsequent time CS can overwrite the value
				rs("tc_year_process") = mySmartUpload.form("jobyear2")
				rs("tc_month_process") = mySmartUpload.form("jobmonth2")
			end if

			if toll_receipt <> "" then  
				if not isnull(rs("tc_toll_receipt")) then
						rs("tc_toll_receipt") = rs("tc_toll_receipt") + "," + ChkString(toll_receipt)
				else
						rs("tc_toll_receipt") = ChkString(toll_receipt)
				end if
			end if
		
			if hotel_receipt <> "" then
				if not isnull(rs("tc_hotel_receipt")) then
						rs("tc_hotel_receipt") = rs("tc_hotel_receipt") + "," + ChkString(hotel_receipt)
				else
						rs("tc_hotel_receipt") = ChkString(hotel_receipt)
				end if
			end if

			if parking_receipt <> "" then  
				if not isnull(rs("tc_parking_receipt")) then
						rs("tc_parking_receipt") = rs("tc_parking_receipt") + "," + ChkString(parking_receipt)
				else
						rs("tc_parking_receipt") = ChkString(parking_receipt)
				end if
			end if

			if fuel_receipt <> "" then  
				if not isnull(rs("tc_fuel_receipt")) then
						rs("tc_fuel_receipt") = rs("tc_fuel_receipt") + "," + ChkString(fuel_receipt)
				else
						rs("tc_fuel_receipt") = ChkString(fuel_receipt)
				end if
			end if

		    rs.Update 	
		rs.Close 	
	'end if 	    

		url = "rmtech_claims.asp?tech_code=" & mySmartUpload.form("tech_code") &" &loginerr=Technician claim has been updated.#articletitle" 

		 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim','addClaims=" & tech_code & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------
 Case "addTechnClaim_back"

	set mySmartUpload = server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.TotalMaxFileSize = 8388608
	mySmartUpload.MaxFileSize = 2097152

	mySmartUpload.Upload 
	
	tempid = mySmartUpload.form("toll_receipt")
	
	intCount = 1
	For each file In mySmartUpload.Files
	 IF not file.IsMissing Then   
		file.SaveAs(Server.MapPath(documentpath & tempid & file.FileName))  		
			job_wrty_photo = tempid & file.FileName
	 End if   
	intCount = intCount + 1     
	Next 



sql = "SELECT tc_claimID,tc_tech_code,tc_submit_date,tc_year, tc_month,tc_total_petrol,tc_total_parking,tc_overwrty_amt,tc_total_toll,tc_total_hotel,tc_total_extramileage,tc_otherdesc1,tc_otheramt1,tc_otherdesc2,tc_otheramt2,tc_deduc1,tc_deducamt1,tc_deduc2,tc_deducamt2 " & _
 	 "FROM tbltech_claim where tc_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' and tc_month ='" & ChkString(Request.Form("jobmonth")) & "' and tc_year ='" & ChkString(Request.Form("jobyear")) & "'"
	set rs = server.CreateObject("adodb.recordset")

	if request("total_mileage") <> "" or request("total_toll") <> ""  or request("total_hotel") <> "" or request("descamt1") <> "" or request("descamt2") <> "" or request("deducdescamt1") <> "" or request("deducdescamt2") <> "" then
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
	    if rs.eof then
		    rs.addnew
	        rs("tc_tech_code") = request.cookies("GAPS")("sloginid")
			rs("tc_submit_date") = ChkString(request("submit_date"))
			rs("tc_year") = request("jobyear")
			rs("tc_month") = request("jobmonth")
			rs("tc_total_petrol") = ChkNumberInt(request("total_petrol"))
			rs("tc_total_extramileage") = ChkNumberInt(request("total_extramileage"))
			rs("tc_total_parking") = ChkNumber2(request("total_parking"))
			rs("tc_total_toll") = ChkNumber2(request("total_toll"))
			rs("tc_total_hotel") = ChkNumber2(request("total_hotel"))
			rs("tc_otherdesc1") = request("desc1")
			rs("tc_otheramt1") = ChkNumber2(request("descamt1"))
			rs("tc_otherdesc2") = request("desc2")
			rs("tc_otheramt2") = ChkNumber2(request("descamt2"))
			rs("tc_deduc1") = request("deducdesc1")
			rs("tc_deducamt1") = ChkNumber2(request("deducdescamt1"))
			rs("tc_deduc2") = request("deducdesc2")
			rs("tc_deducamt2") = ChkNumber2(request("deducdescamt2"))
			rs("tc_overwrty_amt") = ChkNumber2(request("overwrty_amt"))
		    rs.Update 
	    else
	        rs("tc_tech_code") = request.cookies("GAPS")("sloginid")
			rs("tc_submit_date") = ChkString(request("submit_date"))
			rs("tc_year") = request("jobyear")
			rs("tc_month") = request("jobmonth")
			rs("tc_total_petrol") = ChkNumberInt(request("total_petrol"))
			rs("tc_total_parking") = ChkNumber2(request("total_parking"))
			rs("tc_total_toll") = ChkNumber2(request("total_toll"))
			rs("tc_total_hotel") = ChkNumber2(request("total_hotel"))
			rs("tc_total_extramileage") = ChkNumberInt(request("total_extramileage"))
			rs("tc_otherdesc1") = request("desc1")
			rs("tc_otheramt1") = ChkNumber2(request("descamt1"))
			rs("tc_otherdesc2") = request("desc2")
			rs("tc_otheramt2") = ChkNumber2(request("descamt2"))
			rs("tc_deduc1") = request("deducdesc1")
			rs("tc_deducamt1") = ChkNumber2(request("deducdescamt1"))
			rs("tc_deduc2") = request("deducdesc2")
			rs("tc_deducamt2") = ChkNumber2(request("deducdescamt2"))
			rs("tc_overwrty_amt") = ChkNumber2(request("overwrty_amt"))
			rs.Update 
		end if
		rs.Close 	
	end if 	    

		url = "rmtech_claims.asp?tech_code=" & request("tech_code") &" &loginerr=Technician has been updated.#articletitle" 

		 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltech_claim','addClaims=" & request.Cookies("GAPS")("job_tech_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
'---------------------------------------------------------------------------------------------------------------------------------------------

  Case "addTechnician"   

	if request.form("tech_postcode") <> ""  then
	  if request.form("tech_city_id") = "" and request.form("tech_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			tech_id = request("tech_id")
			tech_postcode=request("tech_postcode")
			tech_name=request("tech_name")
			tech_tel1=request("tech_tel1")
			tech_tel2=request("tech_tel2")
			tech_icno=request("tech_icno")
			tech_address=request("tech_address")
			tech_code = request("tech_code")

			sql = "select city_id from tblpostcode where postcode =" & request("tech_postcode")	
			tech_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & tech_city_id	
			tech_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("tech_postcode")	
			tech_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("tech_postcode")	
			tech_state = selectid(sql)
	  	    Response.Redirect "rm_contractor_new.asp?tech_code="&tech_code&"&tech_postcode="&tech_postcode&"&tech_name="&tech_name&"&tech_address="&tech_address&"&tech_tel1="&tech_tel1&"&tech_tel2="&tech_tel2&"&tech_icno="&tech_icno&"&tech_id="&tech_id& "&loginerr=Updated Address.#articletitle" 			
		end if
	else
			Response.Redirect "rm_contractor_new.asp?tech_code="&tech_code&"&tech_postcode="&tech_postcode&"&tech_name="&tech_name&"&tech_address="&tech_address&"&tech_tel1="&tech_tel1&"&tech_tel2="&tech_tel2&"&tech_icno="&tech_icno&"&tech_id="&tech_id& "&loginerr=Updated Address.#articletitle" 			
	end if

	   sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id, tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id, tech_wh_code, " & _
	  "tech_creditlimit, tech_terms,tech_salary " & _
	  "FROM tbltechnician WHERE tech_code = '" & request("tech_code") & "' "	

	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		    rs.addnew
		end if
	        rs("tech_code") = request("tech_code")
			rs("tech_type") = request("tech_type")
			rs("tech_name") = request("tech_name")
			rs("tech_icno") = request("tech_icno")
			rs("tech_address") = request("tech_address")
			rs("tech_postcode") = request("tech_postcode") 
			rs("tech_state") = request("tech_state")  
			rs("tech_state_id") = request("tech_state_id")  
			rs("tech_city")  = tech_city
			rs("tech_city_id")  = tech_city_id
			rs("tech_email")  = request("tech_email")
			rs("tech_tel1")  = request("tech_tel1")
			rs("tech_tel2")  = request("tech_tel2")
			rs("tech_carmodel")   = request("tech_carmodel")
			rs("tech_carplateno")  = request("tech_carplateno")
			rs("tech_carcolour") =  request("tech_carcolour")
			rs("tech_password")  = request("tech_password")
			rs("tech_status")  = request("tech_status")
			rs("tech_area") = tech_area
			rs("tech_area_id") = request("tech_area_id")
			rs("tech_wh_code") = request("tech_wh_code")

			if request("tech_salary") = "" then 
				rs("tech_salary") = 0 
			else
				rs("tech_salary") = request("tech_salary")	
			end if

			if request("tech_creditlimit") <> "" then
				rs("tech_creditlimit") = ChkNumber(request("tech_creditlimit"))
			end if
	
			if request("tech_terms") <> "" then
				rs("tech_terms") = ChkNumber(request("tech_terms"))
			end if
		    rs.Update 
		rs.Close 
		

        url = "rm_contractor_new.asp?tech_code=" & request("tech_code") & "&loginerr=Technician has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltechnician','UpdateTechnician=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	
'----------------------------------------------------------------------------------------------------    
  Case "editTechnician"  ' procedure not used. update/add done in add proc	
	 if request.form("tech_postcode") <> ""  then
			tech_postcode=request.form("tech_postcode")	

			sql = "select city_id from tblpostcode where postcode =" & tech_postcode	
			tech_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & tech_city_id	
			tech_city = selectid(sql)	

			sql = "select state_id from tblpostcode where postcode =" & tech_postcode
			tech_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & tech_postcode 
			tech_state = selectid(sql)	
	 else
			Response.Redirect "rm_contractor_new.asp?tech_code="&request.form("tech_code")&"&loginerr=Tech Not Updated.#articletitle"
	 end if
		  
sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id, tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id, tech_wh_code, " & _
	  "tech_creditlimit, tech_terms " & _
	  "FROM tbltechnician WHERE tech_code = '" & request("tech_code") & "' "	

	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then
	        rs("tech_code") = request("tech_code")
			rs("tech_type") = request("tech_type")
			rs("tech_name") = request("tech_name")
			rs("tech_icno") = request("tech_icno")
			rs("tech_address") = request("tech_address")
			rs("tech_postcode") = request("tech_postcode") 
			rs("tech_state") = tech_state  
			rs("tech_state_id") = tech_state_id  
			rs("tech_city")  = tech_city
			rs("tech_city_id")  = tech_city_id
			rs("tech_email")  = request("tech_email")
			rs("tech_tel1")  = request("tech_tel1")
			rs("tech_tel2")  = request("tech_tel2")
			rs("tech_carmodel")   = request("tech_carmodel")
			rs("tech_carplateno")  = request("tech_carplateno")
			rs("tech_carcolour") =  request("tech_carcolour")
			rs("tech_password")  = request("tech_password")
			rs("tech_status")  = request("tech_status")
			rs("tech_area") = tech_area
			rs("tech_area_id") = request("tech_area_id")
			rs("tech_wh_code") = request("tech_wh_code")
			rs("tech_creditlimit") = request("tech_creditlimit")
			rs("tech_terms") = request("tech_terms")

			if ChkString(request("tech_salary")) = "" then 
				rs("tech_salary") = 0 
			else
				rs("tech_salary") = ChkString(request("tech_salary"))	
			end if

		rs.Update 
		rs.Close 
		end if
		
		'''Update tblwarehouse
		sql = "update tblwarehouse set wh_contact_person='" & request("tech_code") & "' where wh_code='" & request("tech_wh_code") & "'"
		CUD(sql)


        url = "rm_contractor_new.asp?tech_code=" & request("tech_code") & "&loginerr=Technician has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbltechnician','editTechnician=" & ChkString(left(request("tech_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)		
		 	 	 	
'----------------------------------------------------------------------------------------------------     
 
    Case "AddSparepartsRequest"   
	
        sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
            "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
	        "FROM tbltechnician WHERE tech_code = '" & request("sp_tech_code") & "' "
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_tech_code = rs("tech_code")
			sp_tech_name = rs("tech_name")
			sp_tech_icno = rs("tech_icno")
			sp_tech_address = rs("tech_address")
			sp_tech_postcode = rs("tech_postcode") 
			sp_tech_state = rs("tech_state") 
			sp_tech_state_id = rs("tech_state_id") 
			sp_tech_city = rs("tech_city") 
			sp_tech_city_id = rs("tech_city_id") 
			sp_tech_email = rs("tech_email") 
			sp_tech_tel1 = rs("tech_tel1") 
			sp_tech_tel2 = rs("tech_tel2") 
			sp_tech_carmodel = rs("tech_carmodel") 
			sp_tech_carplateno = rs("tech_carplateno") 
			sp_tech_carcolour = rs("tech_carcolour") 
		End If
		rs.Close
	
        ''''Add Spareparts Request   	  
        sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, sp_deliverydate, sp_deliveryby, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest "		
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("sp_tech_code")  = ChkString(sp_tech_code)	
        rs("sp_tech_name")  = ChkString(sp_tech_name)
		rs("sp_tech_address") = ChkString(sp_tech_address)
		rs("sp_tech_postcode") = ChkString(sp_tech_postcode)
		rs("sp_tech_state") = sp_tech_state
		rs("sp_tech_city") = sp_tech_city
		rs("sp_tech_email") = ChkString(sp_tech_email) 
		rs("sp_tech_tel1") = ChkString(sp_tech_tel1) 
		rs("sp_tech_tel2") = ChkString(sp_tech_tel2) 
		rs("sp_tech_carplateno") = ChkString(sp_tech_carplateno) 
		rs("sp_createddate") = ChkDateTimeMySQL(now())
		rs("sp_createdby") = Request.Cookies("GAPS")("sloginid")
		rs("sp_date") =  ChkDateTimeMySQL(now())
		rs("sp_status") = "Open"
		
		if ChkString(Request.Form("sp_deliverydate"))  <> "" then  
		rs("sp_deliverydate") = ChkString(Request.Form("sp_deliverydate")) 
		end if
		
		if ChkString(Request.Form("sp_confirmedreceiveddate"))  <> "" then  
		rs("sp_confirmedreceiveddate") = ChkString(Request.Form("sp_confirmedreceiveddate")) 
		end if
		
		rs("sp_remark") = ChkString(Request.Form("sp_remark")) 
		rs.Update 
		rs.Close      
		
        sql = "select top 1 sp_id from tblsparepartrequest order by sp_id desc "
        sp_id = selectid(sql)
		temp = 100000 + sp_id
        sp_no = "SP" & temp 
  
        sql = "update tblsparepartrequest set sp_no = '" & sp_no & "' where sp_id = " & sp_id
        CUD(sql)
      
        url = "rm_consignment_new.asp?sp_no=" & sp_no & "&loginerr=New Spare Parts Request has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','AddSparepartsRequest=" & ChkString(left(sp_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)


'----------------------------------------------------------------------------------------------------     
 
    Case "editSparepartsRequest"   

        sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
            "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
	        "FROM tbltechnician WHERE tech_code = '" & request("sp_tech_code") & "' "
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_tech_code = rs("tech_code")
			sp_tech_name = rs("tech_name")
			sp_tech_icno = rs("tech_icno")
			sp_tech_address = rs("tech_address")
			sp_tech_postcode = rs("tech_postcode") 
			sp_tech_state = rs("tech_state") 
			sp_tech_state_id = rs("tech_state_id") 
			sp_tech_city = rs("tech_city") 
			sp_tech_city_id = rs("tech_city_id") 
			sp_tech_email = rs("tech_email") 
			sp_tech_tel1 = rs("tech_tel1") 
			sp_tech_tel2 = rs("tech_tel2") 
			sp_tech_carmodel = rs("tech_carmodel") 
			sp_tech_carplateno = rs("tech_carplateno") 
			sp_tech_carcolour = rs("tech_carcolour") 
		End If
		rs.Close
	
        ''''Add Spareparts Request   	
		sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "				
	    rs.Open sql,strconnect,2,2,&H0001
        if not rs.eof then 
        rs("sp_tech_code")  = ChkString(sp_tech_code)	
        rs("sp_tech_name")  = ChkString(sp_tech_name)
		rs("sp_tech_address") = ChkString(sp_tech_address)
		rs("sp_tech_postcode") = ChkString(sp_tech_postcode)
		rs("sp_tech_state") = sp_tech_state
		rs("sp_tech_city") = sp_tech_city
		rs("sp_tech_email") = ChkString(sp_tech_email) 
		rs("sp_tech_tel1") = ChkString(sp_tech_tel1) 
		rs("sp_tech_tel2") = ChkString(sp_tech_tel2) 
		rs("sp_tech_carplateno") = ChkString(sp_tech_carplateno) 
		rs("sp_trackingno") = ChkString(Request.Form("sp_trackingno")) 
		rs("sp_couriercompany") = ChkString(Request.Form("sp_couriercompany")) 
		
		if ChkString(Request.Form("sp_deliverydate"))  <> "" then  
		rs("sp_deliverydate") = ChkString(Request.Form("sp_deliverydate")) 
		end if
		
		if ChkString(Request.Form("sp_confirmedreceiveddate"))  <> "" then  
		rs("sp_confirmedreceiveddate") = ChkString(Request.Form("sp_confirmedreceiveddate")) 
		end if
		
		rs("sp_remark") = ChkString(Request.Form("sp_remark")) 
		rs.Update 
		end if 
		rs.Close
		
        url = "rm_consignment_new.asp?sp_no=" & request("sp_no") & "&loginerr=New Spare Parts Request has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','AddSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
'----------------------------------------------------------------------------------------------------

   Case "SubmitSparepartsRequest"

   	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Submitted"
			rs("sp_submittedby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_submitteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Submitted&loginerr=Spare Parts Request has been Submitted.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','SubmitSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------

   Case "ApproveSparepartsRequest"

   	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Approved"
			rs("sp_approvedby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_approveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Approved&loginerr=Spare Parts Request has been Approved.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','ApproveSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------

   Case "RejectSparepartsRequest"

  	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedby, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Rejected"
			rs("sp_rejectedby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_rejecteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Rejected&loginerr=Spare Parts Request has been Rejected.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','RejectSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		    
'----------------------------------------------------------------------------------------------------

   Case "DeliveredSparepartsRequest"

  	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedby, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Delivered"
			rs("sp_deliveryby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_deliverydate") = ChkDateTimeMySQL(now())
			rs("sp_confirmedreceivedby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_confirmedreceiveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Delivered&loginerr=Spare Parts Request has been Delivered.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','DeliveredSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------

   Case "PostedSparepartsRequest"
 
     	sql1 = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedby, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate, sp_posteddate, sp_postedby " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			rs1("sp_status") = "Posted"
			rs1("sp_postedby") = Request.Cookies("GAPS")("sloginid")
			rs1("sp_posteddate") = ChkDateTimeMySQL(now())
			sp_tech_code = rs1("sp_tech_code")
		rs1.Update 
		rs1.Close 
		end if
		
		
		'sql = "select wh_code from tblwarehouse where wh_contact_person='" & sp_tech_code & "'"
		sql="select tech_wh_code from tbltechnician where tech_code='" & sp_tech_code & "'"
		wh_code = selectid(sql)
		
		'''''Stock-In Detail
		sql1 = "SELECT spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	                   "FROM tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "' order by spd_id"	  
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & wh_code & "' and wst_itm_code = '" & rs1("spd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = wh_code
				rs2("wst_itm_code") = ChkString(rs1("spd_partcode"))
				wst_itm_current_qty = ChkString(rs1("spd_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("spd_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedSparepartsRequest+Add=" & request("sp_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)	
			else
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("spd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("spd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedSparepartsRequest+Update=" & request("sp_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Update 
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("sp_no")
			rs2("stk_reference") = wh_code
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Spareparts-Request"
			rs2("stk_itm_code")  = ChkString(rs1("spd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = wh_code
			rs2("stk_desc")  = ChkString(rs1("spd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("spd_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("spd_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close    
			
			''''Add Stock In Detail - W1	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("spd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then 
				wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("spd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("spd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			end if
			rs2.Close   
			
			'Update Stocktrans - W1
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("sp_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Spareparts-Request"
			rs2("stk_itm_code")  = ChkString(rs1("spd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = wh_code
			rs2("stk_desc")  = ChkString(rs1("spd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("spd_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("spd_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close   

			sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
			Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedSparepartsRequest+DeductW1=" & request("sp_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
			CUD(sql)

		rs1.movenext
		wend
		rs1.close		

	    url = "rm_consignment_view.asp?sp_status=Posted&loginerr=Spare Parts Request has been Posted.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','PostedSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------
   Case "RevertSparepartsRequest"
 
   	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedby, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate, sp_posteddate, sp_postedby, sp_canceldate, sp_cancelby " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Open"
			rs("sp_logby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_logdate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Open&loginerr=Spare Parts Request has been Reverted.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','RevertSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------

   Case "CancelSparepartsRequest"
 
   	sql = "SELECT top 1 sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
				"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
				"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
				"sp_rejecteddate, sp_rejectedby, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
				"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate, sp_posteddate, sp_postedby, sp_canceldate, sp_cancelby " & _
				"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("sp_status") = "Cancel"
			rs("sp_cancelby") = Request.Cookies("GAPS")("sloginid")
			rs("sp_canceldate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

	    url = "rm_consignment_view.asp?sp_status=Cancel&loginerr=Spare Parts Request has been Cancel.#articletitle" 
	  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest','CancelSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 		            
'----------------------------------------------------------------------------------------------------     
 
    Case "addSparepartsRequestDetail"  

        ''''Add Spare Parts Detail	   	  
        sql = "SELECT top 1 spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	          "FROM tblsparepartrequest_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("spd_sp_no") = ChkString(Request.Form("sp_no"))
        rs("spd_tech_code")  = ChkString(Request.Form("spd_tech_code"))	
		rs("spd_partcode")  = ChkString(Request.Form("spd_partcode"))	
		rs("spd_currentstock")  = ChkString(Request.Form("spd_currentstock"))	
		rs("spd_desc")  = ChkString(Request.Form("spd_desc"))	
        rs("spd_unitcost")  = ChkString(Request.Form("spd_unitcost"))
		rs("spd_qty")  = ChkString(Request.Form("spd_qty"))
		rs("spd_subtotal")  = ChkString(Request.Form("spd_subtotal"))
		rs.Update 
		rs.Close      
		
		sql = "select sum(spd_qty) as spd_qty from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
        spd_qty = selectid(sql)
		
        sql = "select sum(spd_subtotal) as spd_subtotal from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
        spd_subtotal = selectid(sql)
		
        sql = "update tblsparepartrequest set sp_totalAmt = " & ChkNumber(spd_subtotal) & ", sp_totalqty=" & chknumber0(spd_qty) & " where sp_no = '" & request("sp_no") & "'"
        CUD(sql)
		
        url = "rm_consignment_new.asp?sp_no=" & request("sp_no") & "&loginerr=Spare Parts Request has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest_detail','AddSparepartsRequest=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)


'----------------------------------------------------------------------------------------------------     
 
    Case "editSparepartsRequestDetail"  

	   sql = "delete from tblsparepartrequest_detail where spd_id=" & request("spd_id")	
	   CUD(sql)
	
         ''''Add Spare Parts Detail	   	  
        sql = "SELECT top 1 spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	          "FROM tblsparepartrequest_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("spd_sp_no") = ChkString(Request.Form("sp_no"))
        rs("spd_tech_code")  = ChkString(Request.Form("spd_tech_code"))	
		rs("spd_partcode")  = ChkString(Request.Form("spd_partcode"))	
		rs("spd_currentstock")  = ChkString(Request.Form("spd_currentstock"))	
		rs("spd_desc")  = ChkString(Request.Form("spd_desc"))	
        rs("spd_unitcost")  = ChkString(Request.Form("spd_unitcost"))
		rs("spd_qty")  = ChkString(Request.Form("spd_qty"))
		rs("spd_subtotal")  = ChkString(Request.Form("spd_subtotal"))
		rs.Update 
		rs.Close  
		
		sql = "select sum(spd_qty) as spd_qty from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
        spd_qty = selectid(sql)
		
        sql = "select sum(spd_subtotal) as spd_subtotal from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
        spd_subtotal = selectid(sql)
		
        sql = "update tblsparepartrequest set sp_totalAmt = " & ChkNumber(spd_subtotal) & ", sp_totalqty=" & chknumber0(spd_qty) & " where sp_no = '" & request("sp_no") & "'"
        CUD(sql)
		
        url = "rm_consignment_new.asp?sp_no=" & request("sp_no") & "&loginerr=Spare Parts Request has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest_detail','editSparepartsRequestDetail=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		 
'----------------------------------------------------------------------------------------------------    
  Case "DelSparepartsRequestDetail"
  
	sql = "delete from tblsparepartrequest_detail where spd_id=" & request("spd_id")	
	CUD(sql)
	
	sql = "select sum(spd_qty) as spd_qty from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
	spd_qty = selectid(sql)
	
	if isnull(spd_qty) then 
	   spd_qty = 0
	end if
	
	sql = "select sum(spd_subtotal) as spd_subtotal from tblsparepartrequest_detail where spd_sp_no = '" & request("sp_no") & "'"
	spd_subtotal = selectid(sql)
	
	if isnull(spd_subtotal) then 
	   spd_subtotal = 0
	end if
	
	sql = "update tblsparepartrequest set sp_totalAmt = " & ChkNumber(spd_subtotal) & ", sp_totalqty=" & chknumber0(spd_qty) & " where sp_no = '" & request("sp_no") & "'"
	CUD(sql)
	
	url = "rm_consignment_new.asp?sp_no=" & request("sp_no") & "&loginerr=Spare Parts Request has been updated.#spareparts" 
	
	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblsparepartrequest_detail','DelSparepartsRequestDetail=" & ChkString(left(request("sp_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)


'----------------------------------------------------------------------------------------------------     
    Case "addRCN"   

if request.form("rcn_cust_cnty_id") <> "" and request.form("rcn_cust_postcode") = "" then 'just return regardless country
		rcn_cust_cnty_id=request.form("rcn_cust_cnty_id")
		rcn_cust_postcode=request("rcn_cust_postcode")
		rcn_cust_name=request("rcn_cust_name")
		rcn_cust_tel1=request("rcn_cust_tel1")
		rcn_cust_tel2=request("rcn_cust_tel2")
		rcn_cust_address=request("rcn_cust_address")
		Response.Redirect "rm_rcn_new.asp?rcn_no="&rcn_code&"&rcn_cust_postcode="&rcn_cust_postcode&"&rcn_cust_name="&rcn_cust_name&"&rcn_cust_address="&rcn_cust_address&"&rcn_cust_tel1="&rcn_cust_tel1&"&rcn_cust_tel2="&rcn_cust_tel2&"&rcn_cust_cnty_id="&rcn_cust_cnty_id&"&loginerr=Updated Address.#articletitle" 
end if

	if request.form("rcn_cust_postcode") <> ""  and request.form("rcn_cust_code") = ""  and request.form("job_cust_cnty_id") = "129" then
	  if request.form("rcn_cust_city_id") = "" and request.form("rcn_cust_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			rcn_cust_postcode=request("rcn_cust_postcode")
			rcn_cust_name=request("rcn_cust_name")
			rcn_cust_tel1=request("rcn_cust_tel1")
			rcn_cust_tel2=request("rcn_cust_tel2")
			rcn_cust_address=request("rcn_cust_address")

			sql = "select city_id from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & rcn_cust_city_id	
			rcn_cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_state = selectid(sql)

		    Response.Redirect "rm_rcn_new.asp?rcn_no="&rcn_code&"&rcn_cust_postcode="&rcn_cust_postcode&"&rcn_cust_name="&rcn_cust_name&"&rcn_cust_address="&rcn_cust_address&"&rcn_cust_tel1="&rcn_cust_tel1&"&rcn_cust_tel2="&rcn_cust_tel2& "&loginerr=Updated Address.#articletitle" 
		end if
	end if

	if request.form("rcn_cust_cnty_id") = "129" then
		sql = "select state_code from tblstate where state_id =" & request("rcn_cust_state_id") 
		state_code = selectid(sql)
	    
		if request.form("rcn_cust_code") <> "" then
			sql ="select cust_city_id from tblcustomer where cust_code ='" & request.form("rcn_cust_code") & "'"
			rcn_cust_city_id = selectid(sql)
		end if

		sql = "select ct_name2 from tblcity where ct_id ='" &  request.form("rcn_cust_city_id") & "'" 
		if request.form("rcn_cust_city_id") <> "" then
			rcn_cust_city = selectid(sql)	
			rcn_cust_city_id = request.form("rcn_cust_city_id")
		end if
	end if

if  request.form("rcn_cust_cnty_id") <> "129" then 
			rcn_cust_city = request.form("rcn_cust_city")
	   	    rcn_cust_city_id = "0"
end if

        ''''Add Job Order	   	  
        sql="SELECT top 1 rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_cnty_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,  " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn"	
		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("rcn_status")  = "Open"   					
        rs("rcn_date") = ChkDateYYYYMMDD(date())
        rs("rcn_cust_code")  = ChkString(Request.Form("rcn_cust_code"))	
        rs("rcn_cust_name")  = ChkString(Request.Form("rcn_cust_name"))
		rs("rcn_cust_address") = ChkString(Request.Form("rcn_cust_address"))
		rs("rcn_cust_postcode") = ChkString(Request.Form("rcn_cust_postcode"))
		
		if request.form("rcn_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs("rcn_cust_state") = rcn_cust_state
				rs("rcn_cust_state_id") = rcn_cust_state_id 
		end if 

		rs("rcn_cust_city") = rcn_cust_city
		rs("rcn_cust_city_id") = rcn_cust_city_id 
		rs("rcn_cust_cnty_id") = ChkString(Request.Form("rcn_cust_cnty_id")) 
		rs("rcn_cust_email") = ChkString(Request.Form("rcn_cust_email")) 
		rs("rcn_cust_tel1") = ChkString(Request.Form("rcn_cust_tel1")) 
		rs("rcn_cust_tel2") = ChkString(Request.Form("rcn_cust_tel2")) 
		rs("rcn_remark") = ChkString(Request.Form("rcn_remark")) 
		rs("rcn_createddate") = ChkDateTimeMySQL(now())
		rs("rcn_createdby") = Request.Cookies("GAPS")("sloginid")
		
		rs("rcn_modeltype") = ChkString(Request.Form("rcn_modeltype")) 
		
		rs.Update 
		rs.Close      
		
        sql = "select top 1 rcn_id from tblrcn order by rcn_id desc "
        rcn_id = selectid(sql)
		temp = 100000 + rcn_id
        rcn_no = "RCN" & temp 
  
        sql = "update tblrcn set rcn_no = '" & rcn_no & "' where rcn_id = " & rcn_id
        CUD(sql)
      
        url = "rm_rcn_new.asp?rcn_no=" & rcn_no & "&loginerr=New RCN has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn','addRCN=" & ChkString(left(rcn_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editRCN"   
	
	if request.form("rcn_cust_postcode") <> "" then
	 		rcn_cust_postcode=request("rcn_cust_postcode")

			sql = "select city_id from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & rcn_cust_city_id	
			rcn_cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("rcn_cust_postcode")	
			rcn_cust_state = selectid(sql)
	else
	   response.redirect("rm_rcn_new.asp?rcn_no=" & request("rcn_no") & "&loginerr=RCN not updated.#articletitle") 
	end if
	
        ''''Edit Job Order	   	  
        sql="SELECT top 1 rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_cnty_id,rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,  " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_no = '" & request("rcn_no") & "' "	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("rcn_cust_code")  = ChkString(Request.Form("rcn_cust_code"))	
			rs("rcn_cust_name")  = ChkString(Request.Form("rcn_cust_name"))
			rs("rcn_cust_address") = ChkString(Request.Form("rcn_cust_address"))
			rs("rcn_cust_postcode") = ChkString(Request.Form("rcn_cust_postcode"))
		
			if request.form("rcn_cust_cnty_id") = "129" then 'state applies to Malaysia only
				rs("rcn_cust_state") = rcn_cust_state
				rs("rcn_cust_state_id") = rcn_cust_state_id 
			end if
			
			rs("rcn_cust_city") = rcn_cust_city
			rs("rcn_cust_city_id") = rcn_cust_city_id 
			rs("rcn_cust_cnty_id") = ChkString(Request.Form("rcn_cust_cnty_id"))
			rs("rcn_cust_email") = ChkString(Request.Form("rcn_cust_email")) 
			rs("rcn_cust_tel1") = ChkString(Request.Form("rcn_cust_tel1")) 
			rs("rcn_cust_tel2") = ChkString(Request.Form("rcn_cust_tel2")) 
			rs("rcn_remark") = ChkString(Request.Form("rcn_remark")) 
			rs("rcn_modeltype") = ChkString(Request.Form("rcn_modeltype")) 
		rs.Update 
		rs.Close 
		end if

        url = "rm_rcn_new.asp?rcn_no=" & request("rcn_no") & "&loginerr=RCN has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn','editRCN=" & ChkString(left(request("rcn_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	 
		 
'----------------------------------------------------------------------------------------------------    
  Case "addRNCDetail"   
	
        ''''Add RCN parts	   	  
        sql = "SELECT top 1 rcnd_id, rcnd_rcn_no, rcnd_job_code, rcnd_partcode, rcnd_desc, rcnd_unitcost, rcnd_qty, rcnd_discountamt, " & _
		      "rcnd_discounttype, rcnd_netcost, rcnd_subtotal FROM tblrcn_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("rcnd_rcn_no") = ChkString(Request.Form("rcn_no"))
        rs("rcnd_job_code")  = ChkString(Request.Form("rcnd_job_code"))	
		rs("rcnd_partcode")  = ChkString(Request.Form("rcnd_partcode"))	
		rs("rcnd_desc")  = ChkString(Request.Form("rcnd_desc"))	
        rs("rcnd_unitcost")  = ChkString(Request.Form("rcnd_unitcost"))
		rs("rcnd_qty") = ChkString(Request.Form("rcnd_qty"))
		rs("rcnd_discountamt")  = ChkString(Request.Form("rcnd_discountamt"))
		rs("rcnd_discounttype")  = ChkString(Request.Form("rcnd_discounttype"))
		
		if ChkString(Request.Form("rcnd_discounttype")) = "%" then 
		rcnd_netcost  = ChkString(Request.Form("rcnd_unitcost")) * (ChkString(Request.Form("rcnd_discountamt")/100))
		rcnd_netcost = ChkString(Request.Form("rcnd_unitcost")) - rcnd_netcost
		rs("rcnd_netcost")  = rcnd_netcost
		rs("rcnd_subtotal")  = rcnd_netcost * ChkString(Request.Form("rcnd_qty"))
		else
		rcnd_netcost  = ChkString(Request.Form("rcnd_unitcost")) -  ChkString(Request.Form("rcnd_discountamt"))
		rs("rcnd_netcost")  = rcnd_netcost
		rs("rcnd_subtotal")  = rcnd_netcost * ChkString(Request.Form("rcnd_qty"))
		end if
		rs.Update 
		rs.Close      
		
        sql = "select sum(rcnd_subtotal) as rcnd_subtotal from tblrcn_detail where rcnd_rcn_no = '" & request("rcn_no") & "'"
        rcnd_subtotal = selectid(sql)
        rcn_gstAmt = rcnd_subtotal * GSTRateBack
		
		sql = "update tblrcn set rcn_totalPartsAmt=" & rcnd_subtotal & ", rcn_gstAmt = " & chknumber0(rcn_gstAmt) & ", rcn_totalAmt=" & rcnd_subtotal & " where rcn_no = '" & ChkString(Request.Form("rcn_no")) & "'"
        CUD(sql)
		
        url = "rm_rcn_new.asp?rcn_no=" & request("rcn_no") & "&loginerr=RCN has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn_detail','addRNCDetail=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------    
  Case "editRCNDetail"   
	
        ''''Add RCN parts	   	  
        sql = "SELECT rcnd_id, rcnd_rcn_no, rcnd_job_code, rcnd_partcode, rcnd_desc, rcnd_unitcost, rcnd_qty, rcnd_discountamt, " & _
		      "rcnd_discounttype, rcnd_netcost, rcnd_subtotal FROM tblrcn_detail where rcnd_id=" & request("rcnd_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("rcnd_partcode")  = ChkString(Request.Form("rcnd_partcode"))	
		rs("rcnd_desc")  = ChkString(Request.Form("rcnd_desc"))	
        rs("rcnd_unitcost")  = ChkString(Request.Form("rcnd_unitcost"))
		rs("rcnd_qty") = ChkString(Request.Form("rcnd_qty"))
		rs("rcnd_discountamt")  = ChkString(Request.Form("rcnd_discountamt"))
		rs("rcnd_discounttype")  = ChkString(Request.Form("rcnd_discounttype"))
		
		if ChkString(Request.Form("rcnd_discounttype")) = "%" then 
		rcnd_netcost  = ChkString(Request.Form("rcnd_unitcost")) * (ChkString(Request.Form("rcnd_discountamt")/100))
		rcnd_netcost = ChkString(Request.Form("rcnd_unitcost")) - rcnd_netcost
		rs("rcnd_netcost")  = rcnd_netcost
		rs("rcnd_subtotal")  = rcnd_netcost * ChkString(Request.Form("rcnd_qty"))
		else
		rcnd_netcost  = ChkString(Request.Form("rcnd_unitcost")) -  ChkString(Request.Form("rcnd_discountamt"))
		rs("rcnd_netcost")  = rcnd_netcost
		rs("rcnd_subtotal")  = rcnd_netcost * ChkString(Request.Form("rcnd_qty"))
		end if
		rs.Update 
		rs.Close  
		end if    
		
        sql = "select sum(rcnd_subtotal) as rcnd_subtotal from tblrcn_detail where rcnd_rcn_no = '" & request("rcn_no") & "'"
        rcnd_subtotal = selectid(sql)
        rcn_gstAmt = rcnd_subtotal * GSTRateBack
		
		sql = "update tblrcn set rcn_totalPartsAmt=" & rcnd_subtotal & ", rcn_gstAmt = " & chknumber0(rcn_gstAmt) & ", rcn_totalAmt=" & rcnd_subtotal & " where rcn_no = '" & ChkString(Request.Form("rcn_no")) & "'"
        CUD(sql)
		
        url = "rm_rcn_new.asp?rcn_no=" & request("rcn_no") & "&loginerr=RCN has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn_detail','editRCNDetail=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "delRCNDetail"
  
	sql = "delete from tblrcn_detail where rcnd_id=" & request("rcnd_id")	
	CUD(sql)

	sql = "select sum(rcnd_subtotal) as rcnd_subtotal from tblrcn_detail where rcnd_rcn_no = '" & request("rcn_no") & "'"
	rcnd_subtotal = selectid(sql)
	
	if isnull(rcnd_subtotal) then 
	   rcnd_subtotal = 0
	end if
	
	rcn_gstAmt = rcnd_subtotal * GSTRateBack
	
	if isnull(rcn_gstAmt) then 
	   rcn_gstAmt = 0
	end if
	
	sql = "update tblrcn set rcn_totalPartsAmt=" & rcnd_subtotal & ", rcn_gstAmt = " & chknumber0(rcn_gstAmt) & ", rcn_totalAmt=" & rcnd_subtotal & " where rcn_no = '" & ChkString(Request("rcn_no")) & "'"
	CUD(sql)
	
	url = "rm_rcn_new.asp?rcn_no=" & request("rcn_no") & "&loginerr=RCN Detail has been deleted.#spareparts" 

	 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn_detail','delRCNDetail=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	 CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "submitRCN"   
  
 sql="SELECT rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,   " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_no = '" & request("rcn_no") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("rcn_status") = "Submitted"
			   rs("rcn_submittedby") = Request.Cookies("GAPS")("sloginid")
			   rs("rcn_submitteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_rcn_view.asp?rcn_status=Submitted&loginerr=RCN has been Submitted.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn','submitRCN=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "PostedRCN"   
  
 sql="SELECT rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,   " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_no = '" & request("rcn_no") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("rcn_status") = "Posted"
			   rs("rcn_postedby") = Request.Cookies("GAPS")("sloginid")
			   rs("rcn_posteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_rcn_view.asp?rcn_status=Posted&loginerr=RCN has been Submitted.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn','PostedInvoice=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		 
'----------------------------------------------------------------------------------------------------    
  Case "CancelRCN"   
  
 sql="SELECT rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,   " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_no = '" & request("rcn_no") & "'"
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("rcn_status") = "Cancel"
			   rs("rcn_cancelledby") = Request.Cookies("GAPS")("sloginid")
			   rs("rcn_cancelleddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_rcn_view.asp?rcn_status=Cancel&loginerr=RCN has been Cancelled.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblrcn','CancelRCN=" & ChkString(left(request("rcn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)			 		 			 
'----------------------------------------------------------------------------------------------------    
   
    Case "addDO"   'this proc not used.DO created auto when inv is created
	    sql = "select ct_name from tblcity where ct_id =" & request("do_cust_city") 
		do_cust_city = selectid(sql)
		
		sql = "select state_name from tblstate where state_id =" & request("do_cust_state") 
		do_cust_state = selectid(sql)
		
		sql = "select state_code from tblstate where state_id =" & request("do_cust_state") 
		state_code = selectid(sql)
	
        ''''Add DO Order	   	  
  sql = "SELECT top 1 do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("do_status")  = "Open"   					
        rs("do_date") = ChkDateYYYYMMDD(date())
        rs("do_cust_code")  = ChkString(Request.Form("do_cust_code"))	
        rs("do_cust_name")  = ChkString(Request.Form("do_cust_name"))
		rs("do_cust_address") = ChkString(Request.Form("do_cust_address"))
		rs("do_cust_postcode") = ChkString(Request.Form("do_cust_postcode"))
		rs("do_cust_state") = do_cust_state
		rs("do_cust_state_id") = ChkString(Request.Form("do_cust_state")) 
		rs("do_cust_city") = do_cust_city
		rs("do_cust_city_id") = ChkString(Request.Form("do_cust_city")) 
		rs("do_cust_email") = ChkString(Request.Form("do_cust_email")) 
		rs("do_cust_tel1") = ChkString(Request.Form("do_cust_tel1")) 
		rs("do_cust_tel2") = ChkString(Request.Form("do_cust_tel2")) 
		rs("do_remark") = ChkString(Request.Form("do_remark")) 
		rs("do_createddate") = ChkDateTimeMySQL(now())
		rs("do_createdby") = Request.Cookies("GAPS")("sloginid")
		rs("do_tech_code") = ChkString(Request.Form("do_tech_code")) 
		
		if ChkString(Request.Form("do_purchase_date"))  <> "" then 
		rs("do_purchase_date") = ChkString(Request.Form("do_purchase_date")) 
		end if
		
		rs("do_onlineWrtyNo") = ChkString(Request.Form("do_onlineWrtyNo")) 
		rs("do_onlineWrtyStatus") = ChkString(Request.Form("do_onlineWrtyStatus")) 
		rs("do_SN_no") = ChkString(Request.Form("do_SN_no")) 
		rs("do_type") = ChkString(Request.Form("do_type")) 
		rs("do_Model") = ChkString(Request.Form("do_Model")) 
		rs("do_model_desc") = ChkString(Request.Form("do_model_desc")) 
		
		if ChkString(Request.Form("do_appointment_date"))  <> "" then
		rs("do_appointment_date") = ChkString(Request.Form("do_appointment_date")) 
		end if
		rs("do_appointment_time") = ChkString(Request.Form("do_appointment_time")) 
		rs("do_appointment_remark") = ChkString(Request.Form("do_appointment_remark")) 
		
		rs("do_inv_no") = ChkString(Request.Form("do_inv_no")) 
		if ChkString(Request.Form("do_inv_date"))  <> "" then
		rs("do_inv_date") = ChkString(Request.Form("do_inv_date")) 
		end if
		
		rs.Update 
		rs.Close      
		
        sql = "select top 1 do_id from tbldo order by do_id desc "
        do_id = selectid(sql)
		'temp = 100000 + do_id
       ' do_no = "DO" & temp 
		do_no = "DOS" & DONumbering(do_id) 
 
        sql = "update tbldo set do_no = '" & do_no & "' where do_id = " & do_id
        CUD(sql)
      
        url = "rm_do_new.asp?do_no=" & do_no & "&loginerr=New DO has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo','addDO=" & ChkString(left(do_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editDO"   
		
	if request.form("do_cust_postcode") <> ""  and request.form("do_cust_cnty_id") = "129" then
			do_cust_postcode=request("do_cust_postcode")
			sql = "select city_id from tblpostcode where postcode =" & do_cust_postcode	
			do_cust_city_id = selectid(sql)
	
			do_cust_city_id = request("do_cust_city_id")
			sql = "select ct_name2 from tblcity where ct_id =" & do_cust_city_id	
			do_cust_city = selectid(sql)

			sql = "select state_id from tblpostcode where postcode =" & do_cust_postcode	
			do_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & do_cust_postcode	
			do_cust_state = selectid(sql)
	elseif request.form("do_cust_postcode") = ""  and request.form("do_cust_cnty_id") = "129" then
			response.redirect("rm_do_new.asp?do_no=" & request("do_no") & "&loginerr=DO not updated.#articletitle")
	end if

	if  request.form("do_cust_cnty_id") <> "129" then 
			do_cust_city = request.form("do_cust_city")
	   	    do_cust_city_id = "0"
	end if
	
	    ''''Edit DO Order	 	   	  
  sql = "SELECT top 1 do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_cnty_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo WHERE do_no = '" & request("do_no") & "' "		    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			rs("do_cust_code")  = ChkString(Request.Form("do_cust_code"))	
			rs("do_cust_name")  = ChkString(Request.Form("do_cust_name"))
			rs("do_cust_address") = ChkString(Request.Form("do_cust_address"))
			rs("do_cust_postcode") = ChkString(Request.Form("do_cust_postcode"))
			
			if request.form("job_cust_cnty_id") = "129" then
				rs("do_cust_state") = do_cust_state
				rs("do_cust_state_id") = do_cust_state_id 
			end if

			rs("do_cust_city") = do_cust_city
			rs("do_cust_city_id") = do_cust_city_id
			rs("do_cust_cnty_id") = ChkString(Request.Form("do_cust_cnty_id"))
			rs("do_cust_email") = ChkString(Request.Form("do_cust_email")) 
			rs("do_cust_tel1") = ChkString(Request.Form("do_cust_tel1")) 
			rs("do_cust_tel2") = ChkString(Request.Form("do_cust_tel2")) 
			rs("do_remark") = ChkString(Request.Form("do_remark")) 
			rs("do_tech_code") = ChkString(Request.Form("do_tech_code")) 
			
			if ChkString(Request.Form("do_purchase_date"))  <> "" then 
			rs("do_purchase_date") = ChkString(Request.Form("do_purchase_date")) 
			end if

			rs("do_onlineWrtyNo") = ChkString(Request.Form("do_onlineWrtyNo")) 
			rs("do_onlineWrtyStatus") = ChkString(Request.Form("do_onlineWrtyStatus")) 
			rs("do_SN_no") = ChkString(Request.Form("do_SN_no")) 
			rs("do_type") = ChkString(Request.Form("do_type")) 
			rs("do_Model") = ChkString(Request.Form("do_Model")) 
			rs("do_model_desc") = ChkString(Request.Form("do_model_desc")) 
			
			if ChkString(Request.Form("do_appointment_date"))  <> "" then
			rs("do_appointment_date") = ChkString(Request.Form("do_appointment_date")) 
			end if
			
			rs("do_appointment_time") = ChkString(Request.Form("do_appointment_time")) 
			rs("do_appointment_remark") = ChkString(Request.Form("do_appointment_remark")) 
			
			rs("do_inv_no") = ChkString(Request.Form("do_inv_no")) 
		    if ChkString(Request.Form("do_inv_date"))  <> "" then
		       rs("do_inv_date") = ChkString(Request.Form("do_inv_date")) 
		    end if
		
		rs.Update 
		rs.Close 
		end if

        url = "rm_do_new.asp?do_no=" & request("do_no") & "&loginerr=DO has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblDO','editDO=" & ChkString(left(request("do_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	 
		 
'----------------------------------------------------------------------------------------------------    
  Case "addDODetail"   
  
        ''''Add DO parts	   	  
        sql = "SELECT top 1 dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, dod_discounttype, " & _
		      "dod_netcost, dod_subtotal FROM tbldo_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("dod_do_no") = ChkString(Request.Form("do_no"))
		rs("dod_inv_no") = ChkString(Request.Form("dod_inv_no"))
        rs("dod_job_code")  = ChkString(Request.Form("dod_job_code"))	
		rs("dod_partcode")  = ChkString(Request.Form("dod_partcode"))	
		rs("dod_desc")  = ChkString(Request.Form("dod_desc"))	
        rs("dod_unitcost")  = ChkString(Request.Form("dod_unitcost"))
		rs("dod_qty") = ChkString(Request.Form("dod_qty"))
		rs("dod_discountamt")  = 0
		rs("dod_discounttype")  = "%"
		rs("dod_netcost")  = 0
		rs("dod_subtotal")  = 0
		rs.Update 
		rs.Close      
		
        sql = "select sum(dod_qty) as dod_qty from tbldo_detail where dod_do_no = '" & request("do_no") & "'"
        dod_qty = selectid(sql)
		
		sql = "update tbldo set do_totalqty=" & dod_qty & " where do_no = '" & ChkString(Request("do_no")) & "'"
        CUD(sql)
		
        url = "rm_do_new.asp?do_no=" & request("do_no") & "&loginerr=DO has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addDODetail','tbldo_detail=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------    
  Case "editDODetail"   
	
        ''''Add DO parts	   	  
         sql = "SELECT dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, dod_discounttype, dod_netcost, dod_subtotal " & _
	          "FROM tbldo_detail where dod_id=" & request("dod_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("dod_partcode")  = ChkString(Request.Form("dod_partcode"))	
		rs("dod_desc")  = ChkString(Request.Form("dod_desc"))	
        rs("dod_unitcost")  = ChkString(Request.Form("dod_unitcost"))
		rs("dod_qty") = ChkString(Request.Form("dod_qty"))
		rs("dod_discountamt")  = ChkString(Request.Form("dod_discountamt"))
		rs("dod_discounttype")  = ChkString(Request.Form("dod_discounttype"))
		
		if ChkString(Request.Form("dod_discounttype")) = "%" then 
		dod_netcost  = ChkString(Request.Form("dod_unitcost")) * (ChkString(Request.Form("dod_discountamt")/100))
		dod_netcost = ChkString(Request.Form("dod_unitcost")) - dod_netcost
		rs("dod_netcost")  = dod_netcost
		rs("dod_subtotal")  = dod_netcost * ChkString(Request.Form("dod_qty"))
		else
		dod_netcost  = ChkString(Request.Form("dod_unitcost")) -  ChkString(Request.Form("dod_discountamt"))
		rs("dod_netcost")  = dod_netcost
		rs("dod_subtotal")  = dod_netcost * ChkString(Request.Form("dod_qty"))
		end if
		rs.Update 
		rs.Close  
		end if    
		
        sql = "select sum(dod_qty) as dod_qty from tbldo_detail where dod_do_no = '" & request("do_no") & "'"
        dod_qty = selectid(sql)
		
		sql = "update tbldo set do_totalqty=" & dod_qty & " where do_no = '" & ChkString(Request("do_no")) & "'"
        CUD(sql)
		
        url = "rm_do_new.asp?do_no=" & request("do_no") & "&loginerr=RCN has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo_detail','editDODetail=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "delDODetail"
  
	sql = "delete from tbldo_detail where dod_id=" & request("dod_id")	
	CUD(sql)

	sql = "select sum(dod_subtotal) as dod_subtotal from tbldo_detail where dod_do_no = '" & request("do_no") & "'"
	dod_subtotal = selectid(sql)
	
	if isnull(dod_subtotal) then 
	   dod_subtotal = 0
	end if
	
	do_gstAmt = dod_subtotal * 0.05660377358
	
	if isnull(do_gstAmt) then 
	   do_gstAmt = 0
	end if
	
	sql = "update tbldo set do_totalPartsAmt=" & dod_subtotal & ", do_gstAmt = " & chknumber0(do_gstAmt) & ", do_totalAmt=" & dod_subtotal & " where do_no = '" & ChkString(Request("do_no")) & "'"
	CUD(sql)
	
	url = "rm_do_new.asp?do_no=" & request("do_no") & "&loginerr=DO Detail has been deleted.#spareparts" 

	 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo_detail','delDODetail=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	 CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "DeliveredDO"   
  
  sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo WHERE do_no = '" & request("do_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("do_status") = "Delivered"
			   rs("do_deliveredby") = Request.Cookies("GAPS")("sloginid")
			   rs("do_delivereddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_do_view.asp?do_status=Delivered&loginerr=DO has been Delivered.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo','DeliveredDO=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "DoneDO"   
  
 sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo WHERE do_no = '" & request("do_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("do_status") = "Done"
			   rs("do_doneby") = Request.Cookies("GAPS")("sloginid")
			   rs("do_donedate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_do_view.asp?do_status=Done&loginerr=DO has been DoneDO.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo','DoneDO=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
		 		 
'----------------------------------------------------------------------------------------------------    
  Case "PostedDO"   
  
 sql1 = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo WHERE do_no = '" & request("do_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("do_status") = "Posted"
			   rs1("do_postedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("do_posteddate") = ChkDateTimeMySQL(now())
		rs1.Update 
		rs1.Close 
		end if
		
		'''''DO Detail
		sql1 = "SELECT dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, " & _
			   "dod_discounttype, dod_netcost, dod_subtotal " & _
			   "FROM tbldo_detail where dod_do_no = '" & request("do_no") & "' order by dod_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add DO Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("dod_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("dod_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("dod_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				rs2.Update 

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedDO=" & request("do_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("do_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "DO"
			rs2("stk_itm_code")  = ChkString(rs1("dod_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("dod_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("dod_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("dod_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close

        url = "rm_do_view.asp?do_status=Posted&loginerr=DO has been Posted.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbldo','PostedDO=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		 
'----------------------------------------------------------------------------------------------------    
  Case "CancelDO"   
  
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, " & _
		  "do_cust_state, do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby,  " & _
		  "do_job_code, do_tech_code, do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, " & _ 
		  "do_emailsentdate, do_submittedby, do_submitteddate, do_deliveredby, do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		  "do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time,  " & _
		  "do_appointment_remark FROM tbldo WHERE do_no = '" & request("do_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("do_status") = "Cancel"
			   rs("do_cancelledby") = Request.Cookies("GAPS")("sloginid")
			   rs("do_cancelleddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

		'''''DO Detail
		sql1 = "SELECT dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, " & _
			   "dod_discounttype, dod_netcost, dod_subtotal " & _
			   "FROM tbldo_detail where dod_do_no = '" & request("do_no") & "' order by dod_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add DO Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("dod_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("dod_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				rs2.Update 
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("do_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "DO-Cancel"
			rs2("stk_itm_code")  = ChkString(rs1("dod_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("dod_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("dod_qty"))
			rs2("stk_balanceqty")  = 0
			rs2("stk_sales_price")  = ChkNumber(rs1("dod_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		
        url = "rm_do_view.asp?do_status=Cancel&loginerr=DO has been Cancelled.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblDO','CancelDO=" & ChkString(left(request("do_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	

'----------------------------------------------------------------------------------------------------    
  Case "addWarehouse"   

	if request.form("wh_postcode") <> ""  and request.form("wh_code") = ""  then
	  if request.form("wh_city_id") = "" and request.form("wh_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			wh_postcode=request("wh_postcode")
			wh_name=request("wh_name")
			wh_tel=request("wh_tel")
			wh_fax=request("wh_fax")
			wh_address=request("wh_address")

			sql = "select city_id from tblpostcode where postcode =" & request("wh_postcode")	
			wh_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & wh_city_id	
			wh_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("wh_postcode")	
			wh_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("wh_postcode")	
			wh_state = selectid(sql)
	  	    Response.Redirect "rm_warehouse_new.asp?wh_code="&wh_code&"&wh_postcode="&wh_postcode&"&wh_name="&wh_name&"&wh_address="&wh_address&"&wh_tel="&wh_tel&"&wh_fax="&wh_fax& "&loginerr=Updated Address.#articletitle" 			
		end if
	end if


	if request.form("wh_postcode") = "" then
	   	    Response.Redirect "rm_warehouse_new.asp?wh_code="&wh_code&"&wh_postcode="&wh_postcode&"&wh_name="&wh_name&"&wh_address="&wh_address&"&wh_tel="&wh_tel&"&wh_fax="&wh_fax& "&loginerr=Updated Address.#articletitle" 			
	end if
		  
sql = "SELECT wh_id, wh_code, wh_name, wh_address, wh_postcode, wh_state_id, wh_state, wh_city_id, wh_city, wh_tel, wh_fax, wh_remark, " & _
      "wh_contact_person, wh_email, wh_status " & _
	  "FROM tblwarehouse WHERE wh_code = '" & request("wh_code") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof then
		    rs.addnew
			rs("wh_name") = request("wh_name")
			rs("wh_address") = request("wh_address")
			rs("wh_postcode") = request("wh_postcode")
			rs("wh_state_id") = request("wh_state_id")
			rs("wh_state") = request("wh_state")
			rs("wh_city_id") = request("wh_city_id") 
			rs("wh_city") =request("wh_city")
			rs("wh_tel")  = request("wh_tel")
			rs("wh_fax")  = request("wh_fax")
			rs("wh_remark")  = request("wh_remark")
			rs("wh_contact_person")  = request("wh_contact_person")
			rs("wh_email")   = request("wh_email")
			rs("wh_status")  = request("wh_status")
		    rs.Update 
			rs.Close 
		end if
	
		
		sql = "select top 1 wh_id from tblwarehouse order by wh_id desc "
        wh_id = selectid(sql)
		temp = wh_id
        wh_code = "W" & temp 
        sql = "update tblwarehouse set wh_code = '" & wh_code & "' where wh_id = " & wh_id
        CUD(sql)
		
		'''Update tbltechnician
		sql = "update tbltechnician set tech_wh_code='" & request("wh_code") & "' where tech_code='" & request("wh_contact_person") & "'"
		CUD(sql)

        url = "rm_warehouse_new.asp?wh_code=" & wh_code & "&loginerr=Store has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse','addWarehouse=" & ChkString(left(wh_code,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 		 
'----------------------------------------------------------------------------------------------------    
  Case "editWarehouse"   

		 if request.form("wh_postcode") <> ""  then
			wh_postcode=request.form("wh_postcode")	

			sql = "select city_id from tblpostcode where postcode =" & wh_postcode	
			wh_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & wh_city_id	
			wh_city = selectid(sql)	

			sql = "select state_id from tblpostcode where postcode =" & wh_postcode
			wh_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & wh_postcode 
			wh_state = selectid(sql)	
		 else
			Response.Redirect "rm_warehouse_new?wh_code="&request.form("wh_code")&"&loginerr=Store Not Updated.#articletitle"
		 end if

       sql = "SELECT wh_id, wh_code, wh_name, wh_address, wh_postcode, wh_state_id, wh_state, wh_city_id, wh_city, wh_tel, wh_fax, wh_remark, " & _
        "wh_contact_person, wh_email, wh_status " & _
	    "FROM tblwarehouse WHERE wh_code = '" & request("wh_code") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then
			rs("wh_name") = request("wh_name")
			rs("wh_address") = request("wh_address")
			rs("wh_postcode") = request("wh_postcode")
			rs("wh_state_id") = wh_state_id
			rs("wh_state") = wh_state
			rs("wh_city_id") = wh_city_id
			rs("wh_city") =wh_city
			rs("wh_tel")  = request("wh_tel")
			rs("wh_fax")  = request("wh_fax")
			rs("wh_remark")  = request("wh_remark")
			rs("wh_contact_person")  = request("wh_contact_person")
			rs("wh_email")   = request("wh_email")
			rs("wh_status")  = request("wh_status")
		    rs.Update 
		rs.Close 
		end if
		
		'''Update tbltechnician
		sql = "update tbltechnician set tech_wh_code='" & request("wh_code") & "' where tech_code='" & request("wh_contact_person") & "'"
		CUD(sql)
		
        url = "rm_warehouse_new.asp?wh_code=" & request("wh_code") & "&loginerr=Store has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse','addWarehouse=" & ChkString(left(request("wh_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    

 Case "AddStock"

		sql = "SELECT md_code,md_desc,md_category,md_model,md_barcode,md_type,md_status,md_unitprice,md_brands,md_rcpprice,md_costprice,md_averageecost, " &_ 
		"md_stock_uom,md_unitprice1,md_unitprice2,md_unitprice3,md_unitprice4,md_logby,md_logdate,md_color,md_size,md_group_type from tblmodel where md_code =  '" & request("md_code") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if rs.eof and request.form("md_code") <> "" and request.form("md_desc") <> "" then 
			rs.AddNew			
			   rs("md_code") = request.form("md_code")
			   rs("md_desc") = request.form("md_desc")
	           rs("md_category") = request.form("md_category")
	           rs("md_model") = request.form("md_model")
	           rs("md_barcode") = request.form("md_barcode")
	           rs("md_type") = request.form("md_type")
	           rs("md_status") = request.form("md_status")
	           rs("md_unitprice") = ChkNumber2(request.form("md_unitprice"))			   
	           rs("md_brands") = request.form("md_brands")
	           rs("md_rcpprice") = ChkNumber2(request.form("md_rcpprice"))
	           rs("md_costprice") = ChkNumber2(request.form("md_costprice"))
			   rs("md_averageecost") = ChkNumber2(request.form("md_costprice"))
               rs("md_stock_uom") = request.form("md_stock_uom") 
	           rs("md_unitprice1") = ChkNumber2(request.form("md_unitprice1"))
	           rs("md_unitprice2") = ChkNumber2(request.form("md_unitprice2"))
	           rs("md_unitprice3") = ChkNumber2(request.form("md_unitprice3"))
	           rs("md_unitprice4") = ChkNumber2(request.form("md_unitprice4"))
	           rs("md_color") = request.form("md_color")
	           rs("md_size") = request.form("md_size")
	           rs("md_group_type") = request.form("md_group_type")
			   rs("md_logby") = Request.Cookies("GAPS")("sloginid")
			   rs("md_logdate") = ChkDateTimeMySQL(now())	  	
		rs.Update
		rs.Close 
		end if

		url = "rm_stock_create.asp?md_code=" & request("md_code") & "&loginerr=Stock has been created.#articletitle" 
		
		sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblmodel','CreateStock=" & ChkString(left(request("md_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------
  Case "editStock"   
		
	 ' sql = "Update tblmodel set md_desc = '" & request.form("md_desc") & "',md_category = '" & request("md_category") & "',md_type = '" & request("md_type") & "'," & _
		'	  "md_status = '" & request("md_status") & "'," & _
		'	  "md_rcpprice = '" & request("md_rcpprice")  & "'" & _
		'	  "WHERE md_code = '" & request("md_code") & "'" 

		   sql = "Update tblmodel set " & _
		      "md_category = '" & request("md_category") & "'," & _
		      "md_model = '" & request("md_model") & "'," & _
			  "md_costprice = '" & request("md_costprice") & "'," & _
			  "md_averageecost = '" & request("md_averageecost") & "'," & _
			  "md_brands = '" & request("md_brands") & "'," & _
			  "md_desc = '" & request("md_desc") & "'," & _
			  "md_type = '" & request("md_type") & "'," & _
			  "md_status = '" & request("md_status") & "'," & _
			  "md_rcpprice = " & request("md_rcpprice") & "," & _
			  "md_unitprice1 = " & request("md_unitprice1") & "," & _
			  "md_unitprice2 = " & request("md_unitprice2") & "," & _
		  	  "md_unitprice3 = " & request("md_unitprice3") & "," & _
			  "md_unitprice4 = " & request("md_unitprice4") & " WHERE md_code = '" & request("md_code") & "'"

		CUD(sql)	  		  
				  
        url = "rm_stock_new.asp?md_code=" & request("md_code") & "&loginerr=Stock has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblmodel','editStock=" & ChkString(left(request("md_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)


'----------------------------------------------------------------------------------------------------    
   
    Case "addStockIn"   
		
	    ''''Add Stock-in Order	   	  
		  sql = "SELECT top 1 st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
				"st_approveddate, st_approvedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
				"FROM tblstockin "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("st_status")  = "Open"   					
        rs("st_date") = ChkString(Request.Form("st_date"))	
        rs("st_referenceno")  = ChkString(Request.Form("st_referenceno"))	
	    rs("st_fromwarehouse")  = ChkString(Request.Form("st_fromwarehouse"))
		rs("st_towarehouse") = ChkString(Request.Form("st_towarehouse"))
		rs("st_remark") = ChkString(Request.Form("st_remark"))
		rs("st_createddate") = ChkString(ChkDateTimeMySQL(now())) 
		rs("st_createdby") = ChkString(Request.Cookies("GAPS")("sloginid")) 
		rs("st_totalqty") = 0
		rs("st_totalaAmt") = 0
		rs.Update 
		rs.Close      
		
        sql = "select top 1 st_id from tblstockin order by st_id desc "
        st_id = selectid(sql)
		temp = 100000 + st_id
        st_no = "STIN" & temp 
  
        sql = "update tblstockin set st_no = '" & st_no & "' where st_id = " & st_id
        CUD(sql)
      
        url = "rm_stockin_new.asp?st_no=" & st_no & "&loginerr=New Stock-In has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockin','addStockIn=" & ChkString(left(st_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
   
    Case "addStockIn"   
		
	    ''''Add Stock-in Order	   	  
		  sql = "SELECT top 1 st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
				"st_approveddate, st_approvedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
				"FROM tblstockin "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("st_status")  = "Open"   					
        rs("st_date") = ChkString(Request.Form("st_date"))	
        rs("st_referenceno")  = ChkString(Request.Form("st_referenceno"))	
	    rs("st_fromwarehouse")  = ChkString(Request.Form("st_fromwarehouse"))
		rs("st_towarehouse") = ChkString(Request.Form("st_towarehouse"))
		rs("st_remark") = ChkString(Request.Form("st_remark"))
		rs("st_createddate") = ChkString(ChkDateTimeMySQL(now())) 
		rs("st_createdby") = ChkString(Request.Cookies("GAPS")("sloginid")) 
		rs("st_totalqty") = 0
		rs("st_totalaAmt") = 0
		rs.Update 
		rs.Close      
		
        sql = "select top 1 st_id from tblstockin order by st_id desc "
        st_id = selectid(sql)
		temp = 100000 + st_id
        st_no = "STIN" & temp 
  
        sql = "update tblstockin set st_no = '" & st_no & "' where st_id = " & st_id
        CUD(sql)
      
        url = "rm_stockin_new.asp?st_no=" & st_no & "&loginerr=New Stock-In has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockin','addStockIn=" & ChkString(left(st_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editStockIn"   
	
        ''''Edit Stock-In Order	 	   	  
sql = "SELECT top 1 st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
		"st_approveddate, st_approvedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "	    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
        rs("st_date") = ChkString(Request.Form("st_date"))	
        rs("st_referenceno")  = ChkString(Request.Form("st_referenceno"))	
	    rs("st_fromwarehouse")  = ChkString(Request.Form("st_fromwarehouse"))
		rs("st_towarehouse") = ChkString(Request.Form("st_towarehouse"))
		rs("st_remark") = ChkString(Request.Form("st_remark"))
		rs.Update 
		rs.Close 
		end if

      url = "rm_stockin_new.asp?st_no=" & request("st_no") & "&loginerr=Stock-In has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockin','editStockIn=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "addStockInDetail"   
		
        ''''Add Stock In Detail	   	  
        sql = "SELECT top 1 std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
	          "FROM tblstockin_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("std_st_no") = ChkString(Request.Form("st_no"))
		rs("std_itm_code") = ChkString(Request.Form("std_itm_code"))
        rs("std_itm_desc")  = ChkString(Request.Form("std_itm_desc"))	
		rs("std_unitcost")  = ChkString(Request.Form("std_unitcost"))	
		rs("std_qty")  = ChkString(Request.Form("std_qty"))	
        rs("std_subtotal")  = ChkString(Request.Form("std_qty")*Request.Form("std_unitcost"))
		rs.Update 
		rs.Close      
		
		
        sql = "select sum(std_qty) as std_qty from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
        std_qty = selectid(sql)
		
		sql = "select sum(std_subtotal) as std_subtotal from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
        std_subtotal = selectid(sql)
		
		sql = "update tblstockin set st_totalqty=" & std_qty & ",st_totalaAmt=" & std_subtotal & " where st_no = '" & Request("st_no") & "'"
        CUD(sql)
		
        url = "rm_stockin_new.asp?st_no=" & request("st_no") & "&loginerr=Stock-In Detail has been added.#spareparts" 
  
        sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockInDetail','tblstockin_detail=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
        CUD(sql)
 
'----------------------------------------------------------------------------------------------------    
  Case "editStockInDetail"   
	
        ''''Add Stock-In Detail   	  
        sql = "SELECT std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
	          "FROM tblstockin_detail where std_id=" & request("std_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("std_itm_code") = ChkString(Request.Form("std_itm_code"))
        rs("std_itm_desc")  = ChkString(Request.Form("std_itm_desc"))	
		rs("std_unitcost")  = ChkString(Request.Form("std_unitcost"))	
		rs("std_qty")  = ChkString(Request.Form("std_qty"))	
        rs("std_subtotal")  = ChkString(Request.Form("std_qty")*Request.Form("std_unitcost"))
		rs.Update 
		rs.Close  
		end if    
		
        sql = "select sum(std_qty) as std_qty from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
        std_qty = selectid(sql)
		
		sql = "select sum(std_subtotal) as std_subtotal from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
        std_subtotal = selectid(sql)
		
		sql = "update tblstockin set st_totalqty=" & std_qty & ",st_totalaAmt=" & std_subtotal & " where st_no = '" & Request("st_no") & "'"
        CUD(sql)
		
        url = "rm_stockin_new.asp?st_no=" & request("st_no") & "&loginerr=Stock-In Detail has been added.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockInDetail','editStockInDetail=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "delStockInDetail"
  
	sql = "delete from tblstockin_detail where std_id=" & request("std_id")	
	CUD(sql)

	sql = "select sum(std_qty) as std_qty from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
	std_qty = selectid(sql)
	if isnull(std_qty) then 
	   std_qty = 0
	end if
	
	sql = "select sum(std_subtotal) as std_subtotal from tblstockin_detail where std_st_no = '" & request("st_no") & "'"
	std_subtotal = selectid(sql)
	if isnull(std_subtotal) then 
	   std_subtotal = 0
	end if
	
	sql = "update tblstockin set st_totalqty=" & std_qty & ",st_totalaAmt=" & std_subtotal & " where st_no = '" & Request("st_no") & "'"
	CUD(sql)
	
        url = "rm_stockin_new.asp?st_no=" & request("st_no") & "&loginerr=Stock-In Detail has been deleted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockInDetail','delStockInDetail=" & ChkString(left(request("std_id"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "SubmitStockIn"   
  
  sql = "SELECT st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
		"st_approveddate, st_approvedby, st_submitteddate, st_submittedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("st_status") = "Submitted"
			   rs("st_submittedby") = Request.Cookies("GAPS")("sloginid")
			   rs("st_submitteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if
		
		
        url = "rm_stockin_view.asp?st_status=Submitted&loginerr=Stock-In Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','SubmitStockIn','tblstockin=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "ApproveStockIn"   
	    sql1 = "SELECT st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
		"st_approveddate, st_approvedby, st_submitteddate, st_submittedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("st_status") = "Approved"
			   rs1("st_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("st_approveddate") = ChkDateTimeMySQL(now())
			   st_no = rs1("st_no")
			   st_fromwarehouse = rs1("st_fromwarehouse")
			   st_towarehouse = rs1("st_towarehouse")
			   st_referenceno = rs1("st_referenceno")
			   st_date = rs1("st_date")
			   st_fromwarehouse = rs1("st_fromwarehouse")
		rs1.Update 
		rs1.Close 
		end if
		
		'Remark - Stock-In Detail
		sql1 = "SELECT std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
		       "FROM tblstockin_detail where std_st_no = '" & st_no & "' order by std_id"	   
		
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			'Remark -- Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & st_towarehouse & "' and wst_itm_code = '" & rs1("std_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = st_towarehouse
				rs2("wst_itm_code") = ChkString(rs1("std_itm_code"))
				wst_itm_current_qty = ChkString(rs1("std_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("std_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				
				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockIn+Add=" & st_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)

			else
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("std_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("std_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockIn+Update=" & st_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Update 
			rs2.Close   
			
			'remark Calculate averagecost and update to tblmodel
            'remark sql = "select SUM(wst_itm_current_qty) from tblwarehouse_stock where wst_itm_code='" & ChkString(rs1("std_itm_code")) & "'"
			
			'remark  - this is the sum that excluding the new incoming stock. This function is called by the screen rm_stockin_new.asp at the point of approving the new stock-30/09/2022
			sql = "Select sum(tblstocktran.stk_qty) as totalqty " & _
					"from tblstocktran where stk_itm_code='" & ChkString(rs1("std_itm_code")) & "'"
			
            totalQty = selectid(sql)
			if isnull(totalQty) or totalQty="" then
				totalQty="0"
			end if	 

			'remark tblmodel_avgcost will be the base for keeping avgcost history
		
			sql  = "select top 1 md_averagecost from tblmodel_avgcost where md_code = '" & ChkString(rs1("std_itm_code")) & "' order by md_date desc"
			md_averageecost = selectid(sql)
			if isnull(md_averageecost) or md_averageecost="" then
				md_averageecost="0"
			end if	 

			totalValue = totalQty * md_averageecost
					
			newTotalQty = totalQty + rs1("std_qty")    
			newTotalValue = totalValue + (rs1("std_qty")*rs1("std_unitcost"))
			
			if newTotalQty <> 0 then
				newAverageCost = newTotalValue / newTotalQty
			else 
				newAverageCost = 0
			end if

			if not isNull(newAverageCost) then
				newAverageCost=round(newAverageCost,2)
			end if 
							
			'remark-  - tblmodel_avgcost - holds the history and tblmodel holds the latest avg cost
			'remark Update Stocktrans - Stock Movement

			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("st_no")
			rs2("stk_reference") = st_towarehouse
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Stock-In"
			rs2("stk_itm_code")  = ChkString(rs1("std_itm_code"))
			rs2("stk_fromwarehouse")  = st_fromwarehouse
			rs2("stk_towarehouse")  = st_towarehouse
			rs2("stk_desc")  = ChkString(rs1("std_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("std_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("std_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("std_unitcost"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close 
	
		if newAveragecost <> md_averagecost then
			sql = "update tblmodel set md_averageecost=" & ChkNumber(newAverageCost) & " where md_code='" & ChkString(rs1("std_itm_code")) & "'"
			CUD(sql)
			
			sql = "insert into tblmodel_avgcost (md_code, md_averagecost, md_date) values ('" & ChkString(rs1("std_itm_code")) & "'," & ChkNumber(newAverageCost) & ", GETDATE())"
			CUD(sql)
		end if
	
			
		 
		rs1.movenext
		wend
		rs1.close

        url = "rm_stockin_view.asp?st_status=Approved&loginerr=Stock-In Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ApproveStockIn','tblstockin=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
			 
'----------------------------------------------------------------------------------------------------    
  Case "CancelStockIn"   
  
sql = "SELECT st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, " & _
		"st_approveddate, st_approvedby, st_submitteddate, st_submittedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("st_status") = "Cancel"
			   rs("st_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs("st_approveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_stockin_view.asp?st_status=Cancel&loginerr=Stock-In Detail has been Cancelled.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','CancelStockIn','tblstockin=" & ChkString(left(request("st_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		 	 		 		 		 			
'----------------------------------------------------------------------------------------------------    
   
    Case "addStockOut"   
	
	    ''''Add Stock-Out Order	   	  
		  sql = "SELECT top 1 so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, " & _
				"so_approveddate, so_approvedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
				"FROM tblstockOut "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("so_status")  = "Open"   					
        rs("so_date") = ChkString(Request.Form("so_date"))	
        rs("so_referenceno")  = ChkString(Request.Form("so_referenceno"))	
	    rs("so_fromwarehouse")  = ChkString(Request.Form("so_fromwarehouse"))
		rs("so_towarehouse") = ChkString(Request.Form("so_towarehouse"))
		rs("so_remark") = ChkString(Request.Form("so_remark"))
		rs("so_createddate") = ChkString(ChkDateTimeMySQL(now())) 
		rs("so_createdby") = ChkString(Request.Cookies("GAPS")("sloginid")) 
		rs("so_totalqty") = 0
		rs("so_totalaAmt") = 0
		rs.Update 
		rs.Close      
		
        sql = "select top 1 so_id from tblstockOut order by so_id desc "
        so_id = selectid(sql)
		temp = 100000 + so_id
        so_no = "STUT" & temp 
  
        sql = "update tblstockOut set so_no = '" & so_no & "' where so_id = " & so_id
        CUD(sql)
      
        url = "rm_stockOut_new.asp?so_no=" & so_no & "&loginerr=New Stock-Out has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockOut','addStockOut=" & ChkString(left(so_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editStockOut"   
	
        ''''Edit Stock-In Order	 	   	  
sql = "SELECT top 1 so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, " & _
		"so_approveddate, so_approvedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "	    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
        rs("so_date") = ChkString(Request.Form("so_date"))	
        rs("so_referenceno")  = ChkString(Request.Form("so_referenceno"))	
	    rs("so_fromwarehouse")  = ChkString(Request.Form("so_fromwarehouse"))
		rs("so_towarehouse") = ChkString(Request.Form("so_towarehouse"))
		rs("so_remark") = ChkString(Request.Form("so_remark"))
		rs.Update 
		rs.Close 
		end if

      url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockOut','editStockOut=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "addStockOutDetail"   

  if isNotExceed(ChkString(Request.Form("sod_itm_code")), ChkString(Request.Form("sod_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
        ''''Add Stock In Detail	   	  
        sql = "SELECT top 1 sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	          "FROM tblstockOut_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("sod_so_no") = ChkString(Request.Form("so_no"))
		rs("sod_itm_code") = ChkString(Request.Form("sod_itm_code"))
        rs("sod_itm_desc")  = ChkString(Request.Form("sod_itm_desc"))	
		rs("sod_unitcost")  = ChkString(Request.Form("sod_unitcost"))	
		rs("sod_qty")  = ChkString(Request.Form("sod_qty"))	
        rs("sod_subtotal")  = ChkString(Request.Form("sod_unitcost")*Request.Form("sod_qty"))
		rs.Update 
		rs.Close      
		
        sql = "select sum(sod_qty) as sod_qty from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
        sod_qty = selectid(sql)

		sql = "select sum(sod_subtotal) as sod_subtotal from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
        sod_subtotal = selectid(sql)
				
		sql = "update tblstockOut set so_totalqty=" & sod_qty & ",so_totalaAmt=" & sod_subtotal & " where so_no = '" & ChkString(Request("so_no")) & "'"
        CUD(sql)
		
        url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out Detail has been added.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockOutDetail','tblstockOut_detail=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
else
		url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out Detail has not been added.#spareparts" 
end if

'----------------------------------------------------------------------------------------------------    
  Case "editStockOutDetail"   
  if isNotExceed(ChkString(Request.Form("sod_itm_code")), ChkString(Request.Form("sod_qty"))) = "True" then ' 10/11/2022  - check against parts qty before issuing
        ''''Add Stock-In Detail   	  
        sql = "SELECT sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	          "FROM tblstockOut_detail where sod_id=" & request("sod_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("sod_itm_code") = ChkString(Request.Form("sod_itm_code"))
        rs("sod_itm_desc")  = ChkString(Request.Form("sod_itm_desc"))	
		rs("sod_unitcost")  = ChkString(Request.Form("sod_unitcost"))	
		rs("sod_qty")  = ChkString(Request.Form("sod_qty"))	
        rs("sod_subtotal")  = ChkString(Request.Form("sod_unitcost")*Request.Form("sod_qty"))
		rs.Update 
		rs.Close  
		end if    
		
        sql = "select sum(sod_qty) as sod_qty from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
        sod_qty = selectid(sql)

		sql = "select sum(sod_subtotal) as sod_subtotal from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
        sod_subtotal = selectid(sql)
				
		sql = "update tblstockOut set so_totalqty=" & sod_qty & ",so_totalaAmt=" & sod_subtotal & " where so_no = '" & ChkString(Request("so_no")) & "'"
        CUD(sql)
		
        url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out Detail has been added.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockOutDetail','editStockOutDetail=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	else
		url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out Detail has not been added.#spareparts" 
	end if
'----------------------------------------------------------------------------------------------------    
  Case "delStockOutDetail"
  
	sql = "delete from tblstockOut_detail where sod_id=" & request("sod_id")	
	CUD(sql)

	sql = "select sum(sod_qty) as sod_qty from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
	sod_qty = selectid(sql)		
	if isnull(sod_qty) then 
	   sod_qty = 0
	end if

	sql = "select sum(sod_subtotal) as sod_subtotal from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
	sod_subtotal = selectid(sql)
	if isnull(sod_subtotal) then 
	   sod_subtotal = 0
	end if
		
	sql = "update tblstockOut set so_totalqty=" & sod_qty & ",so_totalaAmt=" & sod_subtotal & " where so_no = '" & ChkString(Request("so_no")) & "'"
	CUD(sql)
	
	url = "rm_stockOut_new.asp?so_no=" & request("so_no") & "&loginerr=Stock-Out Detail has been deleted.#spareparts"  

	sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockOutDetail','delStockOutDetail=" & ChkString(left(request("sod_id"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "SubmitStockOut"   
  
  sql1 = "SELECT so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, " & _
		"so_approveddate, so_approvedby, so_submitteddate, so_submittedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
				sql2 = "select sod_itm_code from tblstockOut_detail where sod_so_no = '" & request("so_no") & "'"
				set rs2 = server.CreateObject("adodb.recordset")
				rs2.Open sql2,strconnect,3,3,&H0001
				can_be_submit=true	
				if not rs2.eof then  '250624 -  - dont allow submit if there are no parts or stk out details
					rs1("so_status") = "Submitted"
					rs1("so_submittedby") = Request.Cookies("GAPS")("sloginid")
					rs1("so_submitteddate") = ChkDateTimeMySQL(now())
				end if	
		rs1.Update 
		rs1.Close 
		end if
		
		
        url = "rm_stockOut_view.asp?so_status=Submitted&loginerr=Stock-Out Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','SubmitStockOut','tblstockOut=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "ApproveStockOut"   

		cantransfer=true
		errorcode=""	
	
		'-------check if any stock qty exceeds available qty in tech wh -230924
		
		sql3= "SELECT b.so_fromwarehouse, sod_id, sod_so_no, sod_itm_code, sod_qty FROM tblstockOut_detail a "  & _
		"inner join tblstockOut b  on a.sod_so_no = b.so_no where a.sod_so_no =  '" & request("so_no") & "' order by sod_id"	

		set rs3 = server.CreateObject("adodb.recordset")
		set rs4 = server.CreateObject("adodb.recordset")
		rs3.Open sql3,strconnect,3,3,&H0001
		
		while Not rs3.EOF and cantransfer=true
			sql4 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty " & _
	        "FROM tblwarehouse_stock where wst_wh_code = '" &  rs3("so_fromwarehouse") & "' and wst_itm_code = '" & rs3("sod_itm_code") & "'"
			rs4.Open sql4,strconnect,2,2,&H0001

			if not rs4.eof then 
                if rs4("wst_itm_current_qty") < rs3("sod_qty") then
					cantransfer=false					
					errorcode= rs3("sod_so_no") & " / " & rs4("wst_itm_code")

					sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
					Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockOut+Failed=" & errorcode & "','" & ChkDateTimeMySQL(now()) & "')"         
					CUD(sql)
	
					url = "rm_stockOut_view.asp?so_status=Failed&loginerr=Stock-Out Detail has error.#spareparts"  
	            end if						
	
			end if
			rs3.movenext	
			rs4.close
		wend
		rs3.Close
		'--------end of checking transfer can proceed now

  if cantransfer=true then
        sql1 = "SELECT so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, " & _
		"so_approveddate, so_approvedby, so_submitteddate, so_submittedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("so_status") = "Approved"
			   rs1("so_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("so_approveddate") = ChkDateTimeMySQL(now())
			   so_no = rs1("so_no")
			   so_fromwarehouse = rs1("so_fromwarehouse")
			   so_towarehouse = rs1("so_towarehouse")
		rs1.Update 
		rs1.Close 
		end if

		'''''Stock-Out Detail
		sql1 = "SELECT sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	           "FROM tblstockOut_detail where sod_so_no = '" & so_no & "' order by sod_id"	    
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & so_fromwarehouse & "' and wst_itm_code = '" & rs1("sod_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("sod_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("sod_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				rs2.Update 

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockOut+Update=" & so_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if			
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("so_no")
			rs2("stk_reference") = so_fromwarehouse
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Stock-Out"
			rs2("stk_itm_code")  = ChkString(rs1("sod_itm_code"))
			rs2("stk_fromwarehouse")  = so_fromwarehouse
			rs2("stk_towarehouse")  = so_towarehouse
			rs2("stk_desc")  = ChkString(rs1("sod_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sod_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sod_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		
        url = "rm_stockOut_view.asp?so_status=Approved&loginerr=Stock-Out Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ApproveStockOut','tblstockOut=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	end if		 
'----------------------------------------------------------------------------------------------------    
  Case "CancelStockOut"   
  
sql = "SELECT so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, " & _
		"so_approveddate, so_approvedby, so_submitteddate, so_submittedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("so_status") = "Cancel"
			   rs("so_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs("so_approveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_stockOut_view.asp?so_status=Cancel&loginerr=Stock-Out Detail has been Cancelled.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','CancelStockOut','tblstockOut=" & ChkString(left(request("so_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		  	 		 		 		 			 
'----------------------------------------------------------------------------------------------------   
   
    Case "addStockAdj"   
	
	    ''''Add Stock-Adj Order	   	  
		  sql = "SELECT top 1 sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
				"sj_approveddate, sj_approvedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
				"FROM tblstockadj "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("sj_status")  = "Open"   					
        rs("sj_date") = ChkString(Request.Form("sj_date"))	
        rs("sj_referenceno")  = ChkString(Request.Form("sj_referenceno"))	
	    rs("sj_fromwarehouse")  = ChkString(Request.Form("sj_fromwarehouse"))
		rs("sj_towarehouse") = ChkString(Request.Form("sj_towarehouse"))
		rs("sj_remark") = ChkString(Request.Form("sj_remark"))
		rs("sj_createddate") = ChkString(ChkDateTimeMySQL(now())) 
		rs("sj_createdby") = ChkString(Request.Cookies("GAPS")("sloginid")) 
		rs("sj_totalqty") = 0
		rs("sj_totalaAmt") = 0
		rs.Update 
		rs.Close      
		
        sql = "select top 1 sj_id from tblstockadj order by sj_id desc "
        sj_id = selectid(sql)
		temp = 100000 + sj_id
        sj_no = "STAJ" & temp 
  
        sql = "update tblstockadj set sj_no = '" & sj_no & "' where sj_id = " & sj_id
        CUD(sql)
      
        url = "rm_stockAdj_new.asp?sj_no=" & sj_no & "&loginerr=New Stock-Adj has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockadj','addStockAdj=" & ChkString(left(sj_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editStockAdj"   
	
        ''''Edit Stock-Adj Order	 	   	  
sql = "SELECT top 1 sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
		"sj_approveddate, sj_approvedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "	    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
        rs("sj_date") = ChkString(Request.Form("sj_date"))	
        rs("sj_referenceno")  = ChkString(Request.Form("sj_referenceno"))	
	    rs("sj_fromwarehouse")  = ChkString(Request.Form("sj_fromwarehouse"))
		rs("sj_towarehouse") = ChkString(Request.Form("sj_towarehouse"))
		rs("sj_remark") = ChkString(Request.Form("sj_remark"))
		rs.Update 
		rs.Close 
		end if

      url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstockadj','editStockAdj=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "addStockAdjDetail"   
  
        ''''Add Stock Adj Detail	   	  
        sql = "SELECT top 1 sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty , sjd_subtotal, sjd_referid " & _
	          "FROM tblstockadj_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("sjd_sj_no") = ChkString(Request.Form("sj_no"))
		rs("sjd_itm_code") = ChkString(Request.Form("sjd_itm_code"))
        rs("sjd_itm_desc")  = ChkString(Request.Form("sjd_itm_desc"))	
		rs("sjd_unitcost")  = ChkString(Request.Form("sjd_unitcost"))	
		rs("sjd_current_qty")  = ChkString(Request.Form("sjd_current_qty"))	
		rs("sjd_adjust_qty")  = ChkString(Request.Form("sjd_adjust_qty"))	


	if cint(Request.Form("sjd_adjust_qty")) >= 0 then 'cannot key-in adjusted stock < 0
		rs("sjd_diff_qty")  = cint(Request.Form("sjd_adjust_qty"))	- cint(Request.Form("sjd_current_qty")) 
	    rs("sjd_subtotal")  = ChkString(Request.Form("sjd_subtotal"))
		rs.Update 
		rs.Close      
		
        sql = "select sum(sjd_adjust_qty) as sjd_adjust_qty from tblstockadj_detail where sjd_sj_no = '" & request("sj_no") & "'"
        sjd_qty = selectid(sql)
		
		sql = "update tblstockadj set sj_totalqty=" & sjd_qty & " where sj_no = '" & ChkString(Request("sj_no")) & "'"
        CUD(sql)
		
        url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj Detail has been added.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockAdjDetail','tblstockadj_detail=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	else
	    url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj Detail has not been added.#spareparts" 
	end if 
'----------------------------------------------------------------------------------------------------    
  Case "editStockAdjDetail"   
	
        ''''edit Stock-In Detail   	  
        sql = "SELECT sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty , sjd_subtotal, sjd_referid " & _
	          "FROM tblstockadj_detail where sjd_id=" & request("sjd_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("sjd_itm_code") = ChkString(Request.Form("sjd_itm_code"))
        rs("sjd_itm_desc")  = ChkString(Request.Form("sjd_itm_desc"))	
		rs("sjd_unitcost")  = ChkString(Request.Form("sjd_unitcost"))	
		rs("sjd_current_qty")  = ChkString(Request.Form("sjd_current_qty"))	
		rs("sjd_adjust_qty")  = ChkString(Request.Form("sjd_adjust_qty"))	
	
		if cint(Request.Form("sjd_adjust_qty")) >= 0 then 'cannot key-in adjusted stock < 0
				rs("sjd_diff_qty")  = ChkString(Request.Form("sjd_adjust_qty"))	- ChkString(Request.Form("sjd_current_qty")) 
				rs("sjd_subtotal")  = ChkString(Request.Form("sjd_subtotal"))
				rs.Update 
				rs.Close  
			
				sql = "select sum(sjd_adjust_qty) as sjd_adjust_qty from tblstockadj_detail where sjd_sj_no = '" & request("sj_no") & "'"
				sjd_qty = selectid(sql)
		
				sql = "update tblstockadj set sj_totalqty=" & sjd_qty & " where sj_no = '" & ChkString(Request("sj_no")) & "'"
				CUD(sql)
		
				url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj Detail has been updated.#spareparts" 
  
				 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
					Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','editStockAdjDetail','tblstockadj_detail=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
				 CUD(sql)
		else
				url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj Detail has not been added.#spareparts" 
		end if
	end if 

'----------------------------------------------------------------------------------------------------    
  Case "delStockAdjDetail"
  
	sql = "delete from tblstockadj_detail where sjd_id=" & request("sjd_id")	
	CUD(sql)
	
	sql = "select sum(sjd_adjust_qty) as sjd_adjust_qty from tblstockadj_detail where sjd_sj_no = '" & request("sj_no") & "'"
	sjd_qty = selectid(sql)
	
	if isnull(sjd_qty) then 
	   sjd_qty = 0
	end if
	
	sql = "update tblstockadj set sj_totalqty=" & sjd_qty & " where sj_no = '" & ChkString(Request("sj_no")) & "'"
	CUD(sql)
	
	url = "rm_stockAdj_new.asp?sj_no=" & request("sj_no") & "&loginerr=Stock-Adj Detail has been updated.#spareparts" 

	 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','delStockAdjDetail','tblstockadj_detail=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	 CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "SubmitStockAdj"   
  
  sql = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
		"sj_approveddate, sj_approvedby, sj_submitteddate, sj_submittedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("sj_status") = "Submitted"
			   rs("sj_submittedby") = Request.Cookies("GAPS")("sloginid")
			   rs("sj_submitteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if
		
		
        url = "rm_stockAdj_view.asp?sj_status=Submitted&loginerr=Stock-Adj Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','SubmitStockAdj','tblstockadj=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "ApproveStockAdj"   
  
        sql1 = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
		"sj_approveddate, sj_approvedby, sj_submitteddate, sj_submittedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("sj_status") = "Approved"
			   rs1("sj_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("sj_approveddate") = ChkDateTimeMySQL(now())
			   sj_no = rs1("sj_no")
			   sj_fromwarehouse = rs1("sj_fromwarehouse")
			   sj_towarehouse = rs1("sj_towarehouse")
		rs1.Update 
		rs1.Close 
		end if
		
		'''''Stock-Adj Detail
		sql1 = "SELECT sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty, sjd_subtotal, sjd_referid " & _
	                   "FROM tblstockadj_detail where sjd_sj_no = '" & sj_no & "' order by sjd_id"	     
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock Adj Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & sj_fromwarehouse & "' and wst_itm_code = '" & rs1("sjd_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				stk_adjustsuccess=true
				temp_itm_current_qty =  rs2("wst_itm_current_qty") + rs1("sjd_diff_qty")
				if temp_itm_current_qty >= 0 then 
					rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("sjd_diff_qty") 
					rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
					rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
					rs2.Update 
					sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
					Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockAdj+Update=" & sj_no & "','" & ChkDateTimeMySQL(now()) & "')"         
					CUD(sql)
				else 
					stk_adjustsuccess=false
					sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
					Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockAdj+UpdateErr=" & sj_no & "','" & ChkDateTimeMySQL(now()) & "')"         
					CUD(sql)
				end if
			
			end if			
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
				if stk_adjustsuccess=true then
					sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
						   "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
					rs2.Open sql2,strconnect,2,2,&H0001
					rs2.AddNew   
					rs2("stk_voucherno") = request("sj_no")
					rs2("stk_reference") = sj_fromwarehouse
					rs2("stk_date")  = ChkDateTimeMySQL(now())
					rs2("stk_type")  = "Stock-Adj"
					rs2("stk_itm_code")  = ChkString(rs1("sjd_itm_code"))
					rs2("stk_fromwarehouse")  = sj_fromwarehouse
					rs2("stk_towarehouse")  = sj_towarehouse
					rs2("stk_desc")  = ChkString(rs1("sjd_itm_desc"))
					rs2("stk_qty")  = ChkNumber(rs1("sjd_diff_qty"))
					rs2("stk_balanceqty")  = 0
					rs2("stk_sales_price")  = ChkNumber(rs1("sjd_subtotal"))
					rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
					rs2("stk_logdate")  = ChkDateTimeMySQL(now())
					rs2.Update 
					rs2.Close  
				end if
		
		rs1.movenext
		wend
		rs1.close

        url = "rm_stockAdj_view.asp?sj_status=Approved&loginerr=Stock-Adj Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ApproveStockAdj','tblstockadj=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
			 
'----------------------------------------------------------------------------------------------------    
  Case "CancelStockAdj"   
  
sql = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
		"sj_approveddate, sj_approvedby, sj_submitteddate, sj_submittedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("sj_status") = "Cancel"
			   rs("sj_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs("sj_approveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_stockAdj_view.asp?sj_status=Cancel&loginerr=Stock-Adj Detail has been Cancelled.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','CancelStockAdj','tblstockadj=" & ChkString(left(request("sj_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)		 
		 		 		 		 			 
'----------------------------------------------------------------------------------------------------  
   
    Case "addStockTransfer"   
	
	    ''''Add Stock-transfer Order	   	  
		  sql = "SELECT top 1 sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
				"sf_approveddate, sf_approvedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate, sf_couriercompany " & _
				"FROM tblstocktransfer "		
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("sf_status")  = "Open"   					
        rs("sf_date") = ChkString(Request.Form("sf_date"))	
        rs("sf_referenceno")  = ChkString(Request.Form("sf_referenceno"))	
	    rs("sf_fromwarehouse")  = ChkString(Request.Form("sf_fromwarehouse"))
		rs("sf_towarehouse") = ChkString(Request.Form("sf_towarehouse"))
		rs("sf_remark") = ChkString(Request.Form("sf_remark"))
		rs("sf_couriercompany") = ChkString(Request.Form("sf_couriercompany"))
		rs("sf_createddate") = ChkString(ChkDateTimeMySQL(now())) 
		rs("sf_createdby") = ChkString(Request.Cookies("GAPS")("sloginid")) 
		rs("sf_totalqty") = 0
		rs("sf_totalaAmt") = 0
		rs.Update 
		rs.Close      
		
        sql = "select top 1 sf_id from tblstocktransfer order by sf_id desc "
        sf_id = selectid(sql)
		temp = 100000 + sf_id
        sf_no = "STFR" & temp 
  
        sql = "update tblstocktransfer set sf_no = '" & sf_no & "' where sf_id = " & sf_id
        CUD(sql)
      
        url = "rm_stocktfr_new.asp?sf_no=" & sf_no & "&loginerr=New Stock-Transfer has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstocktransfer','addStockTransfer=" & ChkString(left(sf_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editStockTransfer"   
	
        ''''Edit Stock-Transfer Order	 	   	  
sql = "SELECT top 1 sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
		"sf_approveddate, sf_approvedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate,sf_couriercompany " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "	    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
        rs("sf_date") = ChkString(Request.Form("sf_date"))	
        rs("sf_referenceno")  = ChkString(Request.Form("sf_referenceno"))	
	    rs("sf_fromwarehouse")  = ChkString(Request.Form("sf_fromwarehouse"))
		rs("sf_towarehouse") = ChkString(Request.Form("sf_towarehouse"))
		rs("sf_remark") = ChkString(Request.Form("sf_remark"))
		rs("sf_couriercompany") = ChkString(Request.Form("sf_couriercompany"))	
		rs.Update 
		rs.Close 
		end if

      url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblstocktransfer','editStockTransfer=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
'----------------------------------------------------------------------------------------------------    
  Case "addStockTransferDetail"   
  if isNotExceed(ChkString(Request.Form("sfd_itm_code")), ChkString(Request.Form("sfd_qty"))) = "True" then 	'10/11/2022  - check against parts qty before issuing
        ''''Add Stock Transfer Detail	   	  
        sql = "SELECT top 1 sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid, sfd_ex_qty " & _
	          "FROM tblstocktran_detail "	  	
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.Open sql,strconnect,2,2,&H0001
        rs1.AddNew   
        rs1("sfd_st_no") = ChkString(Request.Form("sf_no"))
		rs1("sfd_itm_code") = ChkString(Request.Form("sfd_itm_code"))
        rs1("sfd_itm_desc")  = ChkString(Request.Form("sfd_itm_desc"))	
		rs1("sfd_unitcost")  = ChkString(Request.Form("sfd_unitcost"))	
		rs1("sfd_qty")  = ChkString(Request.Form("sfd_qty"))	
        rs1("sfd_subtotal")  = ChkString(Request.Form("sfd_qty")*Request.Form("sfd_unitcost"))

		'sql2 = "select sum(stk_qty) as sfd_ex_qty from tblstocktran where stk_itm_code= '" & ChkString(Request.Form("sfd_itm_code")) & "' and stk_reference = '" & ChkString(Request.Form("sf_towarehouse")) & "'"
	    sql2 = "select wst_itm_current_qty from tblwarehouse_stock where wst_wh_code='" & ChkString(Request.Form("sf_towarehouse")) & "' and wst_itm_code='" & ChkString(Request.Form("sfd_itm_code")) & "'"
		sfd_ex_qty = selectid(sql2)
		rs1("sfd_ex_qty")  = sfd_ex_qty '20/6 -  - show existing quantity prior to approval

		rs1.Update 
		rs1.Close
				
        sql = "select sum(sfd_qty) as sfd_qty from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
        sfd_qty = selectid(sql)
		
		sql = "select sum(sfd_subtotal) as sfd_subtotal from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
        sfd_subtotal = selectid(sql)
		
		sql = "update tblstocktransfer set sf_totalqty=" & sfd_qty & ",sf_totalaAmt=" & sfd_subtotal & " where sf_no = '" & Request("sf_no") & "'"
        CUD(sql)
		
        url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer Detail has been added.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockTransferDetail','tblstocktransfer_detail=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 else
		url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer Detail has not been added.#spareparts" 
 end if
'----------------------------------------------------------------------------------------------------    
  Case "editStockTransferDetail"   
	 if isNotExceed(ChkString(Request.Form("sfd_itm_code")), ChkString(Request.Form("sfd_qty"))) = "True" then 	'10/11/2022  - check against parts qty before issuing
        ''''Add Stock-Transfer Detail   	  
        sql = "SELECT sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid, sfd_ex_qty " & _
	          "FROM tblstocktran_detail where sfd_id=" & request("sfd_id")		    	
	    set rs = server.CreateObject("adodb.recordset")
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then  
		rs("sfd_itm_code") = ChkString(Request.Form("sfd_itm_code"))
        rs("sfd_itm_desc")  = ChkString(Request.Form("sfd_itm_desc"))	
		rs("sfd_unitcost")  = ChkString(Request.Form("sfd_unitcost"))	
		rs("sfd_qty")  = ChkString(Request.Form("sfd_qty"))	
        rs("sfd_subtotal")  = ChkString(Request.Form("sfd_qty")*Request.Form("sfd_unitcost"))
		rs("sfd_ex_qty")  = ChkString(Request.Form("sfd_ex_qty"))	
		rs.Update 
		rs.Close  
		end if    
		
        sql = "select sum(sfd_qty) as sfd_qty from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
        sfd_qty = selectid(sql)
		
		sql = "select sum(sfd_subtotal) as sfd_subtotal from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
        sfd_subtotal = selectid(sql)
		
		sql = "update tblstocktransfer set sf_totalqty=" & sfd_qty & ",sf_totalaAmt=" & sfd_subtotal & " where sf_no = '" & Request("sf_no") & "'"
        CUD(sql)
		
        url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer Detail has been added.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockTransferDetail','editStockTransferDetail=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 else
		url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer Detail has not been added.#spareparts" 
 end if
'----------------------------------------------------------------------------------------------------    
  Case "delStockTransferDetail"
  
	sql = "delete from tblstocktran_detail where sfd_id=" & request("sfd_id")	
	CUD(sql)

	sql = "select sum(sfd_qty) as sfd_qty from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
	sfd_qty = selectid(sql)
	if isnull(sfd_qty) then 
	   sfd_qty = 0
	end if
	
	sql = "select sum(sfd_subtotal) as sfd_subtotal from tblstocktran_detail where sfd_st_no = '" & request("sf_no") & "'"
	sfd_subtotal = selectid(sql)
	if isnull(sfd_subtotal) then 
	   sfd_subtotal = 0
	end if
	
	sql = "update tblstocktransfer set sf_totalqty=" & sfd_qty & ",sf_totalaAmt=" & sfd_subtotal & " where sf_no = '" & Request("sf_no") & "'"
	CUD(sql)
	
        url = "rm_stocktfr_new.asp?sf_no=" & request("sf_no") & "&loginerr=Stock-Transfer Detail has been deleted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addStockTransferDetail','delStockTransferDetail=" & ChkString(left(request("sfd_id"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "SubmitStockTransfer"   
  
  sql = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
		"sf_approveddate, sf_approvedby, sf_submitteddate, sf_submittedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("sf_status") = "Submitted"
			   rs("sf_submittedby") = Request.Cookies("GAPS")("sloginid")
			   rs("sf_submitteddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if
		
		
        url = "rm_stocktfr_view.asp?sf_status=Submitted&loginerr=Stock-Transfer Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','SubmitStockTransfer','tblstocktransfer=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "ApproveStockTransfer"   

	cantransfer=true
		errorcode=""
		
		'-------check if any transfer qty exceeds available qty in tech wh -230924
		
		sql3= "SELECT b.sf_fromwarehouse, sfd_id, sfd_st_no, sfd_itm_code, sfd_qty, sfd_ex_qty FROM tblstocktran_detail a "  & _
		"inner join tblstocktransfer b " & _
		"on a.sfd_st_no = b.sf_no where a.sfd_st_no =  '" & request("sf_no") & "' order by sfd_id"	

		set rs3 = server.CreateObject("adodb.recordset")
		set rs4 = server.CreateObject("adodb.recordset")
		rs3.Open sql3,strconnect,3,3,&H0001
		
		while Not rs3.EOF and cantransfer=true
			sql4 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty " & _
	        "FROM tblwarehouse_stock where wst_wh_code = '" &  rs3("sf_fromwarehouse") & "' and wst_itm_code = '" & rs3("sfd_itm_code") & "'"
			rs4.Open sql4,strconnect,2,2,&H0001
	
			if not rs4.eof then 
                if rs4("wst_itm_current_qty") < rs3("sfd_qty") then
					cantransfer=false					
					errorcode= rs3("sfd_st_no") & " / " & rs4("wst_itm_code")

					sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
					Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockTransfer+Out+Failed=" & errorcode & "','" & ChkDateTimeMySQL(now()) & "')"         
					CUD(sql)

					url = "rm_stocktfr_view.asp?sf_status=Failedd&loginerr=Stock-Transfer Detail has error.#spareparts" 
	            end if						
	
			end if
			rs3.movenext	
			rs4.close
		wend
		rs3.Close
		'--------end of checking transfer can proceed now

	if cantransfer=true then '23/9/24 only transfer if there are enough qty   
        sql1 = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
		"sf_approveddate, sf_approvedby, sf_submitteddate, sf_submittedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("sf_status") = "Approved"
			   rs1("sf_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("sf_approveddate") = ChkDateTimeMySQL(now())
			   sf_no = rs1("sf_no")
			   sf_fromwarehouse = rs1("sf_fromwarehouse")
			   sf_towarehouse = rs1("sf_towarehouse")
		rs1.Update 
		rs1.Close 
		end if
		
		'''''Stock-Transfer Detail
		sql1 = "SELECT sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid,sfd_ex_qty " & _
		       "FROM tblstocktran_detail where sfd_st_no = '" & sf_no & "' order by sfd_id"	   
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock Transfer out	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _
	               "FROM tblwarehouse_stock where wst_wh_code = '" & sf_fromwarehouse & "' and wst_itm_code = '" & rs1("sfd_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then 
				wst_itm_current_qty  = rs2("wst_itm_current_qty") - rs1("sfd_qty")
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("sfd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockTransfer+Out=" & sf_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Update 
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("sf_no")
			rs2("stk_reference") = sf_fromwarehouse
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Stock-Transfer-Out"
			rs2("stk_itm_code")  = ChkString(rs1("sfd_itm_code"))
			rs2("stk_fromwarehouse")  = sf_fromwarehouse
			rs2("stk_towarehouse")  = sf_towarehouse
			rs2("stk_desc")  = ChkString(rs1("sfd_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sfd_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sfd_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("sfd_unitcost"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close  			
			
			''''Add Stock Transfer In	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _
	               "FROM tblwarehouse_stock where wst_wh_code = '" & sf_towarehouse & "' and wst_itm_code = '" & rs1("sfd_itm_code") & "'"	   
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = sf_towarehouse
				rs2("wst_itm_code") = ChkString(rs1("sfd_itm_code"))
				wst_itm_current_qty = ChkString(rs1("sfd_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("sfd_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockTransfer+In+Add=" & sf_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			else	 
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("sfd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("sfd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','ApproveStockTransfer+In+Update=" & sf_no & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Update 
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("sf_no")
			rs2("stk_reference") = sf_towarehouse
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "Stock-Transfer-In"
			rs2("stk_itm_code")  = ChkString(rs1("sfd_itm_code"))
			rs2("stk_fromwarehouse")  = sf_fromwarehouse
			rs2("stk_towarehouse")  = sf_towarehouse
			rs2("stk_desc")  = ChkString(rs1("sfd_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sfd_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sfd_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("sfd_unitcost"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close  
			
		rs1.movenext
		wend
		rs1.close

        url = "rm_stocktfr_view.asp?sf_status=Approved&loginerr=Stock-Transfer Detail has been Submitted.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ApproveStockTransfer','tblstocktransfer=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
	end if		 
'----------------------------------------------------------------------------------------------------    
  Case "CancelStockTransfer"   
  
sql = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
		"sf_approveddate, sf_approvedby, sf_submitteddate, sf_submittedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("sf_status") = "Cancel"
			   rs("sf_approvedby") = Request.Cookies("GAPS")("sloginid")
			   rs("sf_approveddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_stocktfr_view.asp?sf_status=Cancel&loginerr=Stock-Transfer Detail has been Cancelled.#spareparts"  
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','CancelStockTransfer','tblstocktransfer=" & ChkString(left(request("sf_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	

'----------------------------------------------------------------------------------------------------    
    Case "generateCN"   
	
		generateCN="N" 
		set rs = server.CreateObject("adodb.recordset")
  
	    sql="SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, inv_cust_city, inv_cust_city_id, inv_cust_cnty_id,inv_cust_email, " & _
			  "inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_job_code, inv_tech_code, inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt,  " & _
			  "inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent, inv_emailsentdate, inv_status " & _
			  "FROM tblinvoice where inv_no = '" & request("cn_invoice_no") & "'"
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			inv_no = rs("inv_no")
			inv_cust_code = rs("inv_cust_code")
			inv_cust_name = rs("inv_cust_name") 
			inv_cust_address = rs("inv_cust_address")
			inv_cust_postcode = rs("inv_cust_postcode")
			inv_cust_state = rs("inv_cust_state") 
			inv_cust_state_id = rs("inv_cust_state_id") 
			inv_cust_city = rs("inv_cust_city")
			inv_cust_city_id = rs("inv_cust_city_id") 
			inv_cust_cnty_id = rs("inv_cust_cnty_id") 
			inv_cust_email = rs("inv_cust_email")
			inv_cust_tel1 = rs("inv_cust_tel1") 
			inv_cust_tel2 = rs("inv_cust_tel2")
			inv_createddate = rs("inv_createddate")
			inv_createdby = rs("inv_createdby") 
			inv_job_code = rs("inv_job_code")
			inv_tech_code = rs("inv_tech_code")
			inv_totalqty = rs("inv_totalqty")
			inv_totalPartsAmt = rs("inv_totalPartsAmt")
			inv_labourAmt = rs("inv_labourAmt")
			inv_transportAmt = rs("inv_transportAmt")
			inv_gstAmt = rs("inv_gstAmt") 
			inv_gstRate = rs("inv_gstRate") 
			inv_gstCode = rs("inv_gstCode") 
			inv_totalAmt = rs("inv_totalAmt") 
			generateCN = "Y"
		else
		    response.Redirect("rm_cn_new.asp?loginerr=Invalid Invoice Number, please try again.")	
		end if
		rs.close
	  
	   if generateCN="Y" then 
	    sql = "SELECT top 1 cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_cnty_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("cn_status")  = "Open"   					
        rs("cn_date") = ChkDateYYYYMMDD(date())
        rs("cn_cust_code")  = ChkString(inv_cust_code)	
        rs("cn_cust_name")  = ChkString(inv_cust_name)
		rs("cn_cust_address") = ChkString(inv_cust_address)
		rs("cn_cust_postcode") = ChkString(inv_cust_postcode)
		rs("cn_cust_state") = inv_cust_state
		rs("cn_cust_state_id") = ChkString(inv_cust_state_id) 
		rs("cn_cust_city") = inv_cust_city
		rs("cn_cust_city_id") = ChkString(inv_cust_city_id) 
		rs("cn_cust_cnty_id") = ChkString(inv_cust_cnty_id) 
		rs("cn_cust_email") = ChkString(inv_cust_email) 
		rs("cn_cust_tel1") = ChkString(inv_cust_tel1) 
		rs("cn_cust_tel2") = ChkString(inv_cust_tel2) 
		rs("cn_remark") = ""
		rs("cn_createddate") = ChkDateTimeMySQL(now())
		rs("cn_createdby") = Request.Cookies("GAPS")("sloginid")
		
		rs("cn_job_code") = ""
		rs("cn_do_no") = ""
		rs("cn_inv_no") = ChkString(inv_no) 
		rs.Update 
		rs.Close      
		
        sql = "select top 1 cn_id from tblcn order by cn_id desc "
        cn_id = selectid(sql)
		temp = 100000 + cn_id
        cn_no = "CN" & temp 
  
        sql = "update tblcn set cn_no = '" & cn_no & "' where cn_id = " & cn_id
        CUD(sql)
		
	  
		'''''CN_detail
		sql1 = "SELECT invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_desc, invd_unitcost, invd_qty, invd_discountamt, " & _
				       "invd_discounttype, invd_netcost, invd_subtotal	FROM tblinvoice_detail where invd_inv_no = '" & inv_no & "' order by invd_id"	
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
			   
			           ''''Add CN parts	   	  
					sql = "SELECT top 1 cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, cnd_discounttype, " & _
						  "cnd_netcost, cnd_subtotal FROM tblcn_detail "	  	
					set rs = server.CreateObject("adodb.recordset")
					rs.Open sql,strconnect,2,2,&H0001
					rs.AddNew   
					rs("cnd_cn_no") = ChkString(cn_no)
					rs("cnd_inv_no") = ChkString(rs1("invd_inv_no"))
					rs("cnd_job_code")  = ChkString(rs1("invd_job_code"))
					rs("cnd_partcode")  = ChkString(rs1("invd_partcode"))
					rs("cnd_desc")  = ChkString(rs1("invd_desc"))
					rs("cnd_unitcost")  = ChkString(rs1("invd_unitcost"))
					rs("cnd_qty") =  ChkString(rs1("invd_qty"))
					rs("cnd_discountamt")  = ChkString(rs1("invd_discountamt"))
					rs("cnd_discounttype")  = ChkString(rs1("invd_discounttype"))
					rs("cnd_netcost")  = ChkString(rs1("invd_netcost"))
					rs("cnd_subtotal")  = ChkString(rs1("invd_subtotal"))
					rs.Update 
					rs.Close      
		
		rs1.movenext
		wend
		rs1.close
	
        sql = "select sum(cnd_qty) as cnd_qty from tblcn_detail where cnd_cn_no = '" & cn_no & "'"
	    cnd_qty = selectid(sql)
					
		sql = "select sum(cnd_subtotal) as cnd_subtotal from tblcn_detail where cnd_cn_no = '" & cn_no & "'"
		cn_totalPartsAmt = selectid(sql)
					
		if isnull(cn_totalPartsAmt) then 
		   cn_totalPartsAmt = 0
		end if
					
		cn_gstAmt = cn_totalPartsAmt * GSTRateBack
		cn_totalAmt = cn_totalPartsAmt 
					
		sql = "update tblcn set cn_totalqty=" & chknumber2(cnd_qty) & ", cn_totalPartsAmt=" & chknumber2(cn_totalPartsAmt) & ", cn_gstAmt = " & chknumber2(cn_gstAmt) & ", cn_totalAmt=" & cn_totalAmt & " where cn_no = '" & cn_no & "'"
	    CUD(sql)
		
        url = "rm_cn_new.asp?cn_no=" & cn_no & "&loginerr=CN has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','editCN=" & ChkString(left(cn_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
		else
		url = "rm_cn_new.asp?loginerr=Invalid Invoice Number, please try again." 
		end if 
	
'----------------------------------------------------------------------------------------------------    
   
    Case "addCN"   
	
	if request.form("cn_cust_postcode") <> ""  and request.form("cn_cust_code") = ""  and request.form("cn_cust_cnty_id") = "129" then
	  if request.form("cn_cust_city_id") = "" and request.form("cn_cust_state_id") = "" then 'first time entering & creating auto-fill for state/city	
			cn_cust_postcode=request("cn_cust_postcode")
			cn_cust_name=request("cn_cust_name")
			cn_cust_tel1=request("cn_cust_tel1")
			cn_cust_tel2=request("cn_cust_tel2")
			cn_cust_address=request("cn_cust_address")

			sql = "select city_id from tblpostcode where postcode =" & request("cn_cust_postcode")	
			cn_cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & cn_cust_city_id	
			cn_cust_city = selectid(sql)
	
			sql = "select state_id from tblpostcode where postcode =" & request("cn_cust_postcode")	
			cn_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & request("cn_cust_postcode")	
			cn_cust_state = selectid(sql)

		    Response.Redirect "rm_cn_new.asp?cn_no="&cn_code&"&cn_cust_postcode="&cn_cust_postcode&"&cn_cust_name="&cn_cust_name&"&cn_cust_address="&cn_cust_address&"&cn_cust_tel1="&cn_cust_tel1&"&cn_cust_tel2="&cn_cust_tel2& "&loginerr=Updated Address.#articletitle" 
		end if
	end if

	if request.form("cn_cust_cnty_id") = "129" then
	    sql = "select ct_name from tblcity where ct_id =" & request("cn_cust_city") 
		cn_cust_city = selectid(sql)
		
		sql = "select state_name from tblstate where state_id =" & request("cn_cust_state") 
		cn_cust_state = selectid(sql)
		
		sql = "select state_code from tblstate where state_id =" & request("cn_cust_state") 
		state_code = selectid(sql)
	end if

	if  request.form("cn_cust_cnty_id") <> "129" then 
			cn_cust_city = request.form("cn_cust_city")
	   	    cn_cust_city_id = "0"
	end if
        ''''Add CN Order	   	  
  sql = "SELECT top 1 cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
		rs("cn_status")  = "Open"   					
        rs("cn_date") = ChkDateYYYYMMDD(date())
        rs("cn_cust_code")  = ChkString(Request.Form("cn_cust_code"))	
        rs("cn_cust_name")  = ChkString(Request.Form("cn_cust_name"))
		rs("cn_cust_address") = ChkString(Request.Form("cn_cust_address"))
		rs("cn_cust_postcode") = ChkString(Request.Form("cn_cust_postcode"))

		if request.form("cn_cust_cnty_id") = "129" then 'state applies to Malaysia only
			rs("cn_cust_state") = cn_cust_state
			rs("cn_cust_state_id") = cn_cust_state_id 
		end if

		rs("cn_cust_state") = cn_cust_state
		rs("cn_cust_state_id") = ChkString(Request.Form("cn_cust_state")) 
		rs("cn_cust_city") = cn_cust_city
		rs("cn_cust_city_id") = cn_cust_city 
		rs("cn_cust_email") = ChkString(Request.Form("cn_cust_email")) 
		rs("cn_cust_tel1") = ChkString(Request.Form("cn_cust_tel1")) 
		rs("cn_cust_tel2") = ChkString(Request.Form("cn_cust_tel2")) 
		rs("cn_remark") = ChkString(Request.Form("cn_remark")) 
		rs("cn_createddate") = ChkDateTimeMySQL(now())
		rs("cn_createdby") = Request.Cookies("GAPS")("sloginid")
		
		rs("cn_job_code") = ChkString(Request.Form("cn_job_code")) 
		rs("cn_do_no") = ChkString(Request.Form("cn_do_no")) 
		rs("cn_inv_no") = ChkString(Request.Form("cn_inv_no")) 
		rs.Update 
		rs.Close      
		
        sql = "select top 1 cn_id from tblcn order by cn_id desc"
        cn_id = selectid(sql)
		temp = 100000 + cn_id
        cn_no = "CN" & temp 
  
        sql = "update tblcn set cn_no = '" & cn_no & "' where cn_id = " & cn_id
        CUD(sql)
      
        url = "rm_cn_new.asp?cn_no=" & cn_no & "&loginerr=New CN has been created.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','addCN=" & ChkString(left(cn_no,200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------     
     Case "editCN"   

		if request.form("cn_cust_postcode") <> ""  and request.form("cn_cust_cnty_id") = "129" then
			cn_cust_postcode=request("cn_cust_postcode")
			sql = "select city_id from tblpostcode where postcode =" & cn_cust_postcode	
			cn_cust_city_id = selectid(sql)

			sql = "select ct_name2 from tblcity where ct_id =" & cn_cust_city_id	
			cn_cust_city = selectid(sql)

			sql = "select state_id from tblpostcode where postcode =" & cn_cust_postcode	
			cn_cust_state_id = selectid(sql)

			sql = "select state_name from tblpostcode where postcode =" & cn_cust_postcode	
			cn_cust_state = selectid(sql)
	else
			response.redirect("rm_cn_new.asp?cn_no=" & request("cn_no") & "&loginerr=DO not updated.#articletitle")
	end if
	
        ''''Edit CN Order	 	   	  
   sql = "SELECT top 1 cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "		    
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
        rs("cn_cust_code")  = ChkString(Request.Form("cn_cust_code"))	
        rs("cn_cust_name")  = ChkString(Request.Form("cn_cust_name"))
		rs("cn_cust_address") = ChkString(Request.Form("cn_cust_address"))
		rs("cn_cust_postcode") = ChkString(Request.Form("cn_cust_postcode"))

		if request.form("cn_cust_cnty_id") = "129" then 'state applies to Malaysia only
			rs("cn_cust_state") = cn_cust_state
			rs("cn_cust_state_id") = cn_cust_state_id 
		end if
		
		rs("cn_cust_city") = cn_cust_city
		rs("cn_cust_city_id") =cn_cust_city_id 
		rs("cn_cust_email") = ChkString(Request.Form("cn_cust_email")) 
		rs("cn_cust_tel1") = ChkString(Request.Form("cn_cust_tel1")) 
		rs("cn_cust_tel2") = ChkString(Request.Form("cn_cust_tel2")) 
		rs("cn_remark") = ChkString(Request.Form("cn_remark")) 
		
		rs("cn_job_code") = ChkString(Request.Form("cn_job_code")) 
		rs("cn_do_no") = ChkString(Request.Form("cn_do_no")) 
		rs("cn_inv_no") = ChkString(Request.Form("cn_inv_no")) 
		rs.Update 
		rs.Close 
		end if

        url = "rm_cn_new.asp?cn_no=" & request("cn_no") & "&loginerr=CN has been updated.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','editCN=" & ChkString(left(request("cn_code"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	 
		 
'----------------------------------------------------------------------------------------------------    
  Case "addCNDetail"   
  
        ''''Add CN parts	   	  
        sql = "SELECT top 1 cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, cnd_discounttype, " & _
		      "cnd_netcost, cnd_subtotal FROM tblcn_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("cnd_cn_no") = ChkString(Request.Form("cn_no"))
		rs("cnd_inv_no") = ChkString(Request.Form("cnd_inv_no"))
        rs("cnd_job_code")  = ChkString(Request.Form("cnd_job_code"))	
		rs("cnd_partcode")  = ChkString(Request.Form("cnd_partcode"))	
		rs("cnd_desc")  = ChkString(Request.Form("cnd_desc"))	
        rs("cnd_unitcost")  = ChkString(Request.Form("cnd_unitcost"))
		rs("cnd_qty") = ChkString(Request.Form("cnd_qty"))
		rs("cnd_discountamt")  = ChkString(Request.Form("cnd_discountamt"))
		rs("cnd_discounttype")  = ChkString(Request.Form("cnd_discounttype"))
		
		if ChkString(Request.Form("cnd_discounttype")) = "%" then 
		cnd_netcost  = ChkString(Request.Form("cnd_unitcost")) * (ChkString(Request.Form("cnd_discountamt")/100))
		cnd_netcost = ChkString(Request.Form("cnd_unitcost")) - cnd_netcost
		rs("cnd_netcost")  = cnd_netcost
		rs("cnd_subtotal")  = cnd_netcost * ChkString(Request.Form("cnd_qty"))
		else
		cnd_netcost  = ChkString(Request.Form("cnd_unitcost")) -  ChkString(Request.Form("cnd_discountamt"))
		rs("cnd_netcost")  = cnd_netcost
		rs("cnd_subtotal")  = cnd_netcost * ChkString(Request.Form("cnd_qty"))
		end if
		rs.Update 
		rs.Close      
		
        sql = "select sum(cnd_qty) as cnd_qty from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
        cnd_qty = selectid(sql)
		
        sql = "select sum(cnd_subtotal) as cnd_subtotal from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
        cn_totalPartsAmt = selectid(sql)
		
		if isnull(cn_totalPartsAmt) then 
		   cn_totalPartsAmt = 0
		end if
		
        cn_gstAmt = cn_totalPartsAmt * GSTRateBack
		cn_totalAmt = cn_totalPartsAmt 
		
		sql = "update tblcn set cn_totalqty=" & cnd_qty & ", cn_totalPartsAmt=" & cn_totalPartsAmt & ", cn_gstAmt = " & chknumber2(cn_gstAmt) & ", cn_totalAmt=" & cn_totalAmt & " where cn_no = '" & request("cn_no") & "'"
        CUD(sql)
		
        url = "rm_cn_new.asp?cn_no=" & request("cn_no") & "&loginerr=CN has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','addCNDetail','tblcn_detail=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
 
'----------------------------------------------------------------------------------------------------    
  Case "editCNDetail"   
  
  	   sql = "delete from tblcn_detail where cnd_id=" & request("cnd_id")	
	   CUD(sql)
	
         ''''Add CN parts	   	  
        sql = "SELECT top 1 cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, cnd_discounttype, " & _
		      "cnd_netcost, cnd_subtotal FROM tblcn_detail "	  	
	    set rs = server.CreateObject("adodb.recordset")
	    rs.Open sql,strconnect,2,2,&H0001
        rs.AddNew   
        rs("cnd_cn_no") = ChkString(Request.Form("cn_no"))
		rs("cnd_inv_no") = ChkString(Request.Form("cnd_inv_no"))
        rs("cnd_job_code")  = ChkString(Request.Form("cnd_job_code"))	
		rs("cnd_partcode")  = ChkString(Request.Form("cnd_partcode"))	
		rs("cnd_desc")  = ChkString(Request.Form("cnd_desc"))	
        rs("cnd_unitcost")  = ChkString(Request.Form("cnd_unitcost"))
		rs("cnd_qty") = ChkString(Request.Form("cnd_qty"))
		rs("cnd_discountamt")  = ChkString(Request.Form("cnd_discountamt"))
		rs("cnd_discounttype")  = ChkString(Request.Form("cnd_discounttype"))
		
		if ChkString(Request.Form("cnd_discounttype")) = "%" then 
		cnd_netcost  = ChkString(Request.Form("cnd_unitcost")) * (ChkString(Request.Form("cnd_discountamt")/100))
		cnd_netcost = ChkString(Request.Form("cnd_unitcost")) - cnd_netcost
		rs("cnd_netcost")  = cnd_netcost
		rs("cnd_subtotal")  = cnd_netcost * ChkString(Request.Form("cnd_qty"))
		else
		cnd_netcost  = ChkString(Request.Form("cnd_unitcost")) -  ChkString(Request.Form("cnd_discountamt"))
		rs("cnd_netcost")  = cnd_netcost
		rs("cnd_subtotal")  = cnd_netcost * ChkString(Request.Form("cnd_qty"))
		end if
		rs.Update 
		rs.Close      
		
        sql = "select sum(cnd_qty) as cnd_qty from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
        cnd_qty = selectid(sql)
		
        sql = "select sum(cnd_subtotal) as cnd_subtotal from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
        cn_totalPartsAmt = selectid(sql)
		
		if isnull(cn_totalPartsAmt) then 
		   cn_totalPartsAmt = 0
		end if
		
        cn_gstAmt = cn_totalPartsAmt * GSTRateBack
		cn_totalAmt = cn_totalPartsAmt 
		
		sql = "update tblcn set cn_totalqty=" & cnd_qty & ", cn_totalPartsAmt=" & cn_totalPartsAmt & ", cn_gstAmt = " & chknumber2(cn_gstAmt) & ", cn_totalAmt=" & cn_totalAmt & " where cn_no = '" & request("cn_no") & "'"
        CUD(sql)
		
        url = "rm_cn_new.asp?cn_no=" & request("cn_no") & "&loginerr=CN has been updated.#spareparts" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn_detail','editCNDetail=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "delCNDetail"
  
	sql = "delete from tblcn_detail where cnd_id=" & request("cnd_id")	
	CUD(sql)

	sql = "select sum(cnd_qty) as cnd_qty from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
	cnd_qty = selectid(sql)
	if isnull(cnd_qty) then 
	   cnd_qty = 0
	end if
	
	sql = "select sum(cnd_subtotal) as cnd_subtotal from tblcn_detail where cnd_cn_no = '" & request("cn_no") & "'"
	cn_totalPartsAmt = selectid(sql)
	
	if isnull(cn_totalPartsAmt) then 
	   cn_totalPartsAmt = 0
	   cn_gstAmt = 0
	   cn_totalAmt = 0
	else
		cn_gstAmt = cn_totalPartsAmt * GSTRateBack
		cn_totalAmt = cn_totalPartsAmt    
	end if
	
	sql = "update tblcn set cn_totalqty=" & cnd_qty & ", cn_totalPartsAmt=" & cn_totalPartsAmt & ", cn_gstAmt = " & chknumber2(cn_gstAmt) & ", cn_totalAmt=" & cn_totalAmt & " where cn_no = '" & request("cn_no") & "'"
	CUD(sql)
	
	url = "rm_cn_new.asp?cn_no=" & request("cn_no") & "&loginerr=CN Detail has been deleted.#spareparts" 

	 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn_detail','delCNDetail=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	 CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "SubmitCN"   
  
   sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("cn_status") = "Submitted"
			   rs("cn_returnedby") = Request.Cookies("GAPS")("sloginid")
			   rs("cn_returneddate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_cn_view.asp?cn_status=Submitted&loginerr=CN has been Submitted.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','SubmitCN=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)

'----------------------------------------------------------------------------------------------------    
  Case "DoneCN"   
  
   sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("cn_status") = "Done"
			   rs("cn_doneby") = Request.Cookies("GAPS")("sloginid")
			   rs("cn_donedate") = ChkDateTimeMySQL(now())
		rs.Update 
		rs.Close 
		end if

        url = "rm_cn_view.asp?cn_status=Done&loginerr=CN has been Done.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','DoneCN=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)
		 
		 		 
'----------------------------------------------------------------------------------------------------    
  Case "PostedCN"   
  
   sql1 = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
	    set rs1 = server.CreateObject("adodb.recordset")
	    rs1.ActiveConnection = strconnect
		rs1.Source = sql1
		rs1.CursorLocation  = 3
		rs1.CursorType = 2
        rs1.LockType = 2
		rs1.Open
        if not rs1.eof then 
			   rs1("cn_status") = "Posted"
			   rs1("cn_postedby") = Request.Cookies("GAPS")("sloginid")
			   rs1("cn_posteddate") = ChkDateTimeMySQL(now())
			   cn_inv_no = rs1("cn_inv_no")		
			   cn_totalAmt = rs1("cn_totalAmt")		 	   
		rs1.Update 
		rs1.Close 
		end if
		
		'''''CN Detail
		sql1 = "SELECT cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, " & _
			   "cnd_discounttype, cnd_netcost, cnd_subtotal " & _
			   "FROM tblcn_detail where cnd_cn_no = '" & request("cn_no") & "' order by cnd_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add CN Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("cnd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("cnd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("cnd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				rs2.Update 

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','PostedCN+Add-Update=" & request("cn_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("cn_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "CN"
			rs2("stk_itm_code")  = ChkString(rs1("cnd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("cnd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("cnd_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("cnd_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		
		sql = "Update tblinvoice set inv_cnamount=inv_cnamount+" & cn_totalAmt & " where inv_no = '" & cn_inv_no & "' "
		CUD(sql)
		
		sql = "Update tblinvoice set inv_balance=inv_totalAmt-(inv_cnamount+inv_payment) where inv_no = '" & cn_inv_no & "' "
		CUD(sql)

        url = "rm_cn_view.asp?cn_status=Posted&loginerr=CN has been Posted.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcn','PostedCN=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
		 
'----------------------------------------------------------------------------------------------------    
  Case "CancelCN"   
  
   sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
		"cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
		"cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
		"cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
		"cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
		"FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("cn_status") = "Cancel"
			   rs("cn_cancelledby") = Request.Cookies("GAPS")("sloginid")
			   rs("cn_cancelleddate") = ChkDateTimeMySQL(now())
			   cn_totalAmt = rs("cn_totalAmt")
			   cn_inv_no = rs("cn_inv_no")
		rs.Update 
		rs.Close 
		end if
		
		if cn_totalAmt = "" then 
		   cn_totalAmt = 0
		end if
		
		''Update invoice 
		sql = "update tblinvoice set inv_cnamount=inv_cnamount-" & cn_totalAmt & " where inv_no='" & cn_inv_no & "'"
		CUD(sql)
		
		sql = "update tblinvoice set inv_balance = inv_totalAmt - inv_payment - inv_cnamount where inv_no='" & cn_inv_no & "'"
		CUD(sql)

		'''''CN Detail
		sql1 = "SELECT cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, " & _
			   "cnd_discounttype, cnd_netcost, cnd_subtotal " & _
			   "FROM tblcn_detail where cnd_cn_no = '" & request("cn_no") & "' order by cnd_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add CN Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("cnd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("cnd_qty") 
				rs2("wst_lastupdateby")  = Request.Cookies("GAPS")("sloginid")
				rs2("wst_lastupdatedate")  = ChkDateTimeMySQL(now())
				rs2.Update 

				sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
				Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblwarehouse_stock','CancelCN+Add CN+Update=" & request("cn_no") & "','" & ChkDateTimeMySQL(now()) & "')"         
				CUD(sql)
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = request("cn_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = ChkDateTimeMySQL(now())
			rs2("stk_type")  = "CN-Cancel"
			rs2("stk_itm_code")  = ChkString(rs1("cnd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("cnd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("cnd_qty")*-1)
			rs2("stk_balanceqty")  = 0
			rs2("stk_sales_price")  = ChkNumber(rs1("cnd_subtotal"))
			rs2("stk_logby")  = Request.Cookies("GAPS")("sloginid")
			rs2("stk_logdate")  = ChkDateTimeMySQL(now())

			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		
        url = "rm_cn_view.asp?cn_status=Cancel&loginerr=CN has been Cancelled.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblCN','CancelCN=" & ChkString(left(request("cn_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	

'----------------------------------------------------------------------------------------------------    
  Case "CancelReceipt"   

   sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
	"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
	"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
	"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
	"FROM tblreceipt WHERE receipt_no = '" & request("receipt_no") & "' "
	    set rs = server.CreateObject("adodb.recordset")
	    rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
        rs.LockType = 2
		rs.Open
        if not rs.eof then 
			   rs("receipt_status") = "Cancel"
			   rs("receipt_cancelledby") = Request.Cookies("GAPS")("sloginid")
			   rs("receipt_cancelleddate") = ChkDateTimeMySQL(now())
			   receipt_inv_no = rs("receipt_inv_no")
			   receipt_totalpayment = rs("receipt_totalpayment")
		rs.Update 
		rs.Close 
		end if

		''Update invoice 
		sql = "update tblinvoice set inv_payment=inv_payment-" & receipt_totalpayment & " where inv_no='" & receipt_inv_no & "'"
		CUD(sql)
		
		sql = "update tblinvoice set inv_balance = inv_totalAmt - inv_payment - inv_cnamount where inv_no='" & receipt_inv_no & "'"
		sql = "update tblinvoice set inv_balance = inv_totalAmt - inv_payment - inv_cnamount where inv_no='" & receipt_inv_no & "'"
		CUD(sql)
				
        url = "rm_receipt_view.asp?receipt_status=Cancel&loginerr=Receipt has been Cancelled.#articletitle" 
  
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblreceipt','CancelReceipt=" & ChkString(left(request("receipt_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	
 
'----------------------------------------------------------------------------------------------------     
     Case "editReceipt"   
	
        ''''Edit Receipt   	  
			sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
					"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
					"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
					"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
					"FROM tblreceipt where receipt_no='" & request("receipt_no") & "'"
			set rs = server.CreateObject("adodb.recordset")
			rs.ActiveConnection = strconnect
			rs.Source = sql
			rs.CursorLocation  = 3
			rs.CursorType = 2
			rs.LockType = 2
			rs.Open
			if not rs.eof then  
			rs("receipt_remark") = ChkString(request("receipt_remark"))
			receipt_inv_no = rs("receipt_inv_no") '300724 needed to postback the value to caller
			rs.Update 
			end if
			rs.Close  

         url = "rm_receipt_new.asp?receipt_no="&request("receipt_no")&"&receipt_inv_no="&receipt_inv_no&"&loginerr=Receipt has been updated.#articletitle" 
			
         sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	        Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblreceipt','editReceipt=" & ChkString(left(request("receipt_no"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
         CUD(sql)	   		 	 		 		 		 			
'----------------------------------------------------------------------------------------------------    

		 
End Select
Response.Clear
Response.Redirect(url)
%>

