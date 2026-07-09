<!-- #include file="database/datastore.asp" -->
<%Dim act,url,i,var,img(2),Upload,file
act = Request.QueryString("act")
Function chkEmthy(itm)
  If Len(itm) = 0 Then
    chkEmthy = "  "
  Else
    chkEmthy = itm
  End If
End Function
Function chkCInt(itm)
  If Len(itm) = 0 Then
    chkCInt = 0
  Else
    chkCInt = itm
  End If
End Function

	set rs = server.CreateObject("adodb.recordset")
	
Select Case act
      
'---------------------------------------------------------------------------------------------------

  Case "UpdatePAS"
  
   Lpwd = ""     
   gotpwd = ""   
   gotpwdhistory = ""   
    
    sql = "select top 1 password from tblusers where user_name = '" & Request.Cookies("GAPS")("sloginid") & "' "
	cpass=selectid(sql) 
  	
    Lpwd = LenEncrypt(Request.Cookies("GAPS")("sloginid"),request("current_password"))
	
   if cstr(Lpwd) <> cstr(cpass) then      	  
	  if Request.Cookies("GAPS")("forcechangepassword") = "yes" then
         response.clear 
		 response.Redirect "changepassword.asp?loginerr=Invalid Current Password."		  
	  else
		 response.clear 
		 response.Redirect "changepassword.asp?loginerr=Invalid Current Password."
	  end if	  	  
   end if 
		 	
	'Update Password
	sql1 = "UPDATE tblusers SET "
	sql1 = sql1 & "password= '" & ChkString(LenEncrypt(ChkString(Request.Cookies("GAPS")("sloginid")),ChkString(Request.Form("new_password")))) & "', "
	sql1 = sql1 & "lastchangepassword = '" & ChkDateTimeMySQL(now()) & "' "
	sql1 = sql1 & "where user_name = '" & Request.Cookies("GAPS")("sloginid") & "'"
    CUD(sql1)
  
	Response.Cookies("GAPS")("forcechangepassword") = "no"
	
	  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
		  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblusers','UpdatePAS=" & ChkString(left(Request.Cookies("GAPS")("sloginid"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
	  CUD(sql)
	
	Response.Clear
    Response.Redirect "changepassword.asp?loginerr=Password has been updated."

'---------------------------------------------------------------------------------------------------

  Case "addusers"
   gotsloginid = ""   
   sql = "select user_name from tblusers where user_name = '" & ChkString(Request.Form("user_name")) & "'"
   gotsloginid = selectid(sql)
      
   if gotsloginid = ChkString(Request.Form("user_name")) then   
      url = "mis_user_edit.asp?type=addtblusers&loginerr=User Name: " & Request.Form("user_name") & " has been used."
   else
   
		 set rs = server.CreateObject("adodb.recordset")
		 sql = "SELECT   user_id, createddate, user_name, password, fullname, user_type, staff_id, department, email, contactno, country, address1, address2, " & _
				"city, state, zipcode, user_active, accesslevel, OutletArea, OutletID, OutletRegion, Outletselection, SupplierID, TransID, lastlogindate, lastchangepassword, lastloginIP, log_by, log_date, log_ip, job_tech_code,verify_claim,approve_stk  " & _
				"FROM         tblusers where user_id is not null"
		  rs.ActiveConnection = strconnect 
		  rs.Source = sql
		  rs.LockType = 3	  
		  rs.Open   
		  rs.AddNew 
		  rs("createddate") = ChkDateTimeMySQL(now())
		  rs("user_name") = ChkString(Request.Form("user_name"))	 
		  rs("password") = LenEncrypt(ChkString(Request.Form("user_name")),ChkString(Request.Form("password"))) 	 
		  rs("fullname") = ChkString(Request.Form("fullname"))	 
		  rs("user_type") = ChkString(Request.Form("user_type"))	 
		  rs("staff_id") = ChkString(Request.Form("staff_id"))	 
		  rs("department") = ChkString(Request.Form("department"))	 
		  rs("email") = ChkString(Request.Form("email"))	 
		  rs("contactno") = ChkString(Request.Form("contactno"))	 
		  rs("country") = ChkString(Request.Form("country"))	 
		  rs("address1") = ChkString(Request.Form("address1"))	 
		  rs("address2") = ChkString(Request.Form("address2"))	 
		  rs("city") = ChkString(Request.Form("city"))	 
		  rs("state") = ChkString(Request.Form("state"))	 
		  rs("zipcode") = ChkString(Request.Form("zipcode"))
		  rs("user_active") = ChkString(Request.Form("user_active"))
		  rs("accesslevel") = ChkString(Request.Form("accesslevel"))		  
		  rs("OutletArea") = "1"
		  rs("OutletRegion") = ChkString(Request.Form("OutletRegion"))		  
		  rs("OutletID") = ChkString(Request.Form("target_fromoutletID"))	
		  rs("Outletselection") = "outletarea"	  
		  rs("SupplierID") = ChkString(Request.Form("SupplierID"))
		  rs("TransID") = ChkString(Request.Form("Supplier1"))
		  rs("lastchangepassword") = ChkDateTimeMySQL(now())
		  rs("log_by") = Request.Cookies("GAPS")("sloginid")
		  rs("log_date") = ChkDateTimeMySQL(now())
		  rs("log_ip") = Request.servervariables("remote_addr")
		  rs("job_tech_code") = ChkString(Request.Form("job_tech_code"))	
	  	  rs("verify_claim") = "N"
		  rs("approve_stk") = "N"
		  rs.Update 
		  rs.Close    		 
		  url = "mis_user_view.asp?type=searchdata&loginerr=New User Name: " & request("user_name") & " has been added."
			
		  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
			  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblusers','addtblusers=" & ChkString(left(Request("user_name"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
		  CUD(sql)
    
   end if
   
'----------------------------------------------------------------------------------------------------
  Case "editusers"
  
  sql = "SELECT     user_id, createddate, user_name, password, fullname, user_type, staff_id, department, email, contactno, country, address1, address2, " & _
		"city, state, zipcode, user_active, accesslevel, OutletArea, OutletRegion, OutletID, Outletselection, supplierID,TransID,lastlogindate, lastchangepassword, lastloginIP, log_by, log_date, log_ip, job_tech_code  " & _
		"FROM         tblusers where user_id = " & request("user_id")
  rs.ActiveConnection = strconnect 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  if not rs.eof then
  if Request.Form("oldpassword") <> Request.Form("password") then
  rs("password") = LenEncrypt(ChkString(Request.Form("user_name")),ChkString(Request.Form("password"))) 	 
  end if
  user_name = rs("user_name")
  rs("fullname") = ChkString(Request.Form("fullname"))	 
  rs("user_type") = ChkString(Request.Form("user_type"))	 
  rs("staff_id") = ChkString(Request.Form("staff_id"))	 
  rs("department") = ChkString(Request.Form("department"))	 
  rs("email") = ChkString(Request.Form("email"))	 
  rs("contactno") = ChkString(Request.Form("contactno"))	 
  rs("country") = ChkString(Request.Form("country"))	 
  rs("address1") = ChkString(Request.Form("address1"))	 
  rs("address2") = ChkString(Request.Form("address2"))	 
  rs("city") = ChkString(Request.Form("city"))	 
  rs("state") = ChkString(Request.Form("state"))	
  rs("zipcode") = ChkString(Request.Form("zipcode"))
  rs("user_active") = ChkString(Request.Form("user_active"))
  rs("accesslevel") = ChkString(Request.Form("accesslevel"))
  rs("OutletArea") = "1"
  rs("OutletRegion") = ChkString(Request.Form("OutletRegion"))  
  rs("OutletID") = ChkString(Request.Form("target_fromoutletID")) 
  rs("Outletselection") = "outletarea"	  
  rs("SupplierID") = ChkString(Request.Form("SupplierID"))	
  rs("TransID") = ChkString(Request.Form("Supplier1"))	''Added by bhaskar on 10-Nov-2011
  rs("lastchangepassword") = ChkDateTimeMySQL(now())	 
  rs("log_by") = Request.Cookies("GAPS")("sloginid")
  rs("log_date") = ChkDateTimeMySQL(now())
  rs("log_ip") = Request.servervariables("remote_addr")
  rs("job_tech_code") = ChkString(Request.Form("job_tech_code"))	 
  rs.Update 
  end if
  rs.Close    
  
  'response.write ChkString(Request.Form("target_fromoutletID")) 
  'response.End()
  
  set rs = nothing      
  url = "mis_user_edit.asp?type=editusers&user_id=" & request("user_id") & "&loginerr=User ID: " & request("user_id") & " has been updated."
	
 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblusers','edittblusers=" & ChkString(left(Request("user_name"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
 CUD(sql)

'----------------------------------------------------------------------------------------------------
  Case "deleteusers"
    sql = "DELETE FROM tblusers WHERE user_id=" & Request("user_id") 
    Call CUD(sql)	

    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblusers','deletetblusers=" & Request("user_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)

	url = "mis_user_view.asp?type=searchdata&loginerr=User ID: " & Request("user_id") & " has been deleted."  

'---------------------------------------------------------------------------------------------------

  Case "addnews"
   
  set rs = server.CreateObject("adodb.recordset")
  sql = "SELECT  top 1 news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews "
  rs.ActiveConnection = strconnect 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  rs.AddNew 
  rs("news_date") = date()
  rs("news_grouplevel") = ChkString(Request.Form("news_grouplevel"))	 
  rs("news_type") = ChkString(Request.Form("news_type"))	 
  rs("news_title") = ChkString(Request.Form("news_title"))	 
  rs("news_description") = ChkString(Request.Form("news_description"))	 
  rs("news_active") = ChkString(Request.Form("news_active"))	  
  rs("log_by") = Request.Cookies("GAPS")("sloginid")
  rs("log_date") = ChkDateTimeMySQL(now())
  rs.Update 
  rs.Close    		 
  url = "mis_news_view.asp?loginerr=News has been added."
  
  sql = "select top 1 news_id  FROM  tblNews order by news_id desc"
  news_id = selectid(sql)
  
  '''Update strconnectLingerie Database
  set rs = server.CreateObject("adodb.recordset")
  sql = "SELECT  top 1 news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id, renews_id  FROM  tblNews "
  rs.ActiveConnection = strconnectLingerie 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  rs.AddNew 
  rs("news_date") = date()
  rs("news_grouplevel") = ChkString(Request.Form("news_grouplevel"))	 
  rs("news_type") = ChkString(Request.Form("news_type"))	 
  rs("news_title") = ChkString(Request.Form("news_title"))	 
  rs("news_description") = ChkString(Request.Form("news_description"))	 
  rs("news_active") = ChkString(Request.Form("news_active"))	  
  rs("renews_id") = news_id  
  rs("log_by") = Request.Cookies("GAPS")("sloginid")
  rs("log_date") = ChkDateTimeMySQL(now())
  rs.Update 
  rs.Close    		 
  	
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblNews','addnews=" & ChkString(left(Request("news_title"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)
   
'----------------------------------------------------------------------------------------------------
  Case "editnews"

  sql = "SELECT news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
        "where news_id = " & request("news_id")
  rs.ActiveConnection = strconnect 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  if not rs.eof then  
  rs("news_date") = ChkString(Request.Form("news_date"))	  
  rs("news_grouplevel") = ChkString(Request.Form("news_grouplevel"))	 
  rs("news_title") = ChkString(Request.Form("news_title"))	 
  rs("news_type") = ChkString(Request.Form("news_type"))
  rs("news_description") = ChkString(Request.Form("news_description"))	 
  rs("news_active") = ChkString(Request.Form("news_active"))	  
  rs("log_by") = Request.Cookies("GAPS")("sloginid")
  rs("log_date") = ChkDateTimeMySQL(now())
  rs.Update 
  end if
  rs.Close    
  
'''update Lingerie
  sql = "SELECT news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
        "where renews_id = " & request("news_id")
  rs.ActiveConnection = strconnectLingerie 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  if not rs.eof then  
  rs("news_date") = ChkString(Request.Form("news_date"))	  
  rs("news_grouplevel") = ChkString(Request.Form("news_grouplevel"))	 
  rs("news_title") = ChkString(Request.Form("news_title"))	 
  rs("news_type") = ChkString(Request.Form("news_type"))
  rs("news_description") = ChkString(Request.Form("news_description"))	 
  rs("news_active") = ChkString(Request.Form("news_active"))	  
  rs("log_by") = Request.Cookies("GAPS")("sloginid")
  rs("log_date") = ChkDateTimeMySQL(now())
  rs.Update 
  end if
  rs.Close     
  
  set rs = nothing      
  url = "mis_news_view.asp?loginerr=News: " & request("news_id") & " has been updated."
	
 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblNews','editnews=" & ChkString(left(Request("news_title"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
 CUD(sql)

'----------------------------------------------------------------------------------------------------
  Case "deletenews"
    sql = "DELETE FROM tblNews WHERE news_id=" & Request("news_id") 
    Call CUD(sql)	

    ''''Update Lingerie
    sql = "DELETE FROM tblNews WHERE renews_id=" & Request("news_id") 
    Call CUDLingerie(sql)
	
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblNews','deletenews=" & Request("news_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)

	url = "mis_news_view.asp?loginerr=News: " & request("news_id") & " has been deleted."
'----------------------------------------------------------------------------------------------------

  Case "feedback"
  
	subject = "Gaps.com.my - " & gapstype & " feedback: " & Request.Cookies("GAPS")("sloginid")
	body = body & _
		 "User Name : " & Request.Cookies("GAPS")("sloginid") & "<br>" & _
		 "User Email : " & Request.Cookies("GAPS")("email") & "<br>" & _
		 "Feedback Type : " & request("feedbacktype") & "<br>" & _
		 "Feedback Enquiry : " & request("feedbackenquiry") & "<br>" 		
	sendemail "info@gaps.com.my" ,"alan@redantz.com" , subject ,body
	sendemail "info@gaps.com.my" ,"shyap@asiabrandscorp.com" , subject ,body		
	url = request("sourcepage") & "?loginerr=Thank you for your feedback."

'----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

  Case "addFaultyCode"  
   
	   ''''Add addFaultyCode	   
	sql = "SELECT top 1 fr_id, fr_code, fr_description, fr_type, fr_status FROM tblfaultyreason"
	set rs = server.CreateObject("adodb.recordset")
	rs.ActiveConnection = strconnect
	rs.Source = sql
	rs.CursorLocation  = 3
	rs.CursorType = 2
	rs.LockType = 2
	rs.Open   
	rs.AddNew  
	rs("fr_description") = request("fr_description")	
	rs("fr_type") = request("fr_type")
	rs("fr_status") = request("fr_status")						
	rs.Update 
	rs.Close     
	
	   		 
	sql = "select top 1 fr_id from tblfaultyreason order by fr_id desc "
	fr_id = selectid(sql)	
	fr_code = 1000 + fr_id
	fr_code = "F" & fr_code 
	sql = "update tblfaultyreason set fr_code = '" & fr_code & "' where fr_id = " & fr_id
	CUD(sql)	  
		  
    url = "mis_master_FaultyCode_view.asp?loginerr=Faulty Code has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblfaultyreason','addFaultyCode=" & Request("fr_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editFaultyCode"  
   
	   '''' editFaultyCode	   
	 sql = "SELECT  top 1 fr_id, fr_code, fr_description, fr_type, fr_status FROM tblfaultyreason  where fr_id = " & request("fr_id") & " "
	  set rs = server.CreateObject("adodb.recordset")
	  rs.ActiveConnection = strconnect 
	  rs.Source = sql
	  rs.LockType = 3	  
	  rs.Open   
	  if not rs.eof then
		rs("fr_description") = request("fr_description")	
		rs("fr_type") = request("fr_type")
		rs("fr_status") = request("fr_status")				
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
     url = "mis_master_FaultyCode_view.asp?loginerr=Faulty Code has been updated.&#orderdetail" 
	 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblfaultyreason','editFaultyCode=" & Request("fr_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delFaultyCode"   
   
  sql = "delete from tblfaultyreason where fr_id = " & request("fr_id") 	
  CUD(sql)  	 
  
  url = "mis_master_FaultyCode_view.asp?loginerr=Faulty Code has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblfaultyreason','delFaultyCode=" & Request("fr_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)

'----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

  Case "addFaultyRepair"  
   
	   ''''Add addFaultyRepair	   
	sql = "SELECT  top 1   ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction"
	set rs = server.CreateObject("adodb.recordset")
	rs.ActiveConnection = strconnect
	rs.Source = sql
	rs.CursorLocation  = 3
	rs.CursorType = 2
	rs.LockType = 2
	rs.Open   
	rs.AddNew  
	rs("ra_type") = request("ra_type")	
	rs("ra_category") = request("ra_category")
	rs("ra_repairaction") = request("ra_repairaction")						
	rs.Update 
	rs.Close     

    url = "mis_master_faultyreason_view.asp?loginerr=Faulty Reason has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_repairaction','addFaultyRepair=" & Request("ra_type") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editFaultyRepair"  
   
	   '''' editFaultyRepair	   
	 sql = "SELECT  top 1 ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction  where ra_id = " & request("ra_id") & " "
	  set rs = server.CreateObject("adodb.recordset")
	  rs.ActiveConnection = strconnect 
	  rs.Source = sql
	  rs.LockType = 3	  
	  rs.Open   
	  if not rs.eof then
		rs("ra_type") = request("ra_type")	
		rs("ra_category") = request("ra_category")
		rs("ra_repairaction") = request("ra_repairaction")				
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
     url = "mis_master_faultyreason_view.asp?loginerr=Faulty Reason has been updated.&#orderdetail" 
	 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_repairaction','editFaultyRepair=" & Request("ra_type") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delFaultyRepair"   
   
  sql = "delete from tbljob_repairaction where ra_id = " & request("ra_id") 	
  CUD(sql)  	 
  
  url = "mis_master_faultyreason_view.asp?loginerr=Faulty Reason has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblfaultyreason','delFaultyRepair=" & Request("ra_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)
 
   '----------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

  Case "addFaultyRepairjs"  
   
	   ''''Add addFaultyRepairjs	   
	sql = "SELECT  top 1  ra_repairactionjs FROM tbljob_repairactionjs"
	set rs = server.CreateObject("adodb.recordset")
	rs.ActiveConnection = strconnect
	rs.Source = sql
	rs.CursorLocation  = 3
	rs.CursorType = 2
	rs.LockType = 2
	rs.Open   
	rs.AddNew  
	rs("ra_repairactionjs") = request("ra_repairactionjs")						
	rs.Update 
	rs.Close     

    url = "mis_master_faultyreason_view_js.asp?loginerr=Faulty Reason (Job Sheet) has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tbljob_repairactionjs','addFaultyRepair=" & Request("ra_type") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
     '----------------------------------------------------------------------------------------------------
	  
  Case "addBrands"  
   
	   ''''Add addBrands	   
	  sql = "SELECT  top 1 brand_id, brand_name FROM tblbrand "
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2  
	  rs.Open   
	  rs.AddNew       
		rs("brand_name") = request("brand_name")	
	  rs.Update 
	  rs.Close        		 
		  
		  
    url = "mis_master_Brands_view.asp?loginerr=Brand has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblbrand','addBrands=" & Request("brand_name") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editBrands"  
   
	   '''' addBrands	   
	  sql = "SELECT top 1 brand_id, brand_name FROM tblbrand where brand_id = " & request("brand_id") & " " 
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2	  
	  rs.Open   
	  if not rs.eof then
		rs("brand_name") = request("brand_name")		
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
     url = "mis_master_Brands_view.asp?loginerr=Brand has been updated.&#orderdetail" 
	 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblbrand','editFaultyCode=" & Request("brand_name") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delBrands"   
   
  sql = "delete from tblbrand where brand_id = " & request("brand_id") 	
  CUD(sql)  	 
  
  url = "mis_master_Brands_view.asp?loginerr=Brand has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblbrand','delBrands=" & Request("brand_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)
  
'----------------------------------------------------------------------------------------------------
 '---------------------------------------------------------------------------------------------------

  Case "addPromoCode"  
   
	   ''''Add addPromoCode	   
	sql = "SELECT   top 1  promo_id, promo_code, promo_name, promo_from, promo_to, promo_status, log_by, log_date FROM tblpromotion "
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2  
	  rs.Open   
	  rs.AddNew       
		rs("promo_code") = request("promo_code")				
		rs("promo_name") = request("promo_name")
		rs("promo_status") = request("promo_status")						
	  rs.Update 
	  rs.Close        		 
		  
		  
    url = "mis_master_Promocode_view.asp?loginerr=Promo Code has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblpromotion','addPromoCode=" & Request("promo_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editPromoCode"  
   
	   '''' editFaultyCode	   
	  sql = "SELECT top 1    promo_id, promo_code, promo_name, promo_from, promo_to, promo_status, log_by, log_date FROM tblpromotion  where promo_id = " & request("promo_id") & " "
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2
	  rs.Open   
	  if not rs.eof then
		rs("promo_code") = request("promo_code")				
		rs("promo_name") = request("promo_name")
		rs("promo_status") = request("promo_status")				
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
     url = "mis_master_Promocode_view.asp?loginerr=Promo Code has been added.&#orderdetail" 
	 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblpromotion','editPromoCode=" & Request("promo_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delPromoCode"   
   
  sql = "delete from tblpromotion where promo_id = " & request("promo_id") 	
  CUD(sql)  	 
  
  url = "mis_master_Promocode_view.asp?loginerr=Promo Code has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblpromotion','delPromoCode=" & Request("promo_code") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)

  '----------------------------------------------------------------------------------------------------

  Case "addFaultyAction"  
   
	   ''''Add addFaultyAction	   
		sql = "SELECT top 1 ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction "
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2
		rs.Open
	    rs.AddNew       
		rs("ra_type") = request("ra_type")				
		rs("ra_category") = request("ra_category")	
		rs("ra_repairaction") = request("ra_repairaction")						
	    rs.Update 
	    rs.Close        		 
		  
  url = "mis_master_faultyaction_view.asp?loginerr=Faulty Action has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ra_repairaction','addFaultyAction=" & Request("ra_repairaction") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editFaultyAction"  
   
	   '''' editFaultyAction	   
	 sql = "SELECT top 1  ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction  where ra_id = " & request("ra_id") & " "
		set rs = server.CreateObject("adodb.recordset")
		rs.ActiveConnection = strconnect
		rs.Source = sql
		rs.CursorLocation  = 3
		rs.CursorType = 2
		rs.LockType = 2  
	  rs.Open   
	  if not rs.eof then
		rs("ra_type") = request("ra_type")				
		rs("ra_category") = request("ra_category")	
		rs("ra_repairaction") = request("ra_repairaction")				
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
  url = "mis_master_faultyaction_view.asp?loginerr=Faulty Action has been updated.&#orderdetail" 
	 
     sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ra_repairaction','editFaultyAction=" & Request("ra_repairaction") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delFaultyAction"   
   
  sql = "delete from tbljob_repairaction where ra_id = " & request("ra_id")  		
  CUD(sql)  	 
  
  url = "mis_master_faultyaction_view.asp?Faulty Action has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','ra_repairaction','DelFaultyAction=" & Request("ra_repairaction") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)

 '---------------------------------------------------------------------------------------------------
 '---------------------------------------------------------------------------------------------------

  Case "addcity"  
   
	  sql = "select top 1  state_name FROM tblstate where state_id='" & request("ct_state_id") & "'"
	  state_name = selectid(sql)
	  
	  ''''Add addcity	   
	  sql = "SELECT  top 1   ct_id, ct_cnty_id, ct_state_id, ct_state_code, ct_name FROM tblcity"
	  set rs = server.CreateObject("adodb.recordset")
	  rs.ActiveConnection = strconnect 
	  rs.Source = sql
	  rs.LockType = 3	  
	  rs.Open   
	  rs.AddNew       
		rs("ct_cnty_id") = "129"
		rs("ct_state_id") = request("ct_state_id")
		rs("ct_state_code") = state_name
		rs("ct_name") = request("ct_name")
	  rs.Update 
	  rs.Close        		 
		  
		  
    url = "mis_master_city_view.asp?loginerr=City has been added.&#orderdetail" 
		 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcity','addcity=" & Request("ct_name") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
 '----------------------------------------------------------------------------------------------------
  Case "editcity"  

	  sql = "select top 1  state_name FROM tblstate where state_id='" & request("ct_state_id") & "'"
	  state_name = selectid(sql)
	  
	  '''' editcity	   
	  sql = "SELECT  top 1   ct_id, ct_cnty_id, ct_state_id, ct_state_code, ct_name FROM tblcity where ct_id = " & request("ct_id") 
	  set rs = server.CreateObject("adodb.recordset")
	  rs.ActiveConnection = strconnect 
	  rs.Source = sql
	  rs.LockType = 3	  
	  rs.Open   
	  if not rs.eof then
		rs("ct_state_id") = request("ct_state_id")
		rs("ct_state_code") = state_name
		rs("ct_name") = request("ct_name")	
	  end if							
	  rs.Update 
	  rs.Close        		 
		  
     url = "mis_master_city_view.asp?loginerr=City has been updated.&#orderdetail" 
	 
    sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcity','editcity=" & Request("ct_name") & "','" & ChkDateTimeMySQL(now()) & "')"         
    CUD(sql)
 
  '----------------------------------------------------------------------------------------------------
  Case "delcity"   
   
  sql = "delete from tblcity where ct_id = " & request("ct_id") 
  CUD(sql)  	 
  
  url = "mis_master_city_view.asp?loginerr=City has been deleted.&#orderdetail" 
 
  sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
          Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblcity','delcity=" & Request("ct_id") & "','" & ChkDateTimeMySQL(now()) & "')"         
  CUD(sql)
 
'----------------------------------------------------------------------------------------------------
  Case "Settingusers"
  
  sql = "SELECT user_id, view_cost,verify_claim,approve_stk FROM tblusers where user_id = " & request("user_id")
  set rs = server.CreateObject("adodb.recordset")
  rs.ActiveConnection = strconnect 
  rs.Source = sql
  rs.LockType = 3	  
  rs.Open   
  if not rs.eof then
  rs("view_cost") = ChkString(Request.Form("view_cost"))
  rs("verify_claim") = ChkString(Request.Form("verify_claim"))
  rs("approve_stk") = ChkString(Request.Form("approve_stk"))
  rs.Update 
  end if
  rs.Close    
  
  set rs = nothing      
  url = "mis_user_edit.asp?type=editusers&user_id=" & request("user_id") & "&loginerr=User Setting: " & request("user_id") & " has been updated."
	
 sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
	  Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','tblusers','Settingusers=" & ChkString(left(Request("user_name"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
 CUD(sql)
   
'----------------------------------------------------------------------------------------------------

   
End Select
Response.Clear
Response.Redirect(url)
%>