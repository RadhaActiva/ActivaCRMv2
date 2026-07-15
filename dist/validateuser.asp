<!-- #include file="core/database/dbconnect.asp" -->
<%
if Request("type") = "login" then 

if Request.Cookies("GAPS")("logincount") = "" then
   Response.Cookies("GAPS")("logincount") = 1    
end if

sql = "SELECT user_id,user_name, user_type, email, accesslevel, lastlogindate, lastchangepassword, lastloginIP, log_date, OutletID, OutletArea, OutletRegion, Outletselection, SupplierID,TransID, job_tech_code, view_cost, verify_claim, approve_stk  FROM tblusers WHERE user_name = '" & ChkStringLogin(Request("txtID")) & "' AND " & _
"password = '" & LenEncrypt(ChkString(Request.Form("txtID")),ChkString(Request("txtPASSWORD")))  & "' AND " & _
"user_active = 'Y'"     

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, strconnect,0,1,&H0001
If Not rs.EOF Then
   Response.Cookies("GAPS")("sloginid") = rs("user_name")
   Response.Cookies("GAPS")("slevel") = rs("accesslevel")
   Response.Cookies("GAPS")("email") = rs("email")   
   Response.Cookies("GAPS")("user_type") = rs("user_type")
   Response.Cookies("GAPS")("job_tech_code") = rs("job_tech_code")
   Response.Cookies("GAPS")("view_cost") = rs("view_cost")
   Response.Cookies("GAPS")("verify_claim") = rs("verify_claim")
   Response.Cookies("GAPS")("approve_stk") = rs("approve_stk")

   if rs("lastlogindate") <> "" then
      Response.Cookies("GAPS")("lastlogindate") = ChkDate(rs("lastlogindate"))  
   else
      Response.Cookies("GAPS")("lastlogindate") = ChkDate(date())  
   end if
   
   Response.Cookies("GAPS")("loginstatus") = "yes"
   Response.Cookies("GAPS")("forcechangepassword") = "no"
   Response.Cookies("GAPS")("logincount") = 0      
      
   sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
         Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','users','Login=" & ChkString(left(Request("txtID"),200)) & "','" & ChkDateTimeMySQL(now()) & "')"         
   CUD(sql)

   if rs("user_type") = "Riegen" then
      ' The CRM dashboard is the common landing page for every non-technician.
      loginAccessLevel = ""
      if not IsNull(rs("accesslevel")) then
         loginAccessLevel = LCase(Trim(CStr(rs("accesslevel"))))
      end if
      loginUserName = ""
      if not IsNull(rs("user_name")) then
         loginUserName = LCase(Trim(CStr(rs("user_name"))))
      end if
      if loginUserName = "admin" then
         Response.Redirect("core/mis_home.asp")
      elseif loginAccessLevel = "technician" or loginAccessLevel = "technician2" then
         Response.Redirect("core/rmtech_jobsheet_view.asp?job_status=Accepted")
      else
         Response.Redirect("core/rm_home.asp")
      end if
   end if  
else
   if Request.Cookies("GAPS")("logintry") = Request.Form("txtID") then             
      if cint(Request.Cookies("GAPS")("logincount")) >= 3 then
         Response.Redirect("login.asp?err=1")
      else
         Response.Cookies("GAPS")("logincount") = Request.Cookies("GAPS")("logincount") + 1
         response.Clear()
         Response.Redirect("login.asp?err=1")                  
      end if
   else       
      Response.Cookies("GAPS")("logintry") = Request.Form("txtID")
      Response.Cookies("GAPS")("logincount") = 0
      response.Clear()
      Response.Redirect("login.asp?err=1")
   end if     
end if
rs.close
Set rs = Nothing 
end if

if Request("type") = "logout" then 
   sql = "Update tblusers set lastlogindate = '" & ChkDateTimeMySQL(now()) & "' where user_name = '" & Request.Cookies("GAPS")("sloginid") & "'"
   CUD(sql)

   sql = "INSERT INTO tbluserlog (user_name,user_ip,tablename,tableaction,logdatetime) values ('" & _ 
         Request.Cookies("GAPS")("sloginid") & "','" & Request.servervariables("remote_addr") & "','users','Logout','" & ChkDateTimeMySQL(now()) & "')"         
   CUD(sql)

   Response.Cookies("GAPS")("sloginid") = ""
   Response.Cookies("GAPS")("slevel") = ""
   Response.Cookies("GAPS")("OutletID") = ""
   Response.Cookies("GAPS")("loginstatus") = ""
   Response.Cookies("GAPS")("view_cost") = ""
   Response.Cookies("GAPS")("verify_claim") = ""
   Response.Cookies("GAPS")("approve_stk") = ""
   Response.Redirect("login.asp?err=3")
end if

if request("type") = "forgetpassword" then         
   sql = "SELECT password FROM tblusers WHERE user_name='"& ChkString(Trim(Request("username"))) &"'"     
   spassword = selectid(sql) 
   spassword = LenDecrypt(Request("username"),spassword)
      
   sql = "select email from tblusers WHERE user_name = '" & ChkString(request("username")) & "'"
   email = selectid(sql)    
  
   if spassword <> "" then   
      subject = "RM - Your Password"
      body = "Hi " & ChkString(Trim(Request("username"))) & ",<br><br>" & _
             "Your authentication credentials are:<br><br>" & _          
             "Login ID : " & ChkString(Trim(Request("username"))) & "<br>" & _
             "Password : " & spassword & "<br><br>" & _
             "Please save this email and keep it securely. " & _           
             "Regards<br><br>" & _
             "Admin"      

      sendemail "cwchu@riegen.com.my" ,email , subject ,body
      Response.Redirect "forgetpassword.asp?loginerr=Your Username and Password has been sent" 
   else
      Response.Redirect "forgetpassword.asp?loginerr=Invalid Username."
   end if
end if
%>
