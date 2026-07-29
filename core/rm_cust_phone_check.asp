<!-- #include file="database/datastore.asp" -->

<style>
.btn-select-customer {
    background: linear-gradient(135deg, #ff8a00, #ff6a00);
    color: #ffffff;
    border: none;
    padding: 8px 16px;
    border-radius: 20px;
    font-size: 13px;
    font-family: Arial, sans-serif;
    cursor: pointer;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    transition: all 0.2s ease;
}

.btn-select-customer:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 14px rgba(0,0,0,0.2);
    background: linear-gradient(135deg, #ff9a1a, #ff7a1a);
}

.btn-select-customer:active {
    transform: translateY(0);
    box-shadow: 0 3px 6px rgba(0,0,0,0.15);
}

.btn-select-customer:disabled{
 opacity:0.5;
 cursor:not-allowed;
}

 .titleblue1 {
        font-size: 18px; /* was probably 13px / 14px before */
    }


</style>


<%
'phone1=replace(request("phone1"), "-","")
phone1 = Replace(Replace(Request("phone1"), "-", ""), " ", "")
phone2 = Replace(Replace(Request("phone2"), "-", ""), " ", "")

postfound = 1 'flag to check multiple addressess/postcode if found

if phone1 = "" and phone2 = "" then%>
       <td class="titleblue1">
          <div align="left" style="color:#c00000;">
                <strong>Please enter at least 1 phoe number</strong>
           </div></td><%
    response.End
end if

i = 1

Dim phones
phones = ""

If phone1 <> "" Then
    phones = "'" & Replace(phone1,"'","") & "'"
End If

If phone2 <> "" Then
    If phones <> "" Then phones = phones & ","
    phones = phones & "'" & Replace(phone2,"'","") & "'"
End If

sql = "select job_code, job_date, job_cust_name, job_cust_address, job_cust_postcode, job_cust_tel1, job_cust_tel2," & _
      " job_createddate, job_createdby, job_tech_code, job_status " & _
      "from tbljob where " & _
      " replace(replace(replace(job_cust_tel1,'-',''),' ',''),'+','') in (" & phones & ")" & _
      " or replace(replace(replace(job_cust_tel2,'-',''),' ',''),'+','') in (" & phones & ")"
              
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   'norecord = "There is no record found."
               %>
        <td class="titleblue1">
        <div align="left" style="color:#c00000;">
            <strong>Jobs not found for this number !!</strong><br />
            <strong>Try searching by customer name</strong>
        </div></td>
        <%
    response.End
else

    sql2 = "SELECT DISTINCT job_cust_postcode " & _
       "FROM tbljob " & _
       "WHERE replace(replace(replace(job_cust_tel1,'-',''),' ',''),'+','') in (" & phones & ") " & _
       "or replace(replace(replace(job_cust_tel2,'-',''),' ',''),'+','') in (" & phones & ")"
   
    set rs2 = server.CreateObject("adodb.recordset")
    rs2.ActiveConnection = strconnect
    rs2.Source = sql2
    rs2.CursorLocation  = 3
    rs2.Open
           
    If rs2.RecordCount > 1 Then
        postfound=2
        'Response.Write "More than 1 postcode found."
    Else
        postfound=1
        'Response.Write "Only 1 postcode found."
    End If
end if

If Not rs.EOF Then
if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if

'fetch customer profile to load screen

sql = "SELECT top 1 tblcustomer.cust_id,  " & _
		"tblcustomer.cust_code, tblcustomer.cust_name, " & _
		"tblcustomer.cust_type, tblcustomer.cust_status, tblcustomer.cust_reg_no, tblcustomer.cust_company, tblcustomer.cust_address, tblcustomer.cust_postcode,  " & _
		"tblcustomer.cust_state, tblcustomer.cust_state_id, tblcustomer.cust_city, tblcustomer.cust_city_id,tblcustomer.cust_cnty_id, tblcustomer.cust_email,  " & _
		"tblcustomer.cust_tel1, tblcustomer.cust_tel2, tblcustomer.cust_fax, tblcustomer.cust_website, tblcustomer.cust_password, tblcustomer.cust_gstregno,  " & _
        "tblcustomer.cust_attention, " & _
		"tblonlinewarranty.warrantyno, tblonlinewarranty.serialno, tblonlinewarranty.productmodel, tblonlinewarranty.dealername, tblonlinewarranty.purchase_date, " & _
		"job_total_job = ( select count(job_id) from tbljob where tbljob.job_cust_code = tblcustomer.cust_code) " & _
		"FROM tblcustomer left join tblonlinewarranty on tblcustomer.cust_email=tblonlinewarranty.customeremail " & _
	    "WHERE replace(replace(replace(cust_tel1,'-',''),' ',''),'+','') in (" & phones & ") " & _
        "or replace(replace(replace(cust_tel2,'-',''),' ',''),'+','') in (" & phones & ")"

set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql
rs1.CursorLocation  = 3
rs1.Open
if rs1.eof then
     %>
        <td class="titleblue1">
        <div align="left" style="color:#c00000;">
            <strong>Invalid customer record for the number <%=phones%></strong><br />
            <strong>Contact the admin</strong><br />
        </div></td>
        <%
end if
		
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed
link = "&phone1=" & phone1 & "&phone2=" & phone2 
%>

<%
Function JsExpr(v)
    If IsNull(v) Then v = ""
    v = CStr(v)

    ' 1) Escape backslash first
    v = Replace(v, "\", "\\")

    ' 2) Escape double-quote properly for JS string:  \" 
    v = Replace(v, """", "\""")

    ' 3) Normalize line breaks
    v = Replace(v, vbCrLf, "\n")
    v = Replace(v, vbCr, "\n")
    v = Replace(v, vbLf, "\n")

    ' 4) Prevent breaking out of </script>
    v = Replace(v, "</", "<\/")

    ' Always return a valid JS string literal (quoted)
    JsExpr = """" & v & """"
End Function
%>

 
<script type="text/javascript">
    function selectCustomer() {
        var f = parent.opener.document.forms["formorder"];

        f.job_cust_code.value = <%=JsExpr(rs1("cust_code")) %>;
        parent.opener.document.getElementById("job_total_job").innerHTML =
        <%=JsExpr(rs1("job_total_job")) %>;

        f.job_cust_name.value = <%=JsExpr(rs1("cust_name")) %>;
        f.job_cust_email.value = <%=JsExpr(rs1("cust_email")) %>;
        f.job_cust_address.value = <%=JsExpr(rs1("cust_address")) %>;
        f.job_cust_state.value = <%=JsExpr(rs1("cust_state")) %>;
        f.job_cust_state_id.value = <%=JsExpr(rs1("cust_state_id")) %>;
        f.job_cust_postcode.value = <%=JsExpr(rs1("cust_postcode")) %>;
        f.job_cust_city_code.value = <%=JsExpr(rs1("cust_city_id")) %>;
        f.job_cust_city.value = <%=JsExpr(rs1("cust_city")) %>;
        f.job_cust_tel1.value = <%=JsExpr(rs1("cust_tel1")) %>;
        f.job_cust_tel2.value = <%=JsExpr(rs1("cust_tel2")) %>;
        f.job_cust_cnty_id.value = <%=JsExpr(rs1("cust_cnty_id")) %>;

        window.close();
    }
</script>

        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><strong>Job Sheet Record for Phone No - <%=phones%></strong></div></td>
                      </tr>
                      <tr>
                      <% If postfound = 2 Then %>
                        <td class="titleblue1">
                            <div align="left" style="color:#c00000;">
                                  <strong>Multiple postcode found !!</strong><br />
                                  <strong>Try searching by customer name to load customer record</strong>
                            </div>
                        </td>
                    <% End If %>

                      </tr>
                    </table></td>
                </tr>
                  </table>
                 </td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr><td></td>
               
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					'Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					'Response.Write " <a href='rm_cust_phone_check.asp?num=" & (j-1) * row & link & "&do_status=" & do_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cust_phone_check.asp?num=" & Showed+row & link & "&do_status=" & do_status & "'> Next >></a>"
      ' Response.Write "<a href='rm_cust_phone_check.asp?num=" & Showed+row & link & "> Next >></a>"
	End If
	
                    %>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                  
                  </table>
                        <div align="right">
                        <input type="button" class="btn-select-customer" value="Load Customer Record" onclick="selectCustomer();" <% If postfound = 2 Then Response.Write("disabled") End If %>>
</div>                 <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Job Code </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Job Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Name</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Cust Address</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Created <span> Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Created By</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Tech Code</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Status</strong></font></td>                     
                    </tr>
                    
<% 

if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                     <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" width="40"> <%=j%> </td>
                      <td width="120"> <%=rs("job_code")%></td>
                      <td nowrap="nowrap"> <%=chkdate(rs("job_date"))%></td>
                      <td> <%=rs("job_cust_name")%></td>
                      <td> <%=rs("job_cust_address")%>,<%=rs("job_cust_postcode")%></td>
                      <td> <%=chkdate(rs("job_createddate"))%></td>
                      <td> <%=rs("job_createdby")%></td>
                      <td> <%=rs("job_tech_code")%></td>
                      <td><%=rs("job_status")%></td>              
                    </tr>
<%
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
rs1.close
rs2.close
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
%>  
                      <td></td>
                   <td height="30" align="left" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					'Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_cust_phone_check.asp?num=" & (j-1) * row & link  & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cust_phone_check.asp?num=" & Showed+row & link  & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->