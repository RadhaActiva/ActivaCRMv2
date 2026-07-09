<!-- #include file="database/datastore.asp" -->
<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Dealer_Collection" & searchvalue & year(date()) & month(date()) & day(date()) & ".xls"
Response.Write "<style>.num2{mso-number-format:""0.00"";}</style>"

dealer_code = request("dealer_code")
job_date_from = request("job_date_from")
job_date_to = request("job_date_to")

collection_amt = 0
labouramt = 0
t_totalpymt = 0
t_totalcoll = 0

sql3 = "SELECT cust_name, cust_address, cust_postcode, cust_city, cust_state " & _
       "FROM tblcustomer " & _
       "WHERE cust_code = '" & Replace(dealer_code, "'", "''") & "'"

Set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.Open sql3, strconnect, 1, 1   ' adOpenKeyset, adLockReadOnly
If Not rs3.EOF Then
    cust_name     = rs3("cust_name")
    cust_address  = rs3("cust_address")
    cust_postcode = rs3("cust_postcode")
    cust_city     = rs3("cust_city")
    cust_state    = rs3("cust_state")
End If
rs3.Close
Set rs3 = Nothing
%>
<!-- #include file="database/datastore.asp" -->
<h4>Installation Job Submission Form (Collection Report)</h4>
<table>
<tr><td><Strong>Date Period: </Strong></td><td nowrap="nowrap"><%=job_date_from%> to <%=job_date_to%></td></tr>
<tr><td><Strong>Dealer Name: </Strong></td><td nowrap="nowrap"><%=dealer_code%> - <%=cust_name%></td></tr>
<tr><td><Strong>Address:</Strong></td><td nowrap="nowrap"><%=cust_address%> - <%=cust_postcode%></td></tr>
<tr><td></td><td nowrap="nowrap"><%= cust_city%> , <%=cust_state%></td></tr>
<tr><td></td></tr>
</table>
<table width="100%" border="0" cellpadding="4" cellspacing="0">

   <tr>
                  <table border="0" cellpadding="4" cellspacing="0" bordercolor="#E8E8E8">
                      <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                        <td width="50" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Job Sheet No</span></strong></font></td>
                        <td width="50" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Dealer DO/Invoice No</span></strong></font></td>
                        <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Model Desc</strong></span></font></td>
                        <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Serial No</strong></span></font></td>
                      <td width="109" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Warranty <br />Info</span></strong></font></td>
                        <td width="109" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Type</span></strong></font></td>
                        <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Collection<br />(MYR)</span></strong></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total <br /> Payment</span></strong></font></td>
                    </tr>
 <%
 
i = 1  
sql1 = request.Cookies("GAPS")("sqlexcel")
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>             
                      <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=i%></td>
                        <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("job_code")%></font></strong></td>
                             <td align="left" nowrap="nowrap"> <%=rs1("job_dealer_inv")%></td>
                             <td align="left" nowrap="nowrap"> <%=rs1("job_model_desc")%></td>
                             <td align="left" nowrap="nowrap"> <%=rs1("job_SN_no")%></td>
                             <td align="left" nowrap="nowrap"><%=rs1("job_actual_wrty_status")%></td>
                             <td align="left" nowrap="nowrap"><%=rs1("job_type")%> </td>
                        <%
                            labouramt = ChkNumber(rs1("job_totallabourAmt"))

                            If Not IsNull(rs1("job_type")) Then
                                If rs1("job_type") = "WH" Then
                                    collection_amt = labouramt - 20
                                ElseIf rs1("job_type") = "CF" Then
                                    collection_amt =labouramt - 10
                                End If
                                if collection_amt < 0 then
                                    collection_amt = 0
                                End If
                            End If
                            %>
                       <td class="num2" align="right" nowrap="nowrap" >
                            <%= CDbl(Replace(collection_amt & "", ",", "")) %>
                        </td>
                        <td class="num2" align="right" nowrap="nowrap">
                            <%= CDbl(Replace(rs1("job_totalAmt") & "", ",", "")) %>
                        </td>
                    </tr>
  <%

t_totalcoll = collection_amt + t_totalcoll 
t_totalpymt = rs1("job_totalAmt") + t_totalpymt 

i = i + 1
rs1.MoveNext
wend
rs1.Close
Set rs1 = Nothing
%>
    <tr></tr>
      <tr bgcolor="#F3F3F3">
                  <td height="40" colspan="7" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                  <td class="num2" align="right" nowrap="nowrap" bgcolor="#999999">
                    <strong><%= CDbl(Replace(t_totalcoll & "", ",", "")) %></strong>
                   </td>

                    <td class="num2" align="right" nowrap="nowrap" bgcolor="#999999">
                    <strong><%= CDbl(Replace(t_totalpymt & "", ",", "")) %></strong>
                    </td>
                    </tr>

                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#999999"><strong>Total Job</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=i-1%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                     </tr>
      </tr>
</table>
