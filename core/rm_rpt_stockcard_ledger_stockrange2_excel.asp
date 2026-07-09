<!-- #include file="database/datastore.asp" -->
<%  
'Response.ContentType = "application/vnd.ms-excel"
'Response.AddHeader "content-disposition", " filename=Stock_Ledger_" & searchvalue & year(date()) & month(date()) & day(date()) & ".xls"
searchvalue = request("searchvalue")
searchvalue2 = request("searchvalue2")
wh_code = request("wh_code")
whchk = request("whchk")
job_date_from = request("job_date_from")
job_date_to = request("job_date_to")

totalqty_in = 0
totalqty_out = 0
total_balance = 0

if whchk = "Yes" then
        wh_code = "All"
end if

 sql1 = "select md_desc from tblmodel where md_code = '" & searchvalue & "'"
 set rs1 = server.CreateObject("adodb.recordset")
 rs1.ActiveConnection = strconnect
 rs1.Source = sql1
 rs1.CursorLocation  = 3
 rs1.Open
 if not rs1.eof then
      prod_desc = rs1("md_desc")
 else
       prod_desc = ""
 end if 
  
%>
<!-- #include file="database/datastore.asp" -->
<h2>Stock Ledger Listing from <%=job_date_from%> To <%=Job_date_to%> </h2>
<table>
<tr><td><Strong>Item Code: </Strong></td><td nowrap="nowrap"><Strong><%=searchvalue%> - <%=searchvalue2%></Strong></td></tr>
<tr><td><Strong>Location: </Strong></td><td nowrap="nowrap"><Strong><%=wh_code%></Strong></td></tr>
</table>
<table width="100%" border="0" cellpadding="4" cellspacing="0">

   <tr>
                  <table border="0" cellpadding="4" cellspacing="0" bordercolor="#E8E8E8">
                    <tr valign="top" bgcolor="#88c0a7">
                      <td width="6%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">No</font></strong></td>
                      <td width="15%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Code</font></strong></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Date</strong></td>
                      <td width="20%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc Type</strong></td>
                      <td width="13%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc No</strong></td>
                      <td width="109" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Cust Code</strong></td>
                      <td width="32%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Description</strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty In</strong></td>          
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty Out</strong></td>
                      </tr>
 <%
 
i = 1  
sql4 = request.Cookies("GAPS")("sqlexcel")
set rs4 = server.CreateObject("adodb.recordset")
rs4.ActiveConnection = strconnect
rs4.Source = sql4
rs4.CursorLocation  = 3
rs4.Open
while not rs4.eof

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                  
                      <%
                     ' stk_qty_in=0
                     ' stk_qty_out=0
                     ' if rs4("stk_quantity") >= 0 then
                      '    stk_qty_in = rs4("stk_quantity")
                      'elseif  rs4("stk_quantity") < 0 then 
                       '   stk_qty_out = rs4("stk_quantity")
                      'elseif isnull(rs4("stk_quantity")) then
                       ' stk_qty_in=0
                        'stk_qty_out=0
                      'end if  

                    totalqty_in = totalqty_in + cint(stk_qty_in) 
                    totalqty_out = totalqty_out + cint(stk_qty_out)     
                      %>

                     <tr bgcolor="<%=nbgcolor%>">
                     <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=i%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs4("stk_code_id")%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=chkdate(rs4("stk_date"))%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_type")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_no")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_code")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_description")%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_in%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_out%></td>
                    </tr>
  <%

i = i + 1
rs4.MoveNext
wend
rs4.Close
Set rs4 = Nothing
%>
    <tr></tr>
      <tr bgcolor="#F3F3F3" valign="top">
                      <td height="40" colspan="6" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_in%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_out%></strong></td>                      
      </tr>
</table>
