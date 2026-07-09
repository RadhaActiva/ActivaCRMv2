<!-- #include file="database/datastore.asp" -->
<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Stock_Ledger_" & searchvalue & year(date()) & month(date()) & day(date()) & ".xls"
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

  
%>
<!-- #include file="database/datastore.asp" -->
<h2>Stock Ledger Listing from <%=job_date_from%> To <%=Job_date_to%> </h2>
<table>
<tr><td><Strong>Item Code: </Strong></td><td nowrap="nowrap"><Strong><%=searchvalue%> - <%=searchvalue2%></Strong></td></tr>
<tr><td><Strong>Location: </Strong></td><td nowrap="nowrap"><Strong><%=wh_code%></Strong></td></tr>
</table>
<table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr valign="top" bgcolor="#88c0a7">
                      <td width="6%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">No</font></strong></td>
                      <td width="15%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Code</font></strong></td>
                      <td width="85%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Desc</font></strong></td>
                      <td width="10%" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Quantity</strong></td> 
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
                      partdesc = ""
                      sql = "select md_desc from tblmodel where md_code = '" & rs4("stk_itm_code") & "'  "
                      partdesc = selectid(sql)   
                      %>

                     <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=i%></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs4("stk_itm_code")%></td>                       
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=partdesc%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs4("stk_qty")%></strong></td>
                    </tr>
<%
    totalqty = totalqty + ChkNumberInt(rs4("stk_qty"))

i = i + 1
rs4.MoveNext
wend
rs4.Close
Set rs4 = Nothing
%>
    <tr></tr>
               <tr bgcolor="#F3F3F3">
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                      <td height="40"  align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty%></strong></td>
                    </tr>
</table>
