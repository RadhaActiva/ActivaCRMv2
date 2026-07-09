
<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", "filename=weekly_POR_Report_" & year(date()) & month(date()) & day(date()) & ".xls"
por_docno   = request("por_docno")
%>

<!-- #include file="database/datastore.asp" -->

<%
sql2 = "SELECT por_id,por_docno,por_date,por_remark,por_part_code,por_eta1,por_order_qty1,por_eta2,por_order_qty2,por_last_stockin,por_total_incoming,por_total_last" & _
",FORMAT(por_avg_3, 'N2') as 'por_avg_3',FORMAT(por_avg_6, 'N2') as 'por_avg_6',por_mth1,por_mth2,por_mth3,por_mth4,por_mth5,por_mth6,por_mth1_qty,por_mth2_qty,por_mth3_qty,por_mth4_qty" & _
",por_mth5_qty,por_mth6_qty,por_createdby,por_createddate,por_sw1_qty,por_ex_qty FROM tblpor where por_docno = '" & por_docno &"'"    

set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
i=1 

%>

<h1><u>Summary of Stock Movement on Weekly Basis</u></h1>


<table width="100%" border="0" cellpadding="4" cellspacing="0">
<tr><td width="20%"><strong>Date </strong></td><td align ="left"><%=FormatDateTime(date(),1)%></tr>
<tr><td><strong>For Week </strong></td><td align ="left"><%=DatePart("ww", Now())%></td></tr>
<tr><td><strong>POP Doc No </strong></td><td align ="left"><%=rs1("por_docno")%></td></tr>
<tr><td><strong>DOC Date </strong></td><td align ="left"><%=chkdate(rs1("por_date"))%></td></tr>
</table>
<br>
<br>
<table width="100%" border="0" cellpadding="4" cellspacing="0">
             <tr>
                  <td valign="top" bgcolor="#FFFFFF">
                    <table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></th>
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Part #</strong></font></th>   
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Description</strong></font></th>  
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>SW1<br/> Qty</strong></font></th> 
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Exchange<br/> Qty</strong></font></th> 
                      <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF">
                        
                          <table width="100%" border="1" cellpadding="2" cellspacing="0">
                              <tr><th colspan="2" align="center"><font color="#FFFFFF"><strong>1st Request</strong></th></tr>
                              <tr><th><font color="#FFFFFF"><strong>ETA/ETS</strong></th><th><font color="#FFFFFF"><strong>Order Qty</strong></th></tr>
                            </table></font></th>   
                      
                        <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF">
                       
                            <table width="100%" border="1" cellpadding="2" cellspacing="0">
                              <tr><th colspan="2" align="center"><font color="#FFFFFF"><strong>2nd Request</strong></th></tr>
                              <tr><th><font color="#FFFFFF"><strong>ETA/ETS</strong></th><th><font color="#FFFFFF"><strong>Order Qty</strong></th></tr>
                           </table></font></th>                              

                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total <br/>Incoming Stock</strong></font></th>   
                      <th align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock last<br/> for _Mth</strong></font></th>  
                      <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>3-Month <br/>Average</strong></font></th>  
                      <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>6-Month <br/>Average</strong></font></th>  
                      <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1">
                            <table width="100%" border="1" cellpadding="2" cellspacing="0">
                              <tr><th colspan="6" align="center"><font color="#FFFFFF"><strong>Monthly Usage</strong></font></th></tr>
                              <tr><th><font color="#FFFFFF"><%=rs1("por_mth1")%></font></th><th><font color="#FFFFFF"><%=rs1("por_mth2")%></font></th><th><font color="#FFFFFF"><%=rs1("por_mth3")%></font></th><th><font color="#FFFFFF"><%=rs1("por_mth4")%></font></th><th><font color="#FFFFFF"><%=rs1("por_mth5")%></font></th><th><font color="#FFFFFF"><%=rs1("por_mth6")%></font></th></tr>
                            </table></th> 
                       <th align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Remark</strong></font></th>  
                        </tr> 
                        
                   
  <%
  
while not rs1.eof
if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

     partdesc = ""
     sql = "select md_desc from tblmodel where md_code = '" & rs1("por_part_code") & "'  "
     partdesc = selectid(sql)     
%>
    <tr bgcolor="<%=nbgcolor%>">
                      <td height="25" align="center" valign="top" nowrap="nowrap"><%=i%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_part_code")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=partdesc%></td>
                        <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_sw1_qty")%></td>
                        <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_ex_qty")%></td>
                      <td>
                        <table width="100%" border="1" cellspacing="0">
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chkdate(rs1("por_eta1"))%></td>
				            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_order_qty1")%></td>
                        </table>        
                      </td>
                      <td>
                        <table width="100%" border="1" cellspacing="0">
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chkdate(rs1("por_eta2"))%></td>
				            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_order_qty2")%></td>
                        </table>        
                      </td>
				      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_total_incoming")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("por_total_last"))%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("por_avg_3"))%>&nbsp;</td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("por_avg_6"))%>&nbsp;</td>
           
                      <td>
                        <table width="100%" border="1" cellspacing="0">
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth1_qty")%></td>
				            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth2_qty")%></td>
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth3_qty")%></td>
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth4_qty")%></td>
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth5_qty")%></td>
                            <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_mth6_qty")%></td>
                        </table>      
                          <td align="left" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("por_remark")%>&nbsp;</td>
                      </td>
                    </tr>
  <%

count = count + 1 
i = i + 1
rs1.MoveNext
wend
rs1.Close
Set rs1 = Nothing

If Err.Number <> 0 Then
  Response.Write (Err.Description)   
  Response.End 
End If

%>
   </table>
 </table>
  

