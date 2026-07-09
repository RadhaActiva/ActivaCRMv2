<!-- #include file="header.asp" -->
<%

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

%> 
 <script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_pnl.asp?post=yes";
	document.form1.submit();
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Monthly </font>Inventory Assessment Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_pnl_excel.asp?job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0"/></a></td>
                    </tr>

                  </table></td>
                </tr>
                <tr>
                  <form id="form1" name="form1" method="post" action="rm_rpt_pnl.asp?post=yes">
                  <td valign="top" bgcolor="#FFFFFF">
                    Month Ending&nbsp;&nbsp;<input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>
                   <input type="submit" name="button" id="button3" value="Generate Report" onclick="javascript:DisplayReport();"/>&nbsp; &nbsp; Choose last day of the month ie (31st Aug 2022). Latest Avg Cost will be used as per this date.
                  </td>
                  <tr><td width="104" height="10" align="right"></td></tr>
                 </form>
                </tr>
                    
<% 
               
if Request.QueryString("post") = "yes" then
    'Delete all data
    sql9 = "delete from tblrpr_new_pnl where pnl_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'"
    set rs9 = server.CreateObject("adodb.recordset")
    rs9.ActiveConnection = strconnect
    rs9.Source = sql9
    rs9.Open

jobmonth = month(job_date_to)
jobyear = year(job_date_to)

sql1 = "SELECT sum(tblinvoice.inv_totalPartsAmt) FROM tblinvoice where tblinvoice.inv_id is not null and " & _
	   "month(tblinvoice.inv_date) =" & jobmonth & " and year(tblinvoice.inv_date) =" & jobyear & " and tblinvoice.inv_status = 'Posted'"
total_sales_parts  = selectid(sql1)

'update job_posteddate with job_donedate

sql2= "select sum(t.stotal) from " & _
    "(select stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= b.job_posteddate order by md_date desc) as stotal from tblstocktran a " & _
    "INNER JOIN TBLJOB b ON a.stk_voucherno=b.job_code " & _
	"where month(a.stk_date) = " & jobmonth & " and year(a.stk_date) = " & jobyear & " and a.STK_TYPE = 'Job' and b.job_actual_wrty_status='Over' and b.job_status='Posted' and b.job_inv_no is not NULL)T "

spare_parts_over_wrty = selectid(sql2) 'this is same as finance report 2 & 3

sql14= "select sum(t.stotal) from " & _
    "(select stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= b.job_posteddate order by md_date desc) as stotal from tblstocktran a " & _
    "INNER JOIN TBLJOB b ON a.stk_voucherno=b.job_code " & _
	"where month(a.stk_date) = " & jobmonth & " and year(a.stk_date) = " & jobyear & " and a.STK_TYPE = 'Job' and b.job_actual_wrty_status='Over' and b.job_status='Posted' and b.job_inv_no is NULL)T "

spare_parts_over_wrty_noinv = selectid(sql14)

'sql3= "select sum(t.stotal) from " & _
'	  "(select a.stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' order by md_date desc) as stotal from tblstocktran a " & _
'      "INNER JOIN TBLJOB b ON a.stk_voucherno=b.job_code " & _
'	  "where month(a.stk_date) = " & jobmonth & " and year(a.stk_date) = " & jobyear & "  and a.STK_TYPE = 'Job' and b.job_actual_wrty_status='Under' and b.job_status='Posted')t "

sql3= "select sum(t.stotal) from " & _
      "(select a.stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <=  b.job_posteddate order by md_date desc) as stotal from tblstocktran a " & _
	  "INNER JOIN TBLJOB b ON a.stk_voucherno=b.job_code " & _
	  "where month(a.stk_date) = " & jobmonth & " and year(a.stk_date) = " & jobyear & "  and a.STK_TYPE = 'Job' and b.job_actual_wrty_status='Under' and b.job_status='Posted')t "

spare_parts_under_wrty = selectid(sql3)

'for DO, no valid DO_Posted Data, hence use do_date , which is same as stk_date
sql15=  "select sum(t.stotal) from  " & _
	"(select stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <=  b.do_date order by md_date desc) as stotal from tblstocktran a " & _
	"inner join tbldo b  on b.do_no=a.stk_voucherno " & _
	"where month(a.stk_date) = " & jobmonth & " and year(a.stk_date) = " & jobyear & " " & _
	"AND a.STK_TYPE = 'DO' and a.stk_itm_code not in ('Labour', 'Service')  and b.do_status ='Posted')t "

spare_parts_do = selectid(sql15)

' need to optimize
sql16= 	"select round(sum(t.stotal),2) from " & _
		"(Select stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' order by md_date desc) as stotal " & _ 
		"from tblstocktran b where month(b.stk_date)  = " & jobmonth & " and year(b.stk_date) = " & jobyear & " and b.STK_TYPE = 'Job' " & _
        "and not exists (Select job_code from tbljob where tbljob.job_code = b.stk_voucherno))t "
        '"and not exists (Select job_id from tbljob where year(job_date) >= 2022))t "

spare_parts_err = selectid(sql16)

sql4="select CONVERT(varchar,dateadd(d,-(day('" & job_date_to & "')),'" & job_date_to & "'),112)"
lastmthdate = selectid(sql4) 'gives the last date of the previous month
'this was used in sql5 query at both variable places

sql41= "SELECT cast(DATEADD(mm, DATEDIFF(mm, 0, '" & job_date_to & "'), 0) as date)"
firstdayofthemth = selectid(sql41) 'this is not used at the moment

'after speaking to victor, opening balance should use last day of of the prev month plus avg cost as at that time --10/05/2023
'the opening fig of the new month is the same as the closing fig of the prev month based on the Inventory by Product Group Report

sql5="select round(sum(t.avgcost),2) from " & _
    "(select tblstocktran.stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = tblstocktran.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & lastmthdate & "' order by md_date desc) AS avgcost " & _
   "from tblstocktran  " & _
   "inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
   "where tblstocktran.stk_itm_code NOT IN ('Service','Labour') " & _
   "and tblstocktran.stk_reference is not null and cast(tblstocktran.stk_date as date) <= '" & lastmthdate & "' ) t"
opening_stock = selectid(sql5)
 
'04/06/23 - not using avg cost but PO purchase price as per Stock-In Screen.
sql6= "SELECT sum(st_totalaAmt) " & _
	"FROM tblstockin where st_id is not null and month(st_date) = " & jobmonth & " and year(st_date) = " & jobyear & " and st_status = 'Approved'" 
total_stock_purchases = selectid(sql6)
   
sql12 = "select sum(t.total_stockOut) from " & _
"(select (stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code=tblstocktran.stk_itm_code " & _
"and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' order by md_date desc)) as total_stockOut from tblstocktran where stk_type='Stock-Out' and month(stk_date) = " & jobmonth & " and year(stk_date) = " & jobyear & ")t "
 total_stock_out = selectid (sql12)

 sql13 = "select sum(t.total_stock_adjust) from " & _
"(select (stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code=tblstocktran.stk_itm_code " & _
"and CAST(tblmodel_avgcost.md_date as date) <=  '" & job_date_to & "' order by md_date desc)) as total_stock_adjust from tblstocktran where stk_type = 'Stock-Adj'and month(stk_date) =" & jobmonth & " and year(stk_date) =  " & jobyear & ")t "
 total_stock_adjust = selectid(sql13)

sql17="select sum(t.stotal) from " & _
		"(select stk_qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' order by md_date desc) as stotal from tblstocktran a " & _
		"inner join tblcn b  on b.cn_no=a.stk_voucherno " & _
		"where MONTH(a.stk_date) = " & jobmonth & " and YEAR(a.stk_date) = " & jobyear & "" & _
		"AND a.STK_TYPE = 'CN' and a.stk_itm_code not in ('Labour', 'Service')  and b.cn_status ='Posted')t"
stock_cn= selectid(sql17)

    if isnull(total_stock_purchases) then 
        total_stock_purchases=0
    end if

     if isnull(stock_cn) then 
        stock_cn=0
    end if

    if isnull(total_stock_out) then 
        total_stock_out=0
    end if

    if isnull(total_stock_adjust) then 
        total_stock_adjust=0
    end if

     if isnull(spare_parts_err) then 
        spare_parts_err=0
    end if

    if isnull(spare_parts_do) then 
        spare_parts_do=0
    end if

    if isnull(spare_parts_over_wrty_noinv) then 
        spare_parts_over_wrty_noinv=0
    end if

    if isnull(spare_parts_under_wrty) then 
        spare_parts_under_wrty_noinv=0
    end if

    if isnull(spare_parts_over_wrty) then 
        spare_parts_over_wrty=0
    end if

    if isnull(spare_parts_under_wrty) then 
        spare_parts_under_wrty=0
    end if
    

 sql6= "select sum(t.avgcost) from (select tblstocktran.stk_qty *  (select TOP 1 md_averagecost from tblmodel_avgcost where md_code=tblstocktran.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' " & _
       "order by md_date desc) AS avgcost from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
       "where tblstocktran.stk_itm_code NOT IN ('Service','Labour') " & _
       "and cast(tblstocktran.stk_date as date) <= '" & job_date_to & "' and stk_reference is not null) t" 
FTCRM_closing_stock = selectid(sql6)

total_spare_cost = spare_parts_over_wrty + spare_parts_under_wrty + spare_parts_over_wrty_noinv + spare_parts_do + spare_parts_err
gross_profit = total_sales_parts+total_spare_cost

if not isnull(gross_profit) and not isnull(total_sales_parts) and total_sales_parts > 0 then
    percentage_profit = (gross_profit/total_sales_parts) * 100
end if

total_spare_parts_stock = opening_stock + total_stock_purchases - abs(total_stock_out) + total_stock_adjust + stock_cn
closing_stock_figure = total_spare_parts_stock + spare_parts_over_wrty + spare_parts_under_wrty + spare_parts_over_wrty_noinv + spare_parts_do + spare_parts_err

  if isnull(closing_stock_figure) then 
        closing_stock_figure=0
  end if

  if isnull(FTCRM_closing_stock) then 
        FTCRM_closing_stock=0
  end if

  if isnull(percentage_profit) then 
        percentage_profit=0
  end if

stock_value_diff = round(closing_stock_figure,2) - round(FTCRM_closing_stock,2) 

FTCRM_closing_stock = round(FTCRM_closing_stock,2)
percentage_profit = round(percentage_profit,2)
spare_parts_over_wrty = round(spare_parts_over_wrty,2)

    if isnull(total_stock_purchases) then 
        total_stock_purchases=0
    end if

    if isnull(total_spare_parts_stock) then 
        total_spare_parts_stock=0
    end if

    if isnull(stock_value_diff) then 
        stock_value_diff=0
    end if

    if isnull(total_spare_cost) then 
        total_spare_cost=0
    end if
    
    if isnull(gross_profit) then 
        gross_profit=0
    end if

     if isnull(opening_stock) then 
        opening_stock=0
    end if
    
    
    'update into db for excel report display
    sql8="insert into tblrpr_new_pnl (pnl_date_ending, pnl_total_sales_parts,pnl_spare_parts_over_wrty,pnl_spare_parts_under_wrty,pnl_total_spare_cost,pnl_gross_profit,pnl_percentage_profit,pnl_opening_stock,pnl_total_stock_purchases,pnl_total_spare_parts_stock,pnl_closing_stock_figure,pnl_FTCRM_closing_stock,pnl_stock_value_diff, pnl_logby,pnl_total_stock_out,pnl_total_stock_adjust,pnl_credit_notes,pnl_spare_parts_over_wrty_notinv,pnl_spare_parts_do,pnl_error_job) " & _
    "values('" & job_date_to & "','" & total_sales_parts & "','" & spare_parts_over_wrty & "','" & spare_parts_under_wrty & "','" & total_spare_cost & "','" & gross_profit & "','" & percentage_profit & "','" & opening_stock & "','" & total_stock_purchases & "','" & total_spare_parts_stock & "','" & closing_stock_figure & "','" & FTCRM_closing_stock & "','" & stock_value_diff & "','" & Request.Cookies("GAPS")("sloginid") & "','" & total_stock_out & "','" & total_stock_adjust & "','" & stock_cn & "','" & spare_parts_over_wrty_noinv & "','" & spare_parts_do & "','" & spare_parts_err & "')"
      
    set rs8 = server.CreateObject("adodb.recordset")
    rs8.ActiveConnection = strconnect
    rs8.Source = sql8
    rs8.CursorLocation  = 3
    rs8.Open
end if
%>
              
<% if Request.QueryString("post") = "yes" and total_sales_parts > 0 then 'doesn't allow future reports%> 
    <tr><td width="104" height="15" align="right"></td></tr>
     <td width="17" align="center" bgcolor="#666666"><font color="#FFFFFF"><strong><nowrap>MONTHLY INVENTORY ASSESSMENT ACCOUNT FOR THE MONTH ENDING <%=job_date_to%> </nowrap></strong></font></td>
          <table width="60%" border="0" align="center" cellpadding="0" cellspacing="0">

                        <br>
                        <br>
                        <tr><td></td><td></td><td><b>MYR</b></td></tr>
                        <tr>
                        <td width="60%">Sales Spare Parts</td>
                        <td width="20%"></td>
                        <td><%=formatnumber(total_sales_parts)%></td>                            
                        <td></td>
                        </tr>
                        <P></P>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="20%"><U>LESS : Cost of Sales</U></td></tr>
                        <tr><td width="20%"></td>
                        <td></td>
                        <td></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="20%">Total Spare Part Costs - As Below</td>
                        <td></td>
                        <td><%=formatnumber(total_spare_cost,,,-1)%></td></tr>
                        <tr>
                        <td></td>
                        <td></td>
                        <td></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr></tr>
                        <tr><td valign="top"> <strong><font color="#0000FF"><b>Gross Profit</b></font></strong></td><td></td>
                        <td><font color="#0000FF"><u><%=formatnumber(gross_profit)%></u></font></td></tr>
                        <tr><td valign="top">% Of Gross Profit</td><td></td>
                        <td><%=chknumber2(percentage_profit)%>%</td></tr>
                        <tr></tr>
                        <tr><td width="104" height="10" align="right" class="style1"></td></tr>
                        <tr><td width="60%"><strong>Summary Stock</strong></td></tr>
                        <tr><td></td><td></td><td></td><td><b>MYR</b></td></tr>
                        <tr><td>Opening Stock As at - <%=lastmthdate%></td>
                        <td></td><td></td><td><%=formatnumber(opening_stock)%></td>
                        <tr><td>ADD : Stock Purchases</td>
                        <td></td><td></td><td><%=formatnumber(total_stock_purchases)%></td>
                        <tr><td>ADD : Credit Notes</td>
                        <td></td><td></td><td><%=formatnumber(stock_cn)%></td>
                        
                        <!--<tr><td width="104" height="10" align="right"></td></tr>-->

                        <tr><td width="60%">LESS : Stock-Out</td>
                        <td></td><td></td><td><u><%=formatnumber(total_stock_out,,,-1)%></u></td></tr>
                        <tr><td>(+/-): Stock Adjustments</td>
                        <td></td><td></td><td><%=formatnumber(total_stock_adjust)%></td>
                        <tr><td width="104" height="10" align="right"></td></tr>

                        <tr><td width="60%">Total Spare Parts Stock</td>
                        <td></td><td></td><td><u><%=formatnumber(total_spare_parts_stock)%></u></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="60%">LESS: Cost of Sales - Spare Parts O/Warranty</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_over_wrty,,,-1)%></td></tr>
                        <tr><td width="60%">LESS: Cost of Sales - Spare Parts O/Warranty (Not Inv)</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_over_wrty_noinv,,,-1)%></td></tr>
                        <tr><td width="60%">LESS: Spare Part Costs (W/H & C/F)-U/Warranty</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_under_wrty,,,-1)%></td></tr>
                        <tr><td width="60%">LESS: DO Spare Part Costs</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_do,,,-1)%></td></tr>
                        <tr><td width="60%">LESS: MISC (Error Jobs)</td>
                        <td></td><td></td><td><%=formatnumber(spare_parts_err,,,-1)%></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
              
                        <tr><td width="60%">Closing Stock <%=job_date_to%></td>
                        <td></td><td></td><td><%=formatnumber(closing_stock_figure)%></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="60%">As per FTCRM closing stock <%=job_date_to%> figure shows </td>
                        <td></td><td></td><td><u><%=formatnumber(FTCRM_closing_stock)%></u></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <%if stock_value_diff < 0 then %>
                        <tr><td width="60%"><strong><font color="#FF0000"><b>STOCK VALUE DIFFERENCES</b></font></strong></td>
                        <td></td><td></td><td><font color="#FF0000"><b><u><%=formatnumber(stock_value_diff,,,-1)%></u></b></font></td></tr>
                        <%else%>
                            <tr><td width="60%"><strong><font color="#0000FF"><b>STOCK VALUE DIFFERENCES</b></font></strong></td>
                            <td></td><td></td><td><font color="#0000FF"><b><u><%=formatnumber(stock_value_diff,,,-1)%></u></b></font></td></tr>
                        <%end if%>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                        <tr><td width="104" height="10" align="right"></td></tr>
                 </table>
            <% end if%>
                
<!-- #include file="footer.asp" -->