<!-- #include file="header.asp" -->
<%
if request("searchvalue") <> "" then
   searchvalue = request("searchvalue")
end if
    
Server.ScriptTimeout=600
%>

 <script>
function DisplayReport() 
{
	document.form1.action = "rm_adjust_cost.asp?post=yes";
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
                        <td class="titleblue1"><div align="left">Weighted Cost Adjustment</div></td>
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
                      <td width="20%" align="center" class="titlegrey1">&nbsp;</td>
                    </tr>

                  </table></td>
                </tr>
                <tr>
                  <form id="form1" name="form1" method="post" action="rm_rpt_pnl.asp?post=yes">
                  <td valign="top" bgcolor="#FFFFFF">
                      Stock-Code&nbsp;<input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" size="15" />&nbsp;
                      <%if request.Cookies("GAPS")("sloginid") = "" then %>
                            <input type="submit" name="button" id="button3" value="Recalculate Weighted Avg Cost" onclick="javascript:DisplayReport();"/>
                      <%end if%>
                  </td>
                    <tr><td width="104" height="10" align="right"></td></tr>
                   <tr><td colspan="3"  height="10" align="left" >Leaving the Stock-Code blank will recalibrate all the weighted avg cost starting from 1st Jan 2023</td></tr>
                   <tr><td colspan="3"  height="10" align="left" ><strong>**** Do NOT interrupt or close this page while it's being processed. ****</strong></td></tr>
                  <tr><td width="104" height="10" align="right"></td></tr>
                 </form>
                </tr>
                    
<%
    if Request.QueryString("post") = "yes" then
    
    if searchvalue = "" then 'PERFORM ADJUSTMENT FOR ALL OR A SINGLE PART CODE ONLY
       sql6= "select md_code, md_costprice, (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = tblmodel.md_code and md_entry = 'Y' " & _ 
       " order by md_date desc) as md_averageecost,(select TOP 1 md_date from tblmodel_avgcost where md_code = tblmodel.md_code and md_entry = 'Y'  " & _
       " order by md_date desc) as md_date from tblmodel where md_code NOT IN ('Service', 'Labour')"
    else
        sql6= "select md_code, md_costprice, (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = tblmodel.md_code and md_entry = 'Y' " & _ 
        " order by md_date desc) as md_averageecost,(select TOP 1 md_date from tblmodel_avgcost where md_code = tblmodel.md_code and md_entry = 'Y'  " & _
        " order by md_date desc) as md_date from tblmodel where md_code = '" & searchvalue &"' and md_code NOT IN ('Service', 'Labour')"
    end if

    set rs6 = server.CreateObject("adodb.recordset")
    rs6.ActiveConnection = strconnect
    rs6.Source = sql6
    rs6.CursorLocation  = 3
    rs6.Open
   
  while Not rs6.EOF
    weighted_avg_cost = 0
    average_cost = 0
    searchvalue = rs6("md_code")
    
    if not isnull(rs6("md_averageecost")) then 
       average_cost = Round(rs6("md_averageecost"),2)
    end if 

    sql11= "delete from tblstock_movement"
    set rs11 = server.CreateObject("adodb.recordset")
    rs11.ActiveConnection = strconnect
    rs11.Source = sql11
    rs11.CursorLocation  = 3
    rs11.Open

    if searchvalue <> "" then
        sql12= "delete from tblmodel_avgcost where md_code ='" & searchvalue &"' and md_entry is NULL" 'Y is the base and initial value , should not be deleted
    else
        sql12= "delete from tblmodel_avgcost where md_entry is NULL" 
    end if
   
    set rs12 = server.CreateObject("adodb.recordset")
    rs12.ActiveConnection = strconnect
    rs12.Source = sql12
    rs12.CursorLocation  = 3
    rs12.Open

       sql1="insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_avg_cost, stk_purchase_cost, stk_logby) " & _
       "select top 1 a.stk_itm_code,'"& rs6("md_date") &"','Balance','Balance C/F','0.00',(select sum (a.stk_qty) from tblstocktran a " & _
       "where a.stk_itm_code ='" & searchvalue & "' and cast(a.stk_date as date) <='"& rs6("md_date") &"') as 'bal_cf', '" & average_cost &"', " & _
       "'0.00', '" & Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a where a.stk_itm_code ='" & searchvalue & "'"
 
        set rs1 = server.CreateObject("adodb.recordset")
        rs1.ActiveConnection = strconnect
        rs1.Source = sql1
        rs1.CursorLocation  = 3
        rs1.Open

        sql3= "insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty,stk_bf_qty,stk_avg_cost, stk_purchase_cost, stk_logby)" & _
        "select a.stk_itm_code, a.stk_date,a.stk_type, stk_voucherno,stk_qty,'0.00','0.00',stk_cost_price, '" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a " & _
        "where CAST(stk_date as date) > '"& rs6("md_date") &"' and stk_itm_code = '" & searchvalue & "' order by stk_date "

    set rs3 = server.CreateObject("adodb.recordset")
    rs3.ActiveConnection = strconnect
    rs3.Source = sql3
    rs3.CursorLocation  = 3
    rs3.Open
 
    sql4= "select stk_move_id, stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_purchase_cost, stk_avg_cost,stk_total_cost " & _
    "from tblstock_movement a where stk_code_id = '" & searchvalue & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
    
    initial=0
    prev_stk_avg_cost=0
    balanceBFqty=0
    prev_total_cost=0
    weighted_avg_cost=0
    totalpurchase_cost=0

    while Not rs4.EOF
    totalpurchase_cost = 0
    total_cost= 0
          
    if rs4("stk_doc_no") = "Balance C/F" then
         total_cost = rs4("stk_bf_qty") * (rs4("stk_avg_cost"))
         balanceBFqty = rs4("stk_bf_qty") 
         weighted_avg_cost = rs4("stk_avg_cost")
         prev_total_cost = rs4("stk_total_cost")
         initial=1
    else 
        if rs4("stk_doc_type") = "Stock-In" then
           totalpurchase_cost = rs4("stk_qty") *  rs4("stk_purchase_cost")
        else 
            if not isnull(prev_stk_avg_cost) then
                 totalpurchase_cost = rs4("stk_qty") * Round(prev_stk_avg_cost,2)
            end if 
        end if 
        total_cost = totalpurchase_cost + prev_total_cost
        balanceBFqty = balanceBFqty + rs4("stk_qty")
   
      if total_cost = 0 or total_cost < 0 then ' to cater for certain exception such as tot purchase cost (-83)  + 83 will result in zero. Continous outgoing without incoming can result in this
            total_cost = prev_total_cost
       end if

     if total_cost < 0 or balanceBFqty < 0 then 'this is to prevent sys crash if divided by 0
        response.write " ***ERROR code **** "
        response.write rs4("stk_code_id")
      end if 

      if balanceBFqty <> 0 then 'this is to prevent sys crash if divided by 0
        weighted_avg_cost = abs(total_cost) / abs(balanceBFqty)
      end if   
          
      if rs4("stk_doc_type") <> "Stock-In" and rs4("stk_doc_type") <> "Balance C/F" then 
            weighted_avg_cost = prev_stk_avg_cost
      end if
   end if 

    if not isNull(weighted_avg_cost) then
        weighted_avg_cost=Round(weighted_avg_cost,2)
    end if 

    if not isNull(prev_stk_avg_cost) then
        prev_stk_avg_cost=Round(prev_stk_avg_cost,2)
    end if 

    if weighted_avg_cost <> prev_stk_avg_cost then 
        'if abs(weighted_avg_cost - prev_stk_avg_cost) > 0.02 and initial <> 1 and weighted_avg_cost <> "0" then 
   ' if abs(weighted_avg_cost - prev_stk_avg_cost) > 0.02 and initial <> 1 and weighted_avg_cost <> "0" then 
    if initial <> 1 and weighted_avg_cost <> "0" then 
            SQL10 = "insert into tblmodel_avgcost (md_code, md_averagecost, md_date) values ('" & rs4("stk_code_id") & "','" & weighted_avg_cost & "', '" & rs4("stk_date") & "')"
            set rs10 = server.CreateObject("adodb.recordset")
            rs10.ActiveConnection = strconnect
            rs10.Source = sql10
            rs10.CursorLocation  = 3
            rs10.Open
        end if 
    end if 
        
    if not isNull(totalpurchase_cost) then
       totalpurchase_cost=Round(totalpurchase_cost,2)
    end if 

    if not isNull(total_cost) then
       total_cost=Round(total_cost,2)
    end if
    
        sql7= "update tblstock_movement set stk_bf_qty =  '" & balanceBFqty & "', stk_purchase_totalcost = '" & totalpurchase_cost & "', stk_total_cost =  '" & total_cost & "', " & _ 
        "stk_avg_cost = '" & weighted_avg_cost & "' where stk_move_id ='" & rs4("stk_move_id") & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'"
        set rs7 = server.CreateObject("adodb.recordset")
        rs7.ActiveConnection = strconnect
        rs7.Source = sql7
        rs7.CursorLocation  = 3
        rs7.Open    
    
    prev_stk_avg_cost = weighted_avg_cost
    prev_total_cost = total_cost
    
   if isNumeric(prev_total_cost) = False then 
      prev_total_cost = 0.00
   end if 

   if isNumeric(balanceBFqty) = False then 
      balanceBFqty = 0.00
   end if

    initial=0
    rs4.movenext
	wend
    
    sql4= "update tblmodel set md_averageecost = '" & weighted_avg_cost & "' where md_code = '" & searchvalue & "'"
    CUD(sql4)            
  
    rs6.movenext
    response.write "********  Successfully Updated  ************"
    wend    
   end if
   %>    
            
<!-- #include file="footer.asp" -->