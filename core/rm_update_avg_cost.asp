<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
wh_code = request("wh_code")
whchk = request("whchk")
reccount=request("reccount")
      
    sql6= "select md_code, md_costprice, (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = tblmodel.md_code and CAST(tblmodel_avgcost.md_date as date) <= '20220108' " & _ 
    "order by md_date desc) as md_averageecost from tblmodel" ' where md_code='WCF36-00'"
   
    set rs6 = server.CreateObject("adodb.recordset")
    rs6.ActiveConnection = strconnect
    rs6.Source = sql6
    rs6.CursorLocation  = 3
    rs6.Open
   
  while Not rs6.EOF
    ' 12/11/2022 need to reset the values prior to calc
    weighted_avg_cost = 0
    average_cost = 0
    searchvalue = rs6("md_code")
    
    if not isnull(rs6("md_averageecost")) then 
       average_cost = round(rs6("md_averageecost"),2)
     end if 

    sql11= "delete from tblstock_movement"
    set rs11 = server.CreateObject("adodb.recordset")
    rs11.ActiveConnection = strconnect
    rs11.Source = sql11
    rs11.CursorLocation  = 3
    rs11.Open

      'inserting initialvalue
       sql1="insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_avg_cost, stk_purchase_cost, stk_logby) " & _
       "select top 1 a.stk_itm_code,'2022-01-01','Balance','Balance C/F','0.00',(select sum (a.stk_qty) from tblstocktran a " & _
       "where a.stk_itm_code ='" & searchvalue & "' and a.stk_date <='01/01/2022') as 'bal_cf', '" & average_cost &"', " & _
       "'"& rs6("md_costprice") &"', '" & Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a where a.stk_itm_code ='" & searchvalue & "'"
    
        set rs1 = server.CreateObject("adodb.recordset")
        rs1.ActiveConnection = strconnect
        rs1.Source = sql1
        rs1.CursorLocation  = 3
        rs1.Open

    'inserting from stocktrans table
        sql3= "insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty,stk_bf_qty,stk_avg_cost, stk_purchase_cost, stk_logby)" & _
        "select a.stk_itm_code, a.stk_date,a.stk_type, stk_voucherno,stk_qty,'0.00','0.00',stk_cost_price, '" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a " & _
        "where CAST(stk_date as date) >= '01/01/2022' and stk_itm_code = '" & searchvalue & "' order by stk_date "

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
           totalpurchase_cost = rs4("stk_qty") * round(prev_stk_avg_cost,2)
        end if 
        total_cost = totalpurchase_cost + prev_total_cost
        balanceBFqty = balanceBFqty + rs4("stk_qty")
   
      if total_cost > 0 and balanceBFqty > 0 then 'this is to prevent sys crash if divided by 0
        weighted_avg_cost = total_cost / balanceBFqty       
    end if
    
      if rs4("stk_doc_type") <> "Stock-In" and rs4("stk_doc_type") <> "Balance C/F" then 
            weighted_avg_cost = prev_stk_avg_cost
      end if
   end if 

    if not isNull(weighted_avg_cost) then
        weighted_avg_cost=round(weighted_avg_cost,2)
    end if 

    if not isNull(prev_stk_avg_cost) then
        prev_stk_avg_cost=round(prev_stk_avg_cost,2)
    end if 

    if weighted_avg_cost <> prev_stk_avg_cost then 
     if abs(weighted_avg_cost - prev_stk_avg_cost) > 0.02 and initial <> 1 and weighted_avg_cost <> "0" then 
            SQL10 = "insert into tblmodel_avgcost (md_code, md_averagecost, md_date) values ('" & rs4("stk_code_id") & "','" & weighted_avg_cost & "', '" & rs4("stk_date") & "')"
            set rs10 = server.CreateObject("adodb.recordset")
            rs10.ActiveConnection = strconnect
            rs10.Source = sql10
            rs10.CursorLocation  = 3
            rs10.Open 
        end if 
    end if 
        
    if not isNull(totalpurchase_cost) then
       totalpurchase_cost=round(totalpurchase_cost,2)
    end if 

    if not isNull(total_cost) then
       total_cost=round(total_cost,2)
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

    'if average_cost <> weighted_avg_cost and weighted_avg_cost <> 0 then
        sql4= "update tblmodel set md_averageecost = '" & weighted_avg_cost & "' where md_code = '" & searchvalue & "'"
        set rs4 = server.CreateObject("adodb.recordset")
        rs4.ActiveConnection = strconnect
        rs4.Source = sql4
        rs4.CursorLocation  = 3
        rs4.Open  
    'end if
  
    rs6.movenext
    wend
   %>