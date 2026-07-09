<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<ol>
  <li><a href="a_procestocktrans.asp?type=resetstocktrans" target="_blank">
    
    reset stocktrans and reset tblwarehouse_stock
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processstock-in" target="_blank">process stock-in (Posted)
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processstock-out" target="_blank">process stock-out (Posted) </a></li>
  <li><a href="a_procestocktrans.asp?type=processstock-transfer" target="_blank">process stock-transfer (Posted)
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processsjob" target="_blank">process job
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processsDO" target="_blank">process DO
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processsparepartrequest" target="_blank">process sparepart request
  </a></li>
  <li><a href="a_procestocktrans.asp?type=processsCN" target="_blank">process CN</a></li>
  <li><a href="a_procestocktrans.asp?type=processsAverageCost" target="_blank">process Average Cost</a></li>
   <li><a href="a_procestocktrans.asp?type=processsStockAdjustment" target="_blank">process Stock Adjustment</a></li>
</ol>
    <%

' 1. reset stocktrans and reset tblwarehouse_stock
' 2. process stock-in (Posted)
' 3. process stock-out (Posted)
' 4. process stock-transfer (Posted)
' 5. process job
' 6. process DO
' 7. process sparepart request
' 8. process CN
' 9. process Average Cost
' 10. process Stock Adjustment

'----------------------------------------------------------------------------------------------------    
' 1. reset stocktrans - done.  

 if request("type") = "resetstocktrans" then 
    
	sql1="truncate table tblwarehouse_stock" 
	CUD(sql1) 
	
	sql1="truncate table tblstocktran" 
	CUD(sql1) 
	
	response.write "1. reset stocktrans - done." 

end if

'----------------------------------------------------------------------------------------------------    

' 2. process stock-in (Posted)         
  if request("type") = "processstock-in" then 
    
 		'''''Stock-In Detail
		sql1 = "SELECT tblstockin_detail.std_id, tblstockin_detail.std_st_no, tblstockin_detail.std_itm_code, tblstockin_detail.std_itm_desc, " & _
		       "tblstockin_detail.std_unitcost, tblstockin_detail.std_qty, tblstockin_detail.std_subtotal, tblstockin_detail.std_referid, " & _
			   "tblstockin.st_towarehouse, tblstockin.st_fromwarehouse, tblstockin.st_approvedby, tblstockin.st_approveddate, tblstockin.st_no " & _
		       "FROM tblstockin_detail inner join tblstockin on tblstockin_detail.std_st_no=tblstockin.st_no where tblstockin.st_status='Approved' order by tblstockin_detail.std_id"	   
		'response.write sql1
		'response.End()
		
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("st_towarehouse") & "' and wst_itm_code = '" & rs1("std_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = rs1("st_towarehouse")
				rs2("wst_itm_code") = ChkString(rs1("std_itm_code"))
				wst_itm_current_qty = ChkString(rs1("std_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("std_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = rs1("st_approvedby")
				rs2("wst_lastupdatedate")  =rs1("st_approveddate")
			else
			    wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("std_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("std_qty") 
				rs2("wst_lastupdateby")  = rs1("st_approvedby")
				rs2("wst_lastupdatedate")  = rs1("st_approveddate")
			end if
			rs2.Update 
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("st_no")
			rs2("stk_reference") = rs1("st_towarehouse")
			rs2("stk_date")  = rs1("st_approveddate")
			rs2("stk_type")  = "Stock-In"
			rs2("stk_itm_code")  = ChkString(rs1("std_itm_code"))
			rs2("stk_fromwarehouse")  = rs1("st_fromwarehouse")
			rs2("stk_towarehouse")  = rs1("st_towarehouse")
			rs2("stk_desc")  = ChkString(rs1("std_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("std_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("std_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("std_unitcost"))
			rs2("stk_logby")  = rs1("st_approvedby")
			rs2("stk_logdate")  = rs1("st_approveddate")
			rs2.Update 
			rs2.Close  
			
			'Update Stock - Vector Average Cost
			sql2 = "select sum(stk_cost_price*stk_qty)/sum(stk_qty) as stk_cost_price from tblstocktran where stk_qty > 0 and stk_itm_code='" & ChkString(rs1("std_itm_code")) & "'"
		    stk_cost_price = selectid(sql2)
			
			if isnull(stk_cost_price) or stk_cost_price=0 then 
			   sql2 = "select top 1 md_averageecost from  tblmodel where md_code='" & ChkString(rs1("std_itm_code")) & "' "
			   stk_cost_price = selectid(sql2)
			end if 
			
			sql2 = "update tblmodel set md_averageecost=" & chknumber(stk_cost_price) & " where md_code='" & ChkString(rs1("std_itm_code")) & "'"
			CUD(sql2)		
		rs1.movenext
		wend
		rs1.close 
		
		response.write "2. process stock-in (Posted) - done."
		
	end if	
'----------------------------------------------------------------------------------------------------    
' 3. process stock-out (Posted)

  if request("type") = "processstock-out" then 
  
		'''''Stock-Out Detail
		sql1 = "SELECT tblstockOut_detail.sod_id, tblstockOut_detail.sod_so_no, tblstockOut_detail.sod_itm_code, tblstockOut_detail.sod_itm_desc, " & _
		       "tblstockOut_detail.sod_unitcost, tblstockOut_detail.sod_qty, tblstockOut_detail.sod_subtotal, tblstockOut_detail.sod_referid, " & _
			   "tblstockout.so_towarehouse, tblstockout.so_fromwarehouse, tblstockout.so_approvedby, tblstockout.so_approveddate, tblstockout.so_no " & _
	           "FROM tblstockOut_detail inner join tblstockout on tblstockOut_detail.sod_so_no=tblstockout.so_no " & _
			   "where tblstockout.so_status='approved' order by tblstockOut_detail.sod_id"	    
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("so_fromwarehouse") & "' and wst_itm_code = '" & rs1("sod_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
			    wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("sod_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("sod_qty") 
				rs2("wst_lastupdateby")  = rs1("so_approvedby")
				rs2("wst_lastupdatedate")  = rs1("so_approveddate")
			rs2.Update 
			end if			
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1  stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("so_no")
			rs2("stk_reference") = rs1("so_fromwarehouse")
			rs2("stk_date")  = rs1("so_approveddate")
			rs2("stk_type")  = "Stock-Out"
			rs2("stk_itm_code")  = ChkString(rs1("sod_itm_code"))
			rs2("stk_fromwarehouse")  = rs1("so_fromwarehouse")
			rs2("stk_towarehouse")  = rs1("so_towarehouse")
			rs2("stk_desc")  = ChkString(rs1("sod_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sod_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sod_subtotal"))
			rs2("stk_logby")  = rs1("so_approvedby")
			rs2("stk_logdate")  = rs1("so_approveddate")
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		
		response.write "3. process stock-out (Posted) - done."
		
  end if
  
  
  '----------------------------------------------------------------------------------------------------    

  ' 4. process stock-transfer (Posted)      
  if request("type") = "processstock-transfer" then 
  
		'''''Stock-Transfer Detail
		sql1 = "SELECT tblstocktran_detail.sfd_id, tblstocktran_detail.sfd_st_no, tblstocktran_detail.sfd_itm_code, tblstocktran_detail.sfd_itm_desc, " & _
		       "tblstocktran_detail.sfd_unitcost, tblstocktran_detail.sfd_qty, tblstocktran_detail.sfd_subtotal, tblstocktran_detail.sfd_referid, " & _
			   "tblstocktransfer.sf_no, tblstocktransfer.sf_approveddate, tblstocktransfer.sf_approvedby, tblstocktransfer.sf_fromwarehouse,tblstocktransfer.sf_towarehouse  " & _
		       "FROM tblstocktran_detail inner join tblstocktransfer on tblstocktransfer.sf_no=tblstocktran_detail.sfd_st_no " & _
			   "where tblstocktransfer.sf_status='Approved' order by tblstocktran_detail.sfd_id"	   
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock Transfer out	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _
	               "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("sf_fromwarehouse") & "' and wst_itm_code = '" & rs1("sfd_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then 
			    wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("sfd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("sfd_qty") 
				rs2("wst_lastupdateby")  = rs1("sf_approvedby")
				rs2("wst_lastupdatedate")  = rs1("sf_approveddate")
				rs2.Update 
			end if
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("sf_no")
			rs2("stk_reference") = rs1("sf_fromwarehouse")
			rs2("stk_date")  = rs1("sf_approveddate")
			rs2("stk_type")  = "Stock-Transfer-Out"
			rs2("stk_itm_code")  = ChkString(rs1("sfd_itm_code"))
			rs2("stk_fromwarehouse")  = rs1("sf_fromwarehouse")
			rs2("stk_towarehouse")  = rs1("sf_towarehouse") 
			rs2("stk_desc")  = ChkString(rs1("sfd_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sfd_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sfd_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("sfd_unitcost"))
			rs2("stk_logby")  = rs1("sf_approvedby")
			rs2("stk_logdate")  = rs1("sf_approveddate")
			rs2.Update 
			rs2.Close  			
			
			''''Add Stock Transfer In	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _
	               "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("sf_towarehouse") & "' and wst_itm_code = '" & rs1("sfd_itm_code") & "'"	   
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = rs1("sf_towarehouse")
				rs2("wst_itm_code") = ChkString(rs1("sfd_itm_code"))
				
				wst_itm_current_qty = ChkString(rs1("sfd_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("sfd_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = rs1("sf_approvedby")
				rs2("wst_lastupdatedate")  = rs1("sf_approveddate")
				rs2.Update 
			else
			    wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("sfd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("sfd_qty") 
				rs2("wst_lastupdateby")  = rs1("sf_approvedby")
				rs2("wst_lastupdatedate")  = rs1("sf_approveddate")
				rs2.Update 
			end if
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1  stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_cost_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("sf_no")
			rs2("stk_reference") = rs1("sf_towarehouse")
			rs2("stk_date")  = rs1("sf_approveddate")
			rs2("stk_type")  = "Stock-Transfer-In"
			rs2("stk_itm_code")  = ChkString(rs1("sfd_itm_code"))
			rs2("stk_fromwarehouse")  = rs1("sf_fromwarehouse")
			rs2("stk_towarehouse")  = rs1("sf_towarehouse")
			rs2("stk_desc")  = ChkString(rs1("sfd_itm_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("sfd_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("sfd_subtotal"))
			rs2("stk_cost_price")  = ChkNumber(rs1("sfd_unitcost"))
			rs2("stk_logby")  = rs1("sf_approvedby")
			rs2("stk_logdate")  = rs1("sf_approveddate")
			rs2.Update 
			rs2.Close  
			
		rs1.movenext
		wend
		rs1.close
 
  
  response.write "4. process stock-transfer (Posted) - done."
  end if
 
  '----------------------------------------------------------------------------------------------------    

  ' 5. process job     
  if request("type") = "processsjob" then 
  
		sql1 = "SELECT tbljob_parts.jobp_id, tbljob_parts.job_code, tbljob_parts.jobp_partcode, tbljob_parts.jobp_desc,  " & _
				"tbljob_parts.jobp_unitcost, tbljob_parts.jobp_discountamt, tbljob_parts.jobp_discounttype,  " & _
				"tbljob_parts.jobp_netcost, tbljob_parts.jobp_qty, tbljob_parts.jobp_subtotal, tbljob.job_posteddate,  " & _
				"tbljob.job_postedby, tbltechnician.tech_wh_code  " & _
				"FROM  tbljob inner join tbljob_parts on tbljob.job_code=tbljob_parts.job_code " & _
				"inner join tbltechnician on tbljob.job_tech_code=tbltechnician.tech_code " & _
				"where tbljob.job_status='Posted' order by tbljob_parts.jobp_id "   
		'response.write sql1
		'response.End()
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("tech_wh_code") & "' and wst_itm_code = '" & rs1("jobp_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
			    wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("jobp_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("jobp_qty") 
				rs2("wst_lastupdateby")  = rs1("job_postedby")
				rs2("wst_lastupdatedate")  = rs1("job_posteddate")
				rs2.Update 
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1  stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
				   "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("job_code")
			rs2("stk_reference") = rs1("tech_wh_code")
			rs2("stk_date")  = rs1("job_posteddate")
			rs2("stk_type")  = "Job"
			rs2("stk_itm_code")  = ChkString(rs1("jobp_partcode"))
			rs2("stk_fromwarehouse")  = rs1("tech_wh_code")
			rs2("stk_towarehouse")  = "Job" 
			rs2("stk_desc")  = ChkString(rs1("jobp_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("jobp_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("jobp_subtotal"))
			rs2("stk_logby")  = rs1("job_postedby")
			rs2("stk_logdate")  =rs1("job_posteddate")
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
  
  response.write "5. process job  - done."
  end if
  
  '----------------------------------------------------------------------------------------------------    

  ' 6. process DO   
  if request("type") = "processsDO" then   
      
		'''''DO Detail
		sql1 = "SELECT tbldo_detail.dod_id, tbldo_detail.dod_do_no, tbldo_detail.dod_inv_no, tbldo_detail.dod_job_code, tbldo_detail.dod_partcode, " & _
		       "tbldo_detail.dod_desc, tbldo_detail.dod_unitcost, tbldo_detail.dod_qty, tbldo_detail.dod_discountamt, " & _
			   "tbldo_detail.dod_discounttype, tbldo_detail.dod_netcost, tbldo_detail.dod_subtotal, " & _
			   "tbldo.do_createdby, tbldo.do_createddate, tbldo.do_no " & _
			   "FROM tbldo_detail inner join tbldo on tbldo_detail.dod_do_no=tbldo.do_no where tbldo.do_status='Posted' order by tbldo_detail.dod_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add DO Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("dod_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
			    wst_itm_current_qty = rs2("wst_itm_current_qty") - rs1("dod_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("dod_qty") 
				rs2("wst_lastupdateby")  = rs1("do_createdby")
				rs2("wst_lastupdatedate")  = rs1("do_createddate")
				rs2.Update 
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("do_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = rs1("do_createddate")
			rs2("stk_type")  = "DO"
			rs2("stk_itm_code")  = ChkString(rs1("dod_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("dod_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("dod_qty")*-1)
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("dod_subtotal"))
			rs2("stk_logby")  = rs1("do_createdby")
			rs2("stk_logdate")  = rs1("do_createddate")
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
        
  response.write "6. process DO   - done."
  end if
  
  '----------------------------------------------------------------------------------------------------    

  ' 7. process sparepart request 
  if request("type") = "processsparepartrequest" then 
  
		'''''Stock-In Detail
		sql1 = "SELECT tblsparepartrequest_detail.spd_id, tblsparepartrequest_detail.spd_sp_no, " & _
				"tblsparepartrequest_detail.spd_tech_code, tblsparepartrequest_detail.spd_partcode, tblsparepartrequest_detail.spd_currentstock,  " & _
				"tblsparepartrequest_detail.spd_desc, tblsparepartrequest_detail.spd_unitcost, tblsparepartrequest_detail.spd_qty,  " & _
				"tblsparepartrequest_detail.spd_subtotal, tblsparepartrequest.sp_posteddate, tblsparepartrequest.sp_postedby, tbltechnician.tech_wh_code  " & _
				"FROM tblsparepartrequest inner join tbltechnician on tblsparepartrequest.sp_tech_code=tbltechnician.tech_code  " & _
				"inner join tblsparepartrequest_detail    " & _
				"on tblsparepartrequest_detail.spd_sp_no=tblsparepartrequest.sp_no  " & _
				"where tblsparepartrequest.sp_status='Posted' order by tblsparepartrequest_detail.spd_id "	  
		'response.write sql1
		'response.End()
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock In Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("tech_wh_code") & "' and wst_itm_code = '" & rs1("spd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if rs2.eof then 
				rs2.AddNew   
				rs2("wst_wh_code") = rs1("tech_wh_code")
				rs2("wst_itm_code") = ChkString(rs1("spd_partcode"))
				wst_itm_current_qty = ChkString(rs1("spd_qty"))
				rs2("wst_itm_current_qty")  = ChkString(rs1("spd_qty"))
				rs2("wst_itm_min_qty")  = 0
				rs2("wst_lastupdateby")  = rs1("sp_postedby")
				rs2("wst_lastupdatedate")  = rs1("sp_posteddate")
			else
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("spd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("spd_qty") 
				rs2("wst_lastupdateby")  = rs1("sp_postedby")
				rs2("wst_lastupdatedate")  = rs1("sp_posteddate")
			end if
			rs2.Update 
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1  stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("spd_sp_no")
			rs2("stk_reference") = rs1("tech_wh_code")
			rs2("stk_date")  = rs1("sp_posteddate")
			rs2("stk_type")  = "Spareparts-Request"
			rs2("stk_itm_code")  = ChkString(rs1("spd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = rs1("tech_wh_code")
			rs2("stk_desc")  = ChkString(rs1("spd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("spd_qty"))
			rs2("stk_balanceqty")  = ChkNumber(wst_itm_current_qty)
			rs2("stk_sales_price")  = ChkNumber(rs1("spd_subtotal"))
			rs2("stk_logby")  = rs1("sp_postedby")
			rs2("stk_logdate")  = rs1("sp_posteddate")
			rs2.Update 
			rs2.Close    
			
			''''Add Stock In Detail - W1	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("spd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") - rs1("spd_qty") 
				rs2("wst_lastupdateby")  = rs1("sp_postedby")
				rs2("wst_lastupdatedate")  = rs1("sp_posteddate")
			rs2.Update 
			end if
			rs2.Close   
			
			'Update Stocktrans - W1
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("spd_sp_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = rs1("sp_posteddate")
			rs2("stk_type")  = "Spareparts-Request"
			rs2("stk_itm_code")  = ChkString(rs1("spd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = rs1("tech_wh_code")
			rs2("stk_desc")  = ChkString(rs1("spd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("spd_qty")*-1)
			rs2("stk_balanceqty")  = 0
			rs2("stk_sales_price")  = ChkNumber(rs1("spd_subtotal"))
			rs2("stk_logby")  = rs1("sp_postedby")
			rs2("stk_logdate")  = rs1("sp_posteddate")
			rs2.Update 
			rs2.Close   
		
		rs1.movenext
		wend
		rs1.close		
  
  response.write "7. process sparepart request    - done."
  end if
    
  '----------------------------------------------------------------------------------------------------    

  ' 8. process CN        
  if request("type") = "processsCN" then 
	
		'''''CN Detail
		sql1 = "SELECT tblcn_detail.cnd_id, tblcn_detail.cnd_cn_no, tblcn_detail.cnd_inv_no, tblcn_detail.cnd_job_code, tblcn_detail.cnd_partcode, " & _
				"tblcn_detail.cnd_desc, tblcn_detail.cnd_unitcost, tblcn_detail.cnd_qty, tblcn_detail.cnd_discountamt, " & _ 
				"tblcn_detail.cnd_discounttype, tblcn_detail.cnd_netcost, tblcn_detail.cnd_subtotal, " & _
				"tblcn.cn_postedby, tblcn.cn_posteddate, tblcn.cn_inv_no, tblcn.cn_totalAmt, tblcn.cn_no " & _
				"FROM tblcn_detail inner join tblcn on tblcn_detail.cnd_cn_no=tblcn.cn_no where tblcn.cn_status='Posted' order by cnd_id"	      
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add CN Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = 'W1' and wst_itm_code = '" & rs1("cnd_partcode") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				wst_itm_current_qty = rs2("wst_itm_current_qty") + rs1("cnd_qty") 
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("cnd_qty") 
				rs2("wst_lastupdateby")  = rs1("cn_postedby")
				rs2("wst_lastupdatedate")  = rs1("cn_posteddate")
				rs2.Update 
			end if
			rs2.Close    
		
			'Update Stocktrans - Stock Movement
			sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
			       "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
			rs2.Open sql2,strconnect,2,2,&H0001
			rs2.AddNew   
			rs2("stk_voucherno") = rs1("cn_no")
			rs2("stk_reference") = "W1"
			rs2("stk_date")  = rs1("cn_posteddate")
			rs2("stk_type")  = "CN"
			rs2("stk_itm_code")  = ChkString(rs1("cnd_partcode"))
			rs2("stk_fromwarehouse")  = "W1"
			rs2("stk_towarehouse")  = ""
			rs2("stk_desc")  = ChkString(rs1("cnd_desc"))
			rs2("stk_qty")  = ChkNumber(rs1("cnd_qty"))
			rs2("stk_balanceqty")  = wst_itm_current_qty
			rs2("stk_sales_price")  = ChkNumber(rs1("cnd_subtotal"))
			rs2("stk_logby")  = rs1("cn_postedby")
			rs2("stk_logdate")  = rs1("cn_posteddate")
			rs2.Update 
			rs2.Close   
			
		rs1.movenext
		wend
		rs1.close
		  
  response.write "8. process CN  - done."
  end if
 
 
  '----------------------------------------------------------------------------------------------------    

  ' 9. process Average Cost        
  if request("type") = "processsAverageCost" then  

		
		sql1 = "update tblmodel set md_averageecost=md_costprice"
		
		'''''Stock-In Detail		
		sql1 = "SELECT tblstocktran.stk_itm_code, sum(stk_qty) as totalqty, sum(stk_cost_price * stk_qty) as totalstockAmt,	sum(stk_cost_price * stk_qty)/sum(stk_qty) as AvgAmt " & _
	           "FROM tblstocktran where  stk_qty > 0 group by tblstocktran.stk_itm_code"		
		'response.write sql1
		'response.End()	   
		set rs1 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF				
			'Update Stock - Vector Average Cost			
			sql2 = "update tblmodel set md_averageecost=" & chknumber(rs1("AvgAmt")) & " where md_code='" & ChkString(rs1("stk_itm_code")) & "'"
			CUD(sql2)		
		rs1.movenext
		wend
		rs1.close 
		
		response.write "9. process Average Cost - done."
		  
  end if
  '----------------------------------------------------------------------------------------------------    
  
  ' 10. Stock Adjustment    
  if request("type") = "processsStockAdjustment" then  
  
		'''''Stock-Adj Detail
		sql1 = "SELECT tblstockadj_detail.sjd_id, tblstockadj_detail.sjd_sj_no, tblstockadj_detail.sjd_itm_code, tblstockadj_detail.sjd_itm_desc, " & _
		       "tblstockadj_detail.sjd_unitcost, tblstockadj_detail.sjd_current_qty, tblstockadj_detail.sjd_adjust_qty, tblstockadj_detail.sjd_diff_qty, " & _
			   "tblstockadj_detail.sjd_subtotal, tblstockadj_detail.sjd_referid, " & _
			   "tblstockadj.sj_status, tblstockadj.sj_approvedby, tblstockadj.sj_approveddate, tblstockadj.sj_no, tblstockadj.sj_fromwarehouse, tblstockadj.sj_towarehouse " & _
	           "FROM tblstockadj_detail inner join tblstockadj on tblstockadj_detail.sjd_sj_no=tblstockadj.sj_no " & _
			   "where tblstockadj.sj_status = 'Approved' order by tblstockadj_detail.sjd_id"	     
		'response.write sql1
		set rs1 = server.CreateObject("adodb.recordset")
		set rs2 = server.CreateObject("adodb.recordset")
		rs1.Open sql1,strconnect,3,3,&H0001
		while Not rs1.EOF
				
			''''Add Stock Adj Detail	   	  
			sql2 = "SELECT wst_id, wst_wh_code, wst_itm_code, wst_itm_current_qty, wst_itm_min_qty, wst_itm_remarks, wst_lastupdateby, wst_lastupdatedate " & _ 
				   "FROM tblwarehouse_stock where wst_wh_code = '" & rs1("sj_fromwarehouse") & "' and wst_itm_code = '" & rs1("sjd_itm_code") & "'"
			rs2.Open sql2,strconnect,2,2,&H0001
			if not rs2.eof then
				rs2("wst_itm_current_qty")  = rs2("wst_itm_current_qty") + rs1("sjd_diff_qty") 
				rs2("wst_lastupdateby")  = rs1("sj_approvedby")
				rs2("wst_lastupdatedate")  = rs1("sj_approveddate")
			    rs2.Update 
			end if			
			rs2.Close   
			
			'Update Stocktrans - Stock Movement
				sql2 = "SELECT top 1 stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, " & _
					   "stk_qty, stk_balanceqty, stk_sales_price, stk_logby, stk_logdate FROM tblstocktran "
				rs2.Open sql2,strconnect,2,2,&H0001
				rs2.AddNew   
				rs2("stk_voucherno") = rs1("sj_no") 
				rs2("stk_reference") = rs1("sj_fromwarehouse")
				rs2("stk_date")  = rs1("sj_approveddate")
				rs2("stk_type")  = "Stock-Adj"
				rs2("stk_itm_code")  = ChkString(rs1("sjd_itm_code"))
				rs2("stk_fromwarehouse")  = rs1("sj_fromwarehouse")
				rs2("stk_towarehouse")  = rs1("sj_towarehouse") 
				rs2("stk_desc")  = ChkString(rs1("sjd_itm_desc"))
				rs2("stk_qty")  = ChkNumber(rs1("sjd_diff_qty"))
				rs2("stk_balanceqty")  = 0
				rs2("stk_sales_price")  = ChkNumber(rs1("sjd_subtotal"))
				rs2("stk_logby")  = rs1("sj_approvedby")
				rs2("stk_logdate")  = rs1("sj_approveddate")
				rs2.Update 
				rs2.Close  
		
		rs1.movenext
		wend
		rs1.close
  
  response.write "10. Stock Adjustment - done."
    
  end if
   
   
  '----------------------------------------------------------------------------------------------------    
  
%>

