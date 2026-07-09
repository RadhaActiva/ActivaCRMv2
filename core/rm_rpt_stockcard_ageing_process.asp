<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<%
''1. Delete from tblStock_Agieng
''2. INsert tblStock_Agieng
''3. Update Ageing Report

if request("stock_aging_date") <> "" then 
   stock_aging_date = request("stock_aging_date")
else   
   stock_aging_date = chkdate(date())
end if
	
	'''1. Delete from tblStock_Agieng
	sql = "Delete from tblStock_Agieng"
	CUD(sql)
	
	''2. INsert tblStock_Agieng
	sql = "INSERT INTO tblStock_Agieng (ag_stock_code, ag_current_stock) SELECT   md_code,  " & _
		"(SELECT     SUM(stk_qty) AS totalqty " & _
		"FROM          tblstocktran " & _
		"WHERE      (stk_itm_code = tblmodel.md_code)) AS totalqty " & _
		"FROM         tblmodel " & _
		"WHERE     (md_code IS NOT NULL) " 
	CUD(sql)	
	
	sql1 = "WITH cte AS ( " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 0 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 0 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 12 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 1 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 12 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 24 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 2 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 24 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 36 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 3 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 36 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 48 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 4 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 48 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 60 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 5 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 60 and DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') < 72 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
		"UNION " & _
		"SELECT tblstockin_detail.std_itm_code, SUM(tblstockin_detail.std_qty) as std_qty, 6 AS Agiengyear, 0 AS totalqty, 0 AS md_averageecost " & _
		"FROM tblstockin_detail INNER JOIN tblstockin ON tblstockin_detail.std_st_no = tblstockin.st_no " & _
		"WHERE DATEDIFF(month, st_approveddate, '" & stock_aging_date & "') >= 72 AND st_status='Approved' " & _
		"GROUP BY tblstockin_detail.std_itm_code " & _
	") " & _
	"SELECT std_itm_code, std_qty, Agiengyear, SUM(tblstocktran.stk_qty) AS totalqty, " & _
	"(select TOP 1 md_averagecost from tblmodel_avgcost " & _
	"    where std_itm_code = tblmodel_avgcost.md_code " & _ 
	"    and CAST(tblmodel_avgcost.md_date as date) <= '" & stock_aging_date & "' order by md_date desc ) AS md_averageecost " & _
	"FROM cte " & _
	"LEFT JOIN tblstocktran ON cte.std_itm_code = tblstocktran.stk_itm_code AND tblstocktran.stk_date < '" & ChkDateYYYYMMDD(DateAdd("d",1,stock_aging_date)) & "' and tblstocktran.stk_reference is not null " & _
	"GROUP BY std_itm_code, std_qty, Agiengyear, md_averageecost " & _
	"ORDER BY std_itm_code ASC, Agiengyear ASC " 
	
	avgcost = 0
	agTotalQty = 0
	itemTotalQty = 0
	accumulateQty = 0
	stkinQty = 0
	currentItem = ""
	
	set rs1 = server.CreateObject("adodb.recordset")
	rs1.ActiveConnection = strconnect
	rs1.Source = sql1
	rs1.CursorLocation  = 3
	rs1.Open
	while not rs1.eof 
			  
			If Not IsNull(rs1("std_itm_code")) Then
			
			  if (currentItem <> rs1("std_itm_code")) then			    
				currentItem = rs1("std_itm_code")
				accumulateQty = 0
				
				agTotalQty = chknumber(rs1("totalqty"))		
				sql2 = "Update tblStock_Agieng set ag_current_stock=" & agTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				CUD(sql2)	
				
			  end if 
			  
			  avgcost = chknumber(rs1("md_averageecost"))	
			  itemTotalQty = chknumber(rs1("totalqty"))				
			  
			  stkinQty = chknumber(rs1("std_qty"))
			  sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
		      CUD(sql2)	
			  
			  if rs1("Agiengyear") = "0" then
				' if (itemTotalQty >= stkinQty) then
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y0=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else				    
				'	sql2 = "Update tblStock_Agieng set Y0=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 
				
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y0=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else
					sql2 = "Update tblStock_Agieng set Y0=" & itemTotalQty  & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
				 end if 
				 
				 CUD(sql2)	
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)	
			  end if

		      if rs1("Agiengyear") = "1" then
			  
				' if (itemTotalQty >= stkinQty) then
				'	response.write("match")
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y1=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else
				'   response.write("not match")
				'	sql2 = "Update tblStock_Agieng set Y1=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y1=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y1=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y1=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
				 
				 CUD(sql2)		
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)	
			  end if
			  
			  if rs1("Agiengyear") = "2" then
			    ' if (itemTotalQty >= stkinQty) then
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y2=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else
				'	sql2 = "Update tblStock_Agieng set Y2=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y2=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y2=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y2=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
					
				 CUD(sql2)	
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)	
		      end if 
			  
			  if rs1("Agiengyear") = "3" then
			    ' if (itemTotalQty >= stkinQty) then
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y3=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else
				'	sql2 = "Update tblStock_Agieng set Y3=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y3=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y3=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y3=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
					
				 CUD(sql2)	
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)					 
		      end if 
			  
			  if rs1("Agiengyear") = "4" then
			    ' if (itemTotalQty >= stkinQty) then
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y4=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else
				'	sql2 = "Update tblStock_Agieng set Y4=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y4=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y4=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y4=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
					
				 CUD(sql2)	
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)	
		      end if 
			  
			  if rs1("Agiengyear") = "5" then
			    ' if (itemTotalQty >= stkinQty) then
				'	itemTotalQty = itemTotalQty - stkinQty
				'	sql2 = "Update tblStock_Agieng set Y5=" & stkinQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"						
				' else
				'	sql2 = "Update tblStock_Agieng set Y5=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				'	itemTotalQty = 0
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y5=" & stkinQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y5=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y5=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
				
				 CUD(sql2)	
				 
				 accumulateQty = accumulateQty + stkinQty
				 
				' sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 				
				 CUD(sql2)	
				 				 
		      end if
			  
			  if rs1("Agiengyear") = "6" then
			    ' if ((itemTotalQty >= accumulateQty) And (accumulateQty>0)) then
				'	sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & " where ag_stock_code='" & rs1("std_itm_code") & "'"	
				' else
				'	sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty  & " where ag_stock_code='" & rs1("std_itm_code") & "'"
				' end if 

			     if (stkinQty <= itemTotalQty - accumulateQty) then
					sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
				 else	    
				    if (accumulateQty >= itemTotalQty) then
						sql2 = "Update tblStock_Agieng set Y6=0, ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"
					else
						sql2 = "Update tblStock_Agieng set Y6=" & itemTotalQty - accumulateQty & ", ag_averagecost=" & avgcost & " where ag_stock_code='" & rs1("std_itm_code") & "'"					
					end if
				 end if 
				 
				 CUD(sql2)	
			  end if	
			  
			end if
			  			  						
	rs1.movenext
	wend
	rs1.close
	

	
response.write "Data Ageing process is completed. for Stock Ageing : " & stock_aging_date
%>

