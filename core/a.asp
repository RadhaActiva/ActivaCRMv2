<%Response.Buffer = True%>
<!-- #include file="database/datastore.asp" -->
<!-- #include file="ExcelADO.asp" -->
<%


         
'Update Stock - Vector 
	Cost
			''''
			sql3 = "SELECT tblstocktran.stk_itm_code, sum(stk_qty) as totalqty, round(sum(stk_cost_price * stk_qty),2) as totalstockAmt,	round(sum(stk_cost_price * stk_qty)/sum(stk_qty),2) as AvgAmt " & _
				   "FROM tblstocktran where  stk_qty > 0  group by tblstocktran.stk_itm_code"		
			'response.write sql1
			'response.End()	   
			set rs3 = server.CreateObject("adodb.recordset")
			rs3.Open sql3,strconnect,3,3,&H0001
			while Not rs3.EOF				
				'Update Stock - Vector Average Cost			
				sql2 = "update tblmodel set md_averageecost=" & chknumber(rs3("AvgAmt")) & " where md_code='" & ChkString(rs3("stk_itm_code")) & "'"
				CUD(sql2)	
				response.write 	ChkString(rs3("stk_itm_code")) & "<br>"
			rs3.movenext
			wend
			rs3.close    
		

response.write "end"
%>

