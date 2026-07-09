<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = "Over"
inv_status = request("inv_status")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-30,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("job_tech_code") <> "" then
   job_tech_code = replace(request("job_tech_code"), " ", "")
   arrjob_tech_code = split(job_tech_code,",")
   job_tech_code = replace(job_tech_code, ",", "','")
   listjob_tech_code = listjob_tech_code & job_tech_code
else
   listjob_tech_code = ""
   arrjob_tech_code = split("0,0",",")
end if

function checkTechlList(strv)
for k = 0 to ubound(arrjob_tech_code)
    if arrjob_tech_code(k) = strv then 
	   checkTechlList = true
	   exit for
	else
	   checkTechlList = false
	end if
next
end function
%>
<!--
sql3 = "SELECT tblinvoice_detail.invd_parttype,  " & _  
	"sum(tblinvoice_detail.invd_qty) as invd_qty,  " & _
	"sum(tblinvoice_detail.invd_subtotal) as invd_subtotal,   " & _
	"sum(tblinvoice_detail.invd_itemcost*tblinvoice_detail.invd_qty) as invd_itemcost  " & _
	"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
	"where tblinvoice_detail.invd_id is not null " & _
	"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
-->
<%
'sql3 = "SELECT tblinvoice_detail.invd_parttype,  " & _  
'	"sum(tblinvoice_detail.invd_qty) as invd_qty,  " & _
'	"sum(tblinvoice_detail.invd_subtotal) as invd_subtotal,   " & _
'	"sum(tblmodel.md_averageecost) as md_averagee ,   " & _
'	"sum(tblmodel.md_averageecost*tblinvoice_detail.invd_qty) as md_averageecost ,   " & _
'	"sum(tblinvoice_detail.invd_itemcost*tblinvoice_detail.invd_qty) as invd_itemcost  " & _
'	"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
'	"left join tblmodel on tblmodel.md_code = REPLACE(LTRIM(RTRIM(tblinvoice_detail.invd_partcode)), CHAR(9), '') " & _
'	"where tblinvoice_detail.invd_id is not null " & _
'	"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 

    sql3 = "SELECT tblinvoice_detail.invd_parttype,  " & _  
	"sum(tblinvoice_detail.invd_qty) as invd_qty,  " & _
	"sum(tblinvoice_detail.invd_subtotal) as invd_subtotal,   " & _
	"sum(tblmodel.md_averageecost) as md_averagee ,   " & _
	"sum(tblmodel.md_averageecost*tblinvoice_detail.invd_qty) as md_averageecost ,   " & _
	"sum(tblinvoice_detail.invd_itemcost*tblinvoice_detail.invd_qty) as invd_itemcost  " & _
	"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
	"left join tblmodel on tblmodel.md_code = REPLACE(LTRIM(RTRIM(tblinvoice_detail.invd_partcode)), CHAR(9), '') " & _
	"where tblinvoice_detail.invd_id is not null " & _
	"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	

	if job_tech_code <> "" then 
	   sql3 = sql3 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if

	if inv_status <> "All" and inv_status<>"" then 
	   sql3 = sql3 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
       sql3 = sql3 & " group by invd_parttype "
	   
'response.write sql3	   
'response.End()

response.Cookies("GAPS")("sqlexcel2") = sql3
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql3
rs.CursorLocation  = 3
rs.Open
while not rs.eof
   if rs("invd_parttype") = "Parts" then 
      totalpartinv = totalpartinv + rs("invd_subtotal")
	  totalpartcost = totalpartcost + rs("invd_itemcost")
%>
<!-- Add grand total average cost-->
<%
	   totalaverageecost = totalaverageecost + rs("md_averageecost")
	   totalaveragee = totalaveragee + rs("md_averagee")
   elseif rs("invd_parttype") = "Labour" then 
      totallabourinv = totallabourinv + rs("invd_subtotal")
   elseif rs("invd_parttype") = "Transport" then 
      totalTransportinv = totalTransportinv + rs("invd_subtotal")
   end if

TotalGrandQty =  TotalGrandQty +  rs("invd_qty")
TotalGrandInv =  totalpartinv +  totallabourinv + totalTransportinv
TotalGrandGST = TotalGrandInv * 0.0566037735849057
totalpartSales = totalpartinv * 0.9433962264150943
totalLabourSales = totalLabourinv * 0.9433962264150943
totalTransportSales = totalTransportinv * 0.9433962264150943

rs.movenext
wend
rs.close

i = 1	
%>
<!--sql2 = "SELECT  " & _ 
	"tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address,  " & _
	"tblinvoice.inv_cust_postcode, tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id,  " & _
	"tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, tblinvoice.inv_createddate,  " & _
	"tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt,  " & _
	"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate,  " & _
	"tblinvoice.inv_status, tblinvoice.inv_approvedby, tblinvoice.inv_approveddate, tblinvoice.inv_remark, tblinvoice.inv_postedby, tblinvoice.inv_posteddate, " & _
	"tblinvoice_detail.invd_id, tblinvoice_detail.invd_inv_no, tblinvoice_detail.invd_job_code,  " & _
	"tblinvoice_detail.invd_partcode, tblinvoice_detail.invd_parttype, tblinvoice_detail.invd_desc,  " & _
	"tblinvoice_detail.invd_unitcost, tblinvoice_detail.invd_qty, tblinvoice_detail.invd_discountamt,  " & _
	"tblinvoice_detail.invd_discounttype, tblinvoice_detail.invd_netcost, tblinvoice_detail.invd_subtotal, " & _
	"tblinvoice.inv_date, tblinvoice.inv_status, tblinvoice_detail.invd_itemcost " & _
	"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
	"where tblinvoice_detail.invd_id is not null " & _
	"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
-->
<%	
	'sql2 = "SELECT  " & _ 
	'"tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address,  " & _
	'"tblinvoice.inv_cust_postcode, tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id,  " & _
	'"tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, tblinvoice.inv_createddate,  " & _
	'"tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt,  " & _
	'"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate,  " & _
	'"tblinvoice.inv_status, tblinvoice.inv_approvedby, tblinvoice.inv_approveddate, tblinvoice.inv_remark, tblinvoice.inv_postedby, tblinvoice.inv_posteddate, " & _
	'"tblinvoice_detail.invd_id, tblinvoice_detail.invd_inv_no, tblinvoice_detail.invd_job_code,  " & _
	'"tblinvoice_detail.invd_partcode, tblinvoice_detail.invd_parttype, tblinvoice_detail.invd_desc,  " & _
	'"tblinvoice_detail.invd_unitcost, tblinvoice_detail.invd_qty, tblinvoice_detail.invd_discountamt,  " & _
	'"tblinvoice_detail.invd_discounttype, tblinvoice_detail.invd_netcost, tblinvoice_detail.invd_subtotal, " & _
	'"tblinvoice.inv_date, tblinvoice.inv_status, tblinvoice_detail.invd_itemcost, tblmodel.md_averageecost " & _
	'"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
	'"left join tblmodel on tblmodel.md_code = REPLACE(LTRIM(RTRIM(tblinvoice_detail.invd_partcode)), CHAR(9), '') " & _
	'"where tblinvoice_detail.invd_id is not null " & _
	'"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "

     sql2 = "SELECT  " & _ 
	"tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address,  " & _
	"tblinvoice.inv_cust_postcode, tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id,  " & _
	"tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, tblinvoice.inv_createddate,  " & _
	"tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt,  " & _
	"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate,  " & _
	"tblinvoice.inv_status, tblinvoice.inv_approvedby, tblinvoice.inv_approveddate, tblinvoice.inv_remark, tblinvoice.inv_postedby, tblinvoice.inv_posteddate, " & _
	"tblinvoice_detail.invd_id, tblinvoice_detail.invd_inv_no, tblinvoice_detail.invd_job_code,  " & _
	"tblinvoice_detail.invd_partcode, tblinvoice_detail.invd_parttype, tblinvoice_detail.invd_desc,  " & _
	"tblinvoice_detail.invd_unitcost, tblinvoice_detail.invd_qty, tblinvoice_detail.invd_discountamt,  " & _
	"tblinvoice_detail.invd_discounttype, tblinvoice_detail.invd_netcost, tblinvoice_detail.invd_subtotal, " & _
	"tblinvoice.inv_date, tblinvoice.inv_status, tblinvoice_detail.invd_itemcost,(select TOP 1 md_averagecost from tblmodel_avgcost where md_code=REPLACE(LTRIM(RTRIM(tblinvoice_detail.invd_partcode)), CHAR(9), '') and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_to & "' " & _ 
    "order by md_date desc)  AS md_averageecost " & _
	"FROM tblinvoice_detail inner join tblinvoice on tblinvoice_detail.invd_inv_no = tblinvoice.inv_no " & _
	"left join tblmodel on tblmodel.md_code = REPLACE(LTRIM(RTRIM(tblinvoice_detail.invd_partcode)), CHAR(9), '') " & _
	"where tblinvoice_detail.invd_id is not null " & _
	"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if

	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
       sql2 = sql2 & " order by tblinvoice.inv_no  "

response.Cookies("GAPS")("sqlexcel") = sql2
'response.write sql2
'response.End()
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   norecord = "There is no record found."
end if

If Not rs.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed

link = "&job_tech_type=" & job_tech_type & "&Searchor_date=" & Searchor_date & "&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&inv_status=" & inv_status   
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Sales Invoice Detail (Cross Month)</div></td>
                      </tr>
                    </table></td>
                </tr> 
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_salesanalysis_detail_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_salesanalysis_detail.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Invoice Date<br />
                        </strong></td>
                        <td colspan="3"><div align="left"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Type </td>
                        <td width="14%"><span class="titlegrey1">Technician</span></td>
                        <td width="24%" align="center"><span class="titlegrey1">Invoice Status</span></td>
                        <td rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_tech_type" id="job_tech_type">
                          <option value="">All</option>
                          <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select></td>
                        <td width="14%"><span class="titlegrey1">
                          <select name="job_tech_code" size="6" multiple="multiple" id="job_tech_code">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql1 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if checkTechlList(rs1("tech_code")) then
					  response.write "<option value='" & rs1("tech_code") & "' selected>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs1("tech_code") & "'>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
                          </select>
                        </span></td>
                        <td width="24%" align="center" valign="top">
                          <select name="inv_status" id="inv_status">
                            <option value="All">All</option>
                            <option value="Open" <%if inv_status="Open" then response.write " selected"%>>Open</option>
                            <option value="Submitted" <%if inv_status="Submitted" then response.write " selected"%>>Submitted</option>
                            <option value="Posted" <%if inv_status="Posted" then response.write " selected"%>>Posted</option>
                            <option value="CN" <%if inv_status="CN" then response.write " selected"%>>CN</option>
                          </select>
                       </td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_salesanalysis_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Date</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer Code<br />
                        </span></strong></font><font color="#FFFFFF"><strong><span><br />
                      </span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Type</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong> Inv<br />
                      </strong></font><font color="#FFFFFF"><strong>Qty</strong></font></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Inv<br /> 
                      Amt</span></strong></font></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Gst<br />
                      Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part <br />
                        Sales Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Labour</font><font color="#FFFFFF"> <br />
                        Amt</font></strong></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Transport <br />
                       Amt</font></strong></td>
                      <!--<td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part Cost</font></strong></td>-->
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price (Per item)</font></strong></td>					  
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parts Gross Profit</span></strong></font></td>
                      <td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GP %</span></strong></font></td>
                    </tr>
                    
<%

if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center" nowrap="nowrap"><%=j%>.</td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_invoice_new.asp?inv_no=<%=rs("inv_no")%>" target="_blank"><%=rs("inv_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("inv_status"))%></td>
                      <td align="left"><%=rs("inv_cust_code")%></td>
                      <td align="left"><%=rs("inv_cust_name")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("invd_partcode")%></td>
                      <td align="left"><%=rs("invd_parttype")%></td>
                      <td align="left"><%=rs("invd_desc")%></td>
                      <td align="right" nowrap="nowrap"><strong> <%=rs("invd_qty")%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong>0.00</strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Labour" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="right"><strong><%if rs("invd_parttype") = "Transport" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal"))%></strong></td>
                      <td align="center">
					  <!--ct check 25/3/21-->
					  <!--
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("invd_itemcost")*rs("invd_qty"))%>
                        <%end if%>
                      </strong>
					  -->
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("md_averageecost"))%>
                        <%end if%>
                      </strong>
					  </td>
					  <td align="center">
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(rs("md_averageecost")*rs("invd_qty"))%>
                        <%end if%>
                      </strong>
					  </td>
                      <td align="right"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then response.write chknumber2(rs("invd_subtotal")-(rs("invd_itemcost")*rs("invd_qty")))%>
                        <%end if%>
                      </strong></td>
                      <td align="right" nowrap="nowrap"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                         <%if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
						      if rs("invd_subtotal") > 0 then 
						      response.write chknumber2((rs("invd_subtotal")-((rs("invd_itemcost")*rs("invd_qty"))))/(rs("invd_subtotal")*100))
						      end if
						   else
						      response.write "0.00"
						   end if	  
							  %>
                        <%end if%>
                      %</strong></td>
                    </tr>
<%
invd_qty = invd_qty + ccur(rs("invd_qty"))
invd_subtotal = invd_subtotal + ccur(rs("invd_subtotal"))
inv_gstAmt = inv_gstAmt + 0

if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
invd_parts = invd_parts + ccur(rs("invd_subtotal"))
end if

if rs("invd_parttype") = "Labour" and rs("invd_subtotal") > 0 then 
invd_Labour = invd_Labour + ccur(rs("invd_subtotal"))
end if

if rs("invd_parttype") = "Transport" and rs("invd_subtotal") > 0 then 
invd_Transport = invd_Transport + ccur(rs("invd_subtotal"))
end if
%>
<!--invd_costtotal = invd_costtotal + ccur(rs("invd_itemcost")*rs("invd_qty"))-->
<%

if not ISNULL(rs("md_averageecost")) then 
invd_costtotal = invd_costtotal + ccur(rs("md_averageecost")*rs("invd_qty"))
invd_costtotala = invd_costtotala + ccur(rs("md_averageecost"))
end if

if rs("invd_parttype") = "Parts" and rs("invd_subtotal") > 0 then 
invd_grossamt = invd_grossamt + ccur(rs("invd_subtotal")-(rs("invd_itemcost")*rs("invd_qty")))
end if

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>

             <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Date</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer Code<br />
                        </span></strong></font><font color="#FFFFFF"><strong><span><br />
                      </span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Type</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong> Inv<br />
                      </strong></font><font color="#FFFFFF"><strong>Qty</strong></font></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Inv<br /> 
                      Amt</span></strong></font></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Gst<br />
                      Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part <br />
                        Sales Amt</font></strong></td>
                      <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Labour</font><font color="#FFFFFF"> <br />
                        Amt</font></strong></td>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Transport <br />
                       Amt</font></strong></td>
                      <!--<td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Part Cost</font></strong></td>-->					  
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price (Per Item)</font></strong></td>
					  <td align="center" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Average Cost Price</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parts Gross Profit</span></strong></font></td>
                      <td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GP %</span></strong></font></td>
                    </tr>         
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="9" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=(invd_qty)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_subtotal)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(inv_gstAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_parts)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(invd_Labour)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(invd_Transport)%></strong></td>
                      <td align="right" bgcolor="#999999">
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_costtotala)%>
                        <%end if%>
                      </strong>
					  </td>
					   <td align="right" bgcolor="#999999">
					  <!--ct check 25/3/21-->
					  <!--
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_costtotal)%>
                        <%end if%>
                      </strong>
					  -->
					  <strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_costtotal)%>
                        <%end if%>
                      </strong>
					  </td>
                      <td align="right" bgcolor="#999999"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%=chknumber2(invd_grossamt)%>
                        <%end if%>
                      </strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong>
                        <%if request.Cookies("GAPS")("view_cost")="Y" then %>
                        <%
						if  invd_parts > 0 then 
						    response.write chknumber2((invd_grossamt/invd_parts)*100)
						else
						    response.write "0.00"
						end if
						%>
                        <%end if%>
                      %</strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="9" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=(TotalGrandQty)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(TotalGrandInv)%></strong></td>
                      <!--<td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(TotalGrandGST)%></strong></td>-->
                      <td height="40" align="right" bgcolor="#999999"><strong>0.00</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(totalpartSales)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(totalLabourSales)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalTransportSales)%></strong></td>
                      <!--<td align="right" bgcolor="#999999"><strong><%=chknumber2(totalpartcost)%></strong></td>-->
					  <td align="right" bgcolor="#999999"><strong><%=chknumber2(totalaveragee)%></strong></td>
					  <td align="right" bgcolor="#999999"><strong><%=chknumber2(totalaverageecost)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(totalpartSales-totalpartcost)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%
					  if totalpartSales > 0 then 
					     response.write chknumber2(((totalpartSales-totalpartcost)/totalpartSales)*100)
					  else
					     response.write "0.00"
					  end if
					  %>%</strong></td>
                    </tr>                    
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_salesanalysis_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->