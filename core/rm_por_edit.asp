<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            height: 25px;
        }
        .auto-style2 {
            width: 163px;
        }
        .auto-style3 {
            height: 25px;
            width: 163px;
        }
        .auto-style5 {
            height: 22px;
        }
        .auto-style6 {
            width: 66%;
        }
    </style>
</head>
<%

 editmode= ""
 por_docno = request("por_docno")  'doc-no retained when form is triggerd from action.asp(outside the page)
 editmode = request("editmode")
 dispavg = request("dispavg")
 recalc = request("recalc")
 porreport = request("porreport")
     
 if por_docno = "" then 
    por_docno = request.form("por_docno") 'doc-no retained when form is refreshed
 end if

 por_id = request.form("por_id")  
 partcode = request.form("partcode")
 partdesc = request.form("partdesc")
 por_date = request.form("por_date")
 mth1 = request.form("mth1")
 mth2 = request.form("mth2")
 mth3 = request.form("mth3")
 mth4 = request.form("mth4")
 mth5 = request.form("mth5")
 mth6 = request.form("mth6")
 eta1 = request.form("eta1")   
 eta2 = request.form("eta2")

 avgmth1 = request.form("avgmth1")
 avgmth2 = request.form("avgmth2")
 avgmth3 = request.form("avgmth3")
 avgmth4 = request.form("avgmth4")
 avgmth5 = request.form("avgmth5")
 avgmth6 = request.form("avgmth6")

 avgfor3mth = request.form("avgfor3mth")
 avgfor6mth = request.form("avgfor6mth")
    

 orderqty1 = request.form("orderqty1")
 orderqty2 = request.form("orderqty2")
 totalincomingstk = request.form("totalincomingstk")
 laststockin = request.form("laststockin")
 stklastformonths = request.form("stklastformonths")
 sw1_qty = request.form("sw1_qty")
 ex_qty = request.form("ex_qty")

if por_docno <> "" AND partcode  <> "" and mth1 <> "" and dispavg="Yes" then      

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -1, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -1, por_date)) & "'"
            avgmth1 = selectid(sql)  

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -2, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -2, por_date)) & "'"
            avgmth2 = selectid(sql)  

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -3, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -3, por_date)) & "'"
            avgmth3 = selectid(sql)  

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -4, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -4, por_date)) & "'"
            avgmth4 = selectid(sql)  

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -5, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -5, por_date)) & "'"
            avgmth5 = selectid(sql)  

            sql =  "select (sum(stk_qty) * -1) from tblstocktran where stk_type in ('Stock-Out', 'Stock-Transfer-Out', 'DO') and stk_fromwarehouse='W1' and stk_itm_code = '" & partcode & "' " & _
            "and month(stk_date)= '" & Month(DateAdd("m", -6, por_date)) & "' and year (stk_date) = '" & Year(DateAdd("m", -6, por_date)) & "'"
            avgmth6 = selectid(sql)  

           ' sql = "select (sum(stk_qty)) from tblstocktran where stk_type in ('Stock-In') and stk_itm_code = '" & partcode & "' and  stk_date >=dateadd(month,datediff(month,0,'" & por_date & "')-12,0)"
            sql = "select (sum(stk_qty)) from tblstocktran where stk_itm_code = '" & partcode & "' and stk_reference ='W1'" 'get the current qty
            totalincomingstk = selectid(sql)
            sw1_qty = selectid(sql)
            
            sql = "select CONCAT(stk_qty,' / ', FORMAT (stk_date,'dd-MM-yyyy '))   from tblstocktran where stk_type in ('Stock-In') and stk_itm_code = '" & partcode & "' order by stk_id desc"
            laststockin = selectid(sql)

            avgfor3mth = (ChkNumberInt(avgmth1)+ChkNumberInt(avgmth2)+ChkNumberInt(avgmth3))/3
            avgfor6mth = (ChkNumberInt(avgmth1)+ChkNumberInt(avgmth2)+ChkNumberInt(avgmth3)+ChkNumberInt(avgmth4)+ChkNumberInt(avgmth5)+ChkNumberInt(avgmth6))/6            
            dispavg=""
            stype = "addPOR"
else    
       if por_docno = "" then
            por_docno = "POP"  &  Right("0" & Day(Now),2)  & Right("0" & Month(Now),2) & Year(Now)
            por_date = Date()
            mth1 = MonthName(Month(DateAdd("m", -1, por_date)),True) & "-" & Year(DateAdd("m", -1, por_date))
            mth2 = MonthName(Month(DateAdd("m", -2, por_date)),True) & "-" & Year(DateAdd("m", -2, por_date))
            mth3 = MonthName(Month(DateAdd("m", -3, por_date)),True) & "-" & Year(DateAdd("m", -3, por_date))
            mth4 = MonthName(Month(DateAdd("m", -4, por_date)),True) & "-" & Year(DateAdd("m", -4, por_date))
            mth5 = MonthName(Month(DateAdd("m", -5, por_date)),True) & "-" & Year(DateAdd("m", -5, por_date))
            mth6 = MonthName(Month(DateAdd("m", -6, por_date)),True) & "-" & Year(DateAdd("m", -6, por_date))   
            stype = "addPOR"
        end if
end if

if request("por_docno") <> "" and partcode  <> "" and mth1 <> "" and recalc="Yes" then     

    sql = "select (sum(stk_qty)) from tblstocktran where stk_itm_code = '" & partcode & "' and stk_reference ='W1'" 'get the current qty
    totalincomingstk = selectid(sql)
    sw1_qty = selectid(sql)

    newtotal=0
    newtotal = totalincomingstk + ChkNumberInt(orderqty1) + ChkNumberInt(orderqty2) + ChkNumberInt(ex_qty)
    
    totalincomingstk = newtotal        
    
    if  ChkNumberInt(avgfor3mth) >= 1 then 
        stklastformonths = totalincomingstk / round(avgfor3mth,2)
    else
        stklastformonths = 0
    end if
    recalc=""
    stype = "addPOR"
end if

'invoke if record triggered from report or edit record    
if (request("por_docno") <> "" and request("por_id") <> "" and porreport="Yes")  OR (request("por_docno") <> "" and request("por_id") <> "" and editmode = "Yes") then
        set rs3 = server.CreateObject("adodb.recordset")
        sql1="SELECT por_id,por_docno,por_date,por_remark,por_part_code,por_eta1,por_order_qty1,por_eta2,por_order_qty2,por_last_stockin,por_total_incoming" & _
	            ",por_total_last,por_avg_3,por_avg_6,por_mth1,por_mth2,por_mth3,por_mth4,por_mth5,por_mth6,por_mth1_qty,por_mth2_qty,por_mth3_qty,por_mth4_qty" & _
	            ",por_mth5_qty,por_mth6_qty,por_createdby,por_createddate,por_sw1_qty,por_ex_qty FROM tblpor where por_docno = '" & request("por_docno") &"' and por_id = '" & request("por_id") &"'" 
        rs3.ActiveConnection = strconnect
        rs3.Source = sql1
        rs3.CursorLocation  = 3
        rs3.Open
 
        If Not rs3.EOF Then
            por_id = rs3("por_id") 
            por_docno = rs3("por_docno") 
			por_date = rs3("por_date")
			por_remark= rs3("por_remark")
			partcode= rs3("por_part_code")

            'part description fetched form tblmodel
            partdesc = ""
            sql = "select md_desc from tblmodel where md_code = '" & rs3("por_part_code") & "'  "
            partdesc = selectid(sql)   

			eta1 = rs3("por_eta1")
			orderqty1 = rs3("por_order_qty1")
			eta2 = 	rs3("por_eta2")    
    		orderqty2 = rs3("por_order_qty2")
			laststockin = rs3("por_last_stockin")
			totalincomingstk = rs3("por_total_incoming")
			stklastformonths = rs3("por_total_last")
			avgfor3mth = rs3("por_avg_3")
			avgfor6mth = rs3("por_avg_6")
			mth1 = rs3("por_mth1")
			mth2 = rs3("por_mth2")
			mth3 = rs3("por_mth3")
			mth4 = rs3("por_mth4")
			mth5 = rs3("por_mth5")
			mth6 = rs3("por_mth6")
			avgmth1 = rs3("por_mth1_qty")
            avgmth2 = rs3("por_mth2_qty")
            avgmth3 = rs3("por_mth3_qty")
            avgmth4 = rs3("por_mth4_qty")
            avgmth5 = rs3("por_mth5_qty")
            avgmth6 = rs3("por_mth6_qty")
            por_createdby  =rs3("por_createdby")
            por_createddate =rs3("por_createddate")
            sw1_qty = rs3("por_sw1_qty")
            ex_qty = rs3("por_ex_qty")
            rs3.Close            
            editmode = "No"
            stype = "addPOR"       
      
    end if 
end if
    
%> 
<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{
  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorder.action = orderlinks;
	document.formorder.submit();
   }
}

function pageReload()
{
    //alert("hello");
    //document.formorder.submit();
    //window.location.reload();
    //Response.redirect("rm_por_edit.asp?partcode=" & partcode)
}
// -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Create/Edit </font>Parts Order Plan</div></td>
                          <td width="20%" align="center" class="titlegrey1">                         
                            <a href="rm_por_rpt_excel.asp?por_docno=<%=por_docno%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a>                       
                     </td>                        
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" >
                    <thead>
                      <tr class="head_row">
                        <td colspan="2" nowrap="nowrap" bgcolor="#DADADA"><strong>Stock Information </strong></td>
                      </tr>
                    </thead>
                    <form name="formorder" id="formorder" method="post" action="action.asp?type=<%=stype%>" >
                      <tr class="head_row">
                        <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
                          <tr>
                            <td align="left" bgcolor="#CD6155" width="30%"><font color="#FFFFFF"><strong>POP Doc No</strong><br />
                             <strong><font size="1">(System Generate) </font></strong></font></td>
                            <td nowrap="nowrap" class="auto-style2"><strong><input name="por_docno" type="text" style="background-color:#EBEDEF" value="<%=por_docno%>" readonly/></strong></td>
                            <td height="22" nowrap="nowrap" width="30%" bgcolor="#CD6155"><font color="#FFFFFF">POP Date</font></td>
                            <td><input name="por_date" type="text" id="por_date" style="background-color:#EBEDEF" value="<%=chkdate(por_date)%>" size="15" maxlength="20" readonly/><input name="por_id" type="hidden" value="<%=por_id%>" /></td>                            
                          </tr>
                          
                          <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
                            <td nowrap="nowrap"><%=por_createdby%> @ <%=chkdatetime(por_createddate)%></td>                            
                            </tr>
                            <tr>
                             <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF">Remark</font></td>
                            <td colspan="3"><textarea name="por_remark" rows="2" cols="50" class="text" id="por_remark"><%=por_remark%></textarea></td>
                            </tr>                      
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Spare Part Code</strong></font></strong></td>
                              <td>
                            <input name="partcode" type="text" id="partcode" value="<%=partcode%>" size="15" maxlength="15" />
                              [<a href="javascript:popup('rm_spareparts_list.asp?searchitem=md_type&amp;searchvalue=<%=cust_code%>&amp;formname=formorder&fieldname=partcode&fieldname1=partdesc','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td> 
                              </td><td><input type="submit" style="height:40px;width:160px;font-weight:bold;background-color: #1b6bcf; 
        color: white;" name="button" id="button" value="Display Average" formaction='rm_por_edit.asp?por_docno=<%=por_docno%>&dispavg=Yes' formmethod="post"/></td>
                            </tr>   
                      <tr>
                          <td bgcolor="#CD6155"><strong><font color="#FFFFFF"><strong>Part Description</strong></font></strong></td>
                          <td><textarea name="partdesc" rows="2" cols="50" class="text" id="textarea" readonly><%=partdesc%></textarea></td>
                      </tr>
                            <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155" class="auto-style1"> <font color="#FFFFFF" >1st Part Request  (ETA/ETS)</font></td>
                            <td nowrap="nowrap"  class="auto-style1"><input type="text" name="eta1" value="<%=ChkDate(eta1)%>" />
                                <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.eta1;
                                calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                            </td>
                            <td bgcolor="#CD6155" class="auto-style1"> <font color="#FFFFFF">Order Qty</font></td>
                            <td colspan="1" align="left" class="auto-style1"><input type="text" style="width: 80px;" maxlength=6 name="orderqty1" value="<%=orderqty1%>" /></td>
                            </tr>
                             <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"> <font color="#FFFFFF" class="auto-style6">2nd Part Request  (ETA/ETS)</font></td>                         
                            <td nowrap="nowrap"  class="auto-style1"><input type="text" name="eta2" value="<%=ChkDate(eta2)%>" /> 
                                <font color="#000000"><strong><a href="javascript:void(null)" onclick="window.dateField = document.formorder.eta2;
                                calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td></td>
                            <td bgcolor="#CD6155"><font color="#FFFFFF">Order Qty</font></td>
                            <td colspan="1" align="left"><input type="text"style="width: 80px;" maxlength=6 name="orderqty2" value="<%=orderqty2%>"  /></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"> <font color="#FFFFFF">SW1 Qty</font></td>
                            <td nowrap="nowrap" class="auto-style3"><input type="text" style="background-color:#EBEDEF ;width: 80px;" name="sw1_qty" value="<%=sw1_qty%>" readonly/>&nbsp;&nbsp;</td>
                            <td nowrap="nowrap"  bgcolor="#CD6155"> <font color="#FFFFFF">Exchange Qty</font></td>
                            <td><input type="text" style="width: 80px;" name="ex_qty" maxlength=6 value="<%=ex_qty%>"/></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap"  bgcolor="#CD6155" class="auto-style1"><font color="#FFFFFF">Last Stock-In</font></td>
                            <td nowrap="nowrap" class="auto-style3"><input type="text" style="background-color:#EBEDEF;width: 120px;" name="laststockin" value="<%=laststockin%>" readonly /></td>
                            <td nowrap="nowrap"  bgcolor="#CD6155" class="auto-style1"> <font color="#FFFFFF">Total Incoming Stock</font></td>
                            <td class="auto-style1"><input type="text" style="background-color:#EBEDEF;width: 80px;" name="totalincomingstk" value="<%=round(ChkNumberInt(totalincomingstk),2)%>" readonly/></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap"  bgcolor="#CD6155"> <font color="#FFFFFF">Stock last for X mths</font></td>
                            <td nowrap="nowrap" class="auto-style3"><input type="text" style="background-color:#EBEDEF;width: 80px;" name="stklastformonths" value="<%=ChkNumber2Decimal(stklastformonths)%>" readonly />&nbsp[based on 3-mth avg]</td> 
                            <td></td><td></td>
                            <td><input type="submit" style="height:40px;width:160px;font-weight:bold;background-color: #1b6bcf;color: white;" name="calc" id="calc" value="ReCalculate" formaction='rm_por_edit.asp?por_docno=<%=por_docno%>&recalc=Yes' formmethod="post"/></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap" bgcolor="#CD6155"> <font color="#FFFFFF" >3-month Avg</font></td>
                            <td nowrap="nowrap" class="auto-style3"><input type="text" style="background-color:#EBEDEF;width: 80px;" name="avgfor3mth" value="<%=ChkNumber2Decimal(avgfor3mth)%>" readonly/></td>
                            <td nowrap="nowrap"  bgcolor="#CD6155"> <font color="#FFFFFF">6-month Avg</font></td>
                            <td><input type="text" style="background-color:#EBEDEF;width: 80px;" name="avgfor6mth" value="<%=ChkNumber2Decimal(avgfor6mth)%>" readonly/></td>
                             <td><input type="submit" style="height:40px;width:160px;font-weight:bold;background-color: #1b6bcf; color: white;" name="btnSave" id="btnSave" value="Save"/></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap" class="auto-style6"></td>
                            </tr>
                            <tr>
                            <td nowrap="nowrap"><Strong>Monthly Usage</Strong></td>
                            </tr>
                        </table>
                          <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">                            
                              <tr>
                                  <td class="auto-style5"><input type="text" style="background-color:#AED6F1; text-align:center" align="center" name="mth1" value="<%=mth1%>" readonly></td>
                                  <td class="auto-style5"><input type="text"  style="background-color:#AED6F1; text-align:center" align="center" name="mth2" value="<%=mth2%>" readonly></td>
                                  <td class="auto-style5"><input type="text"  style="background-color:#AED6F1; text-align:center" align="center" name="mth3" value="<%=mth3%>" readonly></td>
                                  <td class="auto-style5"><input type="text"  style="background-color:#AED6F1; text-align:center" align="center" name="mth4" value="<%=mth4%>" readonly></td>
                                  <td class="auto-style5"><input type="text" style="background-color:#AED6F1; text-align:center" align="center" name="mth5" value="<%=mth5%>" readonly></td>
                                  <td class="auto-style5"><input type="text" style="background-color:#AED6F1; text-align:center" align="center" name="mth6" value="<%=mth6%>" readonly></td>
                              </tr>
                              <tr>
                                  <td><input type="text" name="avgmth1" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth1%>" readonly"></td>
                                  <td><input type="text" name="avgmth2" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth2%>" readonly></td>
                                  <td><input type="text" name="avgmth3" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth3%>" readonly></td>
                                  <td><input type="text" name="avgmth4" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth4%>" readonly></td>
                                  <td><input type="text" name="avgmth5" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth5%>" readonly></td>
                                  <td><input type="text" name="avgmth6" style="background-color:#EBEDEF; text-align:center" value="<%=avgmth6%>" readonly></td>
                              </tr>
                           </table> 

                        </td>
                      </tr>
                      <tr class="head_row">
                        <td colspan="2" valign="top">&nbsp;</td>
                      </tr>
                    </form>                  

 
                    <%if por_docno <> "" then %>                    
                    <tr class="head_row">
                      <td colspan="2"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E8E8E8">
                      <tr class="head_row">
                          <td height="24" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>No</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF">POP Doc No</font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Part Code</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Part Desc</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>ETA 1</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Order Qty</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>ETA 2</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Order Qty</strong></font></strong></td>
                          <td align="center" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Action</strong></font></strong></td>
                        </tr>       
                         
 <%				i = 1
				sql="SELECT por_id,por_docno,por_date,por_remark,por_part_code,por_eta1,por_order_qty1,por_eta2,por_order_qty2,por_last_stockin,por_total_incoming" & _
	            ",por_total_last,por_avg_3,por_avg_6,por_mth1,por_mth2,por_mth3,por_mth4,por_mth5,por_mth6,por_mth1_qty,por_mth2_qty,por_mth3_qty,por_mth4_qty" & _
	            ",por_mth5_qty,por_mth6_qty,por_createdby,por_createddate FROM tblpor where por_docno = '" & por_docno & "'" 	   
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql,strconnect,3,3,&H0001
                while Not rs1.EOF
     
     partdesc = ""
     sql = "select md_desc from tblmodel where md_code = '" & rs1("por_part_code") & "'  "
     partdesc = selectid(sql)       
     %>                         
                        <tr valign="top">
                          <td align="center"><%=i%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("por_docno")%></td>
                          <td bgcolor="#FFFFFF" style="text-align: left"><%=rs1("por_part_code")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=partdesc%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=chkdate(rs1("por_eta1"))%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("por_order_qty1")%></td>
                          <td align="center" bgcolor="#FFFFFF" class='tktTotals'><%=chkdate(rs1("por_eta2"))%></td>
                          <td align="right" bgcolor="#FFFFFF" class='tktTotals'><%=rs1("por_order_qty2")%></td>
                          <td align="center">
                             <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_por_edit.asp?por_id=<%=rs1("por_id")%>&por_docno=<%=rs1("por_docno")%>&editmode=Yes#POR'" />
                             <input type="button" name="button9" id="button23" value="Del" onclick="javascript:confirmDel('<%=rs1("por_id")%>','action.asp?type=delPOR&por_id=<%=rs1("por_id")%>&por_docno=<%=rs1("por_docno")%>')"/></td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>                    </table></td>
                      </tr>
                      </table></td>
                    </tr>
                    
                <%end if%>
       <!-- #include file="footer.asp" -->