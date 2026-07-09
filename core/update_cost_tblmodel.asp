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
      
    sql1= "select md_code, md_costprice from tblmodel"
    set rs1 = server.CreateObject("adodb.recordset")
    rs1.ActiveConnection = strconnect
    rs1.Source = sql1
    rs1.CursorLocation  = 3
    rs1.Open
   
  while Not rs1.EOF
    weighted_avg_cost = 0
    searchvalue = rs1("md_code")
    
        sql2= "update tblmodel set md_averageecost = (select TOP 1 md_averagecost from tblmodel_avgcost where md_code = '" & rs1("md_code") & "' order by md_date desc) where md_code = '" & rs1("md_code") & "'"
        set rs2 = server.CreateObject("adodb.recordset")
        rs2.ActiveConnection = strconnect
        rs2.Source = sql2
        rs2.CursorLocation  = 3
        rs2.Open  
    
  rs1.movenext
  wend
%>

