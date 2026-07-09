<!-- #include file="database/datastore.asp" -->

<script language="javascript">

    function grabarea()
    {
        alert('Hi');
        //document.getElementById('cust_name').value = s;
        //document.getElementById('cust_city_id').value = c;
        //document.formorder.submit();
    }

// -->
</script>

<%
response.expires=-1

'count=500
dim a(500)
dim b(500) 
'Fill up array with names

'get the q parameter from URL
q=ucase(request.querystring("q"))
p=request.querystring("p") 'postcode para
c=request.querystring("c") 'city not used

cnt=1
'sql = "select postcode,area_name from tblpostcode where post_office = '" & cityname & "'"
sql = "select postcode,area_name from tblpostcode where postcode = '" & p & "'"
set rs = server.CreateObject("adodb.recordset")
rs.Open sql,strconnect,0,1,&H0001
while Not rs.EOF and cnt<=500
			a(cnt) = rs("area_name") 'Fill up array with areanames
            b(cnt) = rs("postcode")
            cnt=cnt+1
rs.MoveNext
wend
rs.close

'lookup all hints from array if length of q>0
if len(q)>0 then
  hint=""
  for i=1 to 500
    if q=ucase(mid(a(i),1,len(q))) then
      if hint="" then    
         response.write("<a href='rm_new_address.asp?areatext=" & a(i) & "-" & b(i) & "'>"&a(i)& " , " & "</a>")       
      else        
         response.write("<a href='rm_new_address.asp?areatext=" & a(i) & "-" & b(i) & "'>"&a(i)& " , " & "</a>")     
      end if
    end if
  next
end if    

'Output "no suggestion" if no hint were found
'or output the correct values
if hint="" then
 ' response.write("no suggestion")
else
   ' response.write("<a href='getareas.asp?areatext=>" & a(i) & " - " & b(i) & "'></a>")
  response.write(hint)
end if
%>