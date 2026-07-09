<!-- #include file="database/datastore.asp" -->

<%
stateid=request.querystring("stateid")
cityid=request.querystring("cityid")
postcode=request.querystring("postcode")
areatext=request.querystring("areatext")

if areatext <> "" then
        fullareapostcode = Split(areatext,"-") 'splits area and postcode 
        selectedarea = fullareapostcode(0)
        'selectedpostcode = fullareapostcode(1)
        Areaname=selectedarea 'copy the selected text into the input text as well in case user fully typed manually
end if
'response.write areatext
%>

<html>
<head>
<script>
function showHint(str) {
    
    var sid = document.getElementById("stateid").value;
    var cid = document.getElementById("cityid").value;
    var pid = document.getElementById("postcode").value;
    
    if (str.length == 0) {
        document.getElementById("txtHint").innerHTML = "";
        return;
    } else {
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (this.readyState == 4 && this.status == 200) {
                document.getElementById("txtHint").innerHTML = this.responseText;
            }
        };

        var url="getArea.asp";
        url+="?q="+str;
        url+="&s="+sid;
        url+="&c="+cid;
        url+="&p="+pid;
    
        xmlhttp.open("GET", url, true);
        xmlhttp.send();
    }
}
</script>

<script language="javascript">
    function AcceptArea()
    {
        var area = document.getElementById('Areaname').value;
        //var postcode = document.getElementById('selectedpostcode').value;
        //alert(area);
        var existext = parent.opener.document.forms['formorder'].job_cust_address.value;
        existext = existext + " , " + area;
        parent.opener.document.forms['formorder'].job_cust_address.value = existext;
        //parent.opener.document.forms['formorder'].cust_postcode.value = postcode;
        parent.window.close();
    }
// -->
</script>

</head>
<body>

<p><b>Search an area name below:</b></p>
<form name="myform">
Area Name (eg Jalan/Taman/etc): 
<input type="text" name="Areaname" id="Areaname" value="<%=Areaname%>" onkeyup="showHint(this.value)" size="100" maxlength="100">
<input type="Submit" name="Submit1" id="Submit1" value="Submit" onClick="AcceptArea()" /> 
<input type="hidden" name="selectedarea" id="selectedarea" value="<%=selectedarea%>" /> 
<input type="hidden" name="selectedpostcode" id="selectedpostcode" value="<%=selectedpostcode%>" /> 
<input type="hidden" name="stateid" id="stateid" value="<%=stateid%>" /> 
<input type="hidden" name="cityid" id="cityid" value="<%=cityid%>" /> 
<input type="hidden" name="postcode" id="postcode" value="<%=postcode%>" /> 
</form>
<p>Suggestions: <span id="txtHint"></span></p>
 
</body>
</html>

