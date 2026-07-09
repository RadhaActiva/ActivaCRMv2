<title>Riegen Markerting CRM</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link href="inc/gaps.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="inc/popup.js"></script>
<script language="JavaScript">
function confirmDel(id,url) {
if (confirm("Are you sure you want to delete ID: "+id)) {
    location.href=url;
  }
}

function confirmDuplicateQuote(id,url) {
if (confirm("Are you sure you want to Duplicate " + id + " to New Quotation?")) {
    location.href=url;
  }
}

function confirmDelete(url) {
if (confirm("Are you sure you want to delete this record?")) {
    location.href=url;
  }
}

function confirmAction(id,url) {
if (confirm("Are you sure you want to Confirm this: "+id)) {
    location.href=url;
  }
}

function checkCR(evt) {
	var evt  = (evt) ? evt : ((event) ? event : null);
	var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
	if ((evt.keyCode == 13) && (node.type=="text")) {return false;}
}
document.onkeypress = checkCR;
</script>