var win;
function popup(theURL,winName,features) {
  if (win && win.open && !win.closed) win.close();
  win = window.open(theURL,winName,'top=0,left=0,toolbar=Yes,location=Yes,status=Yes,menubar=Yes,' + features);
}

function openPopup(myLink)
  {
  if(! window.focus)return;
  var myWin=window.open(myLink,"link","toolbar=1,status=1,menubar=1,scrollbars=1,resizable=1,width=700,height=550");
  myWin.focus();
}