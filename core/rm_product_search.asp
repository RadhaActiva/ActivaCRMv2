
<!-- #include file="database/datastore.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" type="text/css" href="scripts/style/general.css">
</head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen CRM</title>
<%
fromDate = now()
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="inc/gaps.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="inc/popup.js"></script>

<style>		
.leftMenu {	text-align: left; }		
.centerMenu { text-align: center;}		
.rightMenu { text-align: right;	}
		
a.MenuLabelLink	{ COLOR: #FFFFFF;	FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLink:link { COLOR: #FFFFFF;	FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLink:visited	{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None;	}
a.MenuLabelLink:hover{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuLabelLinkOn {	COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLink { COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLink:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLinkOn { COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
.myMenu { position: absolute; visibility: hidden; z-index: 5; }				
.myMenuLabelleft { padding: 15px 0px 15px 0px; text-align: center; }		
.myMenuLabelcenter { padding: 15px 0px 0px 0px; text-align: center; }		
.myMenuLabelright { padding: 0px 0px 0px 0px; text-align: right; }		
.myMenuItemleft { padding: 0px 0px 0px 0px; text-align: left; }		
.myMenuItemcenter { padding: 0px 0px 0px 0px; text-align: center; }		
.myMenuItemright { padding: 0px 0px 0px 0px; text-align: right; }		
		
#myTest { 
width: 800px;
padding: 0px 0px 0px 0px;
z-index: 1;
}
</style>
<script language="JavaScript1.2" src="inc/api.js" type="text/javascript">
<!-- 
-->
</script>
<script language="JavaScript1.2" src="inc/menucode.js" type="text/javascript">
<!-- 
-->
</script>

<script>
function popup(theURL,winName,features) {
  window.open(theURL,winName,'top=0,left=0,toolbar=yes,location=yes,status=yes,menubar=no,' + features);
}
</script>
<script>
var screenwidth = 850;
var screenHeight = 650;
</script>


<body>


<table width="99%" border="1" cellpadding="2" cellspacing="0" bordercolor="#CCFFCC">
  <tr bgcolor="#CCFFCC"> 
    <td colspan="11"><strong>Search / Filter Box</strong> </td>
  </tr>
   <form name="form1" action="searchboxStockInItems.asp?searchdata=yes" method="post">  
  <tr bgcolor="#CCFFCC"> 
    <td><strong><font color="#000000">Group</font></strong> </td>
    <td align="left" nowrap><strong> <font color="#000000">Code</font></strong> 
    </td>
    <td> <strong><font color="#000000">Description</font></strong></td>
    <td><strong><font color="#000000">Brand</font></strong></td>
    <td><strong><font color="#000000">Type</font></strong></td>
    <td><strong><font color="#000000">Model</font></strong></td>
    <td><strong><font color="#000000">Sub-Type1</font></strong></td>
    <td><strong><font color="#000000">Sub-Type2</font></strong></td>
    <td><strong><font color="#000000">Capacity</font></strong></td>
    <td valign="bottom" nowrap><strong><font color="#000000">Price Level</font></strong></td>
    <td rowspan="2" align="right" valign="bottom" nowrap><input type="submit" name="searchbox" value="Sch">
      <input type="button" name="Reset" id="Reset" value="R" onClick="document.location.href='searchboxPOItems.asp?ItemGroup=&searchdata=Reset'"></td>
  </tr>
  
 
  <tr> 
    <td><select name="ItemGroup" onChange="document.forms[0].submit();" style="{width:80px;}">
       <option value="">- All -</option>
        <option value="AC" >AC</option>
        <option value="AC Installation" >AC Installation</option>
        <option value="AC Material" >AC Material</option>
        <option value="ACS" >ACS</option>
        <option value="ASD" >ASD</option>
        <option value="Parts" >Parts</option>
        <option value="Service" >Service</option>
      </select></td>
    <td nowrap><strong> </strong><strong> 
      <input name="Code" type="text" size="10" value="">
      </strong></td>
    <td><strong> 
      <input name="Description" type="text" size="10" value="">
      </strong></td>
    <td><select name="Brand" onChange="document.forms[0].submit();" style="{width:80px;}">
	<option value="">- All -</option>
        <option value='0'>0</option><option value='ASD'>ASD</option><option value='CN'>CN</option><option value='Daikin'>Daikin</option><option value='DEWPOINT'>DEWPOINT</option><option value='DongYue'>DongYue</option><option value='eBoy'>eBoy</option><option value='FORANE'>FORANE</option><option value='FRESCO'>FRESCO</option><option value='GC'>GC</option><option value='InsulFlex'>InsulFlex</option><option value='K-Flex'>K-Flex</option><option value='KoolMan'>KoolMan</option><option value='KRUGER'>KRUGER</option><option value='LG'>LG</option><option value='LOCAL'>LOCAL</option><option value='ME'>ME</option><option value='Met Tube'>Met Tube</option><option value='Mitsubishi Electric'>Mitsubishi Electric</option><option value='MOX'>MOX</option><option value='PU'>PU</option><option value='RC'>RC</option><option value='Samsung'>Samsung</option><option value='Super Lon'>Super Lon</option><option value='Trane'>Trane</option><option value='WireMan'>WireMan</option><option value='York'>York</option>
      </select></td>
    <td><strong>
      <select name="ItemType" onChange="document.forms[0].submit();" style="{width:80px;}">
       <option value="">- All -</option>
        <option value='0'>0</option><option value='AIR CLEANER'>AIR CLEANER</option><option value='Bracket'>Bracket</option><option value='Conduit'>Conduit</option><option value='Copper Pipe'>Copper Pipe</option><option value='Customised'>Customised</option><option value='DU'>DU</option><option value='Elec Acc'>Elec Acc</option><option value='Gas'>Gas</option><option value='Insulation'>Insulation</option><option value='Other'>Other</option><option value='Pipe'>Pipe</option><option value='Project'>Project</option><option value='PT'>PT</option><option value='PU'>PU</option><option value='SM'>SM</option><option value='SM22'>SM22</option><option value='SM410'>SM410</option><option value='SS'>SS</option><option value='SS22'>SS22</option><option value='SS22i'>SS22i</option><option value='SS410'>SS410</option><option value='SS410i'>SS410i</option><option value='Trunking'>Trunking</option><option value='Ventilation Fan'>Ventilation Fan</option><option value='VRV'>VRV</option><option value='VRV22'>VRV22</option><option value='VRV410'>VRV410</option><option value='Wire '>Wire </option>
      </select>
      </strong></td>
    <td><strong>
      <input name="ModelParts" type="text" size="10" value="">
      </strong></td>
    <td><strong> 
      <select name="SubType1" onChange="document.forms[0].submit();" style="{width:100px;}">
        <option value="">- All -</option>
         <option value='0'>0</option><option value='AC'>AC</option><option value='All'>All</option><option value='Apartment'>Apartment</option><option value='CC'>CC</option><option value='CE'>CE</option><option value='Ceiling'>Ceiling</option><option value='CK'>CK</option><option value='CK, CE'>CK, CE</option><option value='CK, CE, Ducted'>CK, CE, Ducted</option><option value='CN'>CN</option><option value='Common'>Common</option><option value='Ducted'>Ducted</option><option value='FS'>FS</option><option value='Ipoh'>Ipoh</option><option value='Local'>Local</option><option value='MS'>MS</option><option value='OD'>OD</option><option value='Other'>Other</option><option value='Panel'>Panel</option><option value='PT'>PT</option><option value='PU'>PU</option><option value='PVC'>PVC</option><option value='PVC/I/PVC'>PVC/I/PVC</option><option value='RC'>RC</option><option value='Refrigerant'>Refrigerant</option><option value='SM'>SM</option><option value='Truncking'>Truncking</option><option value='Tube'>Tube</option><option value='Venti'>Venti</option><option value='VRV'>VRV</option><option value='Wall'>Wall</option><option value='WM'>WM</option><option value='WM, CK, CE, Ducted'>WM, CK, CE, Ducted</option>
      </select>
      </strong></td>
    <td><strong> 
      <select name="SubType2" onChange="document.forms[0].submit();" style="{width:100px;}">
        <option value="">- All -</option>
        <option value='0'>0</option><option value='0.51mm'>0.51mm</option><option value='0.56mm'>0.56mm</option><option value='0.61mm'>0.61mm</option><option value='0.71mm'>0.71mm</option><option value='0.81mm'>0.81mm</option><option value='0.91mm'>0.91mm</option><option value='1 1/4" (32mm)'>1 1/4" (32mm)</option><option value='1 Core'>1 Core</option><option value='1" (25mm)'>1" (25mm)</option><option value='1/2" (12mm)'>1/2" (12mm)</option><option value='1/2" (13mm)'>1/2" (13mm)</option><option value='1/4" (6mm)'>1/4" (6mm)</option><option value='3 Core'>3 Core</option><option value='3/4" (19mm)'>3/4" (19mm)</option><option value='3/8" (9mm)'>3/8" (9mm)</option><option value='4 Core'>4 Core</option><option value='6/8" (19mm)'>6/8" (19mm)</option><option value='6/8`'>6/8`</option><option value='Add'>Add</option><option value='Add-Basic'>Add-Basic</option><option value='Add-Drain c/w insulation'>Add-Drain c/w insulation</option><option value='Add-Drain w/o insulation'>Add-Drain w/o insulation</option><option value='Add-Gold'>Add-Gold</option><option value='Add-Silver'>Add-Silver</option><option value='Add-Standard'>Add-Standard</option><option value='All'>All</option><option value='AOS'>AOS</option><option value='AOS-Gas'>AOS-Gas</option><option value='B/Steel'>B/Steel</option><option value='Cement Plastering'>Cement Plastering</option><option value='CF'>CF</option><option value='CF-Gold'>CF-Gold</option><option value='Class "0"'>Class "0"</option><option value='Conceal Drain Pipe'>Conceal Drain Pipe</option><option value='Conceal Drain Pipe (FR)'>Conceal Drain Pipe (FR)</option><option value='Conceal Piping'>Conceal Piping</option><option value='Conceal Piping (FR)'>Conceal Piping (FR)</option><option value='Crome'>Crome</option><option value='Electrical Wiring Point'>Electrical Wiring Point</option><option value='GC'>GC</option><option value='GS'>GS</option><option value='GS-Gold'>GS-Gold</option><option value='ID'>ID</option><option value='MS'>MS</option><option value='OD'>OD</option><option value='Other'>Other</option><option value='Panel'>Panel</option><option value='Pkg-Basic'>Pkg-Basic</option><option value='Pkg-Basic (w/o pipe)'>Pkg-Basic (w/o pipe)</option><option value='Pkg-Gold'>Pkg-Gold</option><option value='Pkg-Silver'>Pkg-Silver</option><option value='Pkg-Silver (w/o pipe)'>Pkg-Silver (w/o pipe)</option><option value='Pkg-Standard'>Pkg-Standard</option><option value='Pkg-Standard (w/o pipe)'>Pkg-Standard (w/o pipe)</option><option value='Pressure Test'>Pressure Test</option><option value='PU'>PU</option><option value='R22'>R22</option><option value='R410a'>R410a</option><option value='Repair'>Repair</option><option value='Repair-Mean Leak'>Repair-Mean Leak</option><option value='Repair-Mean Leak+Pressure Test'>Repair-Mean Leak+Pressure Test</option><option value='Repair-Revacuum'>Repair-Revacuum</option><option value='Repair-Transmission Wire'>Repair-Transmission Wire</option><option value='Replacement'>Replacement</option><option value='Replacement-Component-Brazing'>Replacement-Component-Brazing</option><option value='Replacement-Compressor'>Replacement-Compressor</option><option value='Replacement-Condenser Coil'>Replacement-Condenser Coil</option><option value='Replacement-Evaporator'>Replacement-Evaporator</option><option value='S/Steel'>S/Steel</option><option value='Set Wired'>Set Wired</option><option value='Set Wireless'>Set Wireless</option><option value='Site'>Site</option><option value='Upgrade'>Upgrade</option><option value='Upgrade Bracket'>Upgrade Bracket</option><option value='Upgrade Copper'>Upgrade Copper</option><option value='Upgrade Insulation'>Upgrade Insulation</option><option value='Upgrade Wiring'>Upgrade Wiring</option><option value='VSC'>VSC</option><option value='white'>white</option><option value='Wired'>Wired</option><option value='Wireless'>Wireless</option>
      </select>
      </strong></td>
    <td><strong>	
      <select name="Capacity" onChange="document.forms[0].submit();" style="{width:80px;}">
        <option value="">- All -</option>
        <option value='0'>0</option><option value='0.75'>0.75</option><option value='1'>1</option><option value='1.0'>1.0</option><option value='1.0 hp'>1.0 hp</option><option value='1.0, 1.5'>1.0, 1.5</option><option value='1.0, 1.5, 2.0'>1.0, 1.5, 2.0</option><option value='1.0, 1.5, 2.0, 2.5'>1.0, 1.5, 2.0, 2.5</option><option value='1.5'>1.5</option><option value='1.5 hp'>1.5 hp</option><option value='1.75'>1.75</option><option value='10'>10</option><option value='10, 12, 15, 18, 20'>10, 12, 15, 18, 20</option><option value='11.3'>11.3</option><option value='11.300000000000001'>11.300000000000001</option><option value='12.0, 15'>12.0, 15</option><option value='12.5'>12.5</option><option value='13'>13</option><option value='13.6'>13.6</option><option value='13.6kg'>13.6kg</option><option value='15'>15</option><option value='17.5'>17.5</option><option value='18'>18</option><option value='18.0, 20.0'>18.0, 20.0</option><option value='2'>2</option><option value='2.0'>2.0</option><option value='2.0 hp'>2.0 hp</option><option value='2.0, 2.5, 3.0, 3.5'>2.0, 2.5, 3.0, 3.5</option><option value='2.5'>2.5</option><option value='2.5 hp'>2.5 hp</option><option value='2.5, 3.0, 3.5, 4.0, 4.5, 5.0'>2.5, 3.0, 3.5, 4.0, 4.5, 5.0</option><option value='2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0'>2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0</option><option value='2.6000000000000001'>2.6000000000000001</option><option value='2.75'>2.75</option><option value='20'>20</option><option value='20, 25, 30, 35, 40'>20, 25, 30, 35, 40</option><option value='25'>25</option><option value='3'>3</option><option value='3.0'>3.0</option><option value='3.0 hp'>3.0 hp</option><option value='3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0'>3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0</option><option value='3.5'>3.5</option><option value='3.75'>3.75</option><option value='30'>30</option><option value='32'>32</option><option value='35'>35</option><option value='4'>4</option><option value='4.0'>4.0</option><option value='4.0, 5.0'>4.0, 5.0</option><option value='4.5'>4.5</option><option value='40'>40</option><option value='5'>5</option><option value='5.0 hp'>5.0 hp</option><option value='5.5'>5.5</option><option value='6'>6</option><option value='6, 7, 8, 9'>6, 7, 8, 9</option><option value='6.0, 6.5'>6.0, 6.5</option><option value='6.5'>6.5</option><option value='7'>7</option><option value='7.0, 8.0'>7.0, 8.0</option><option value='7.5'>7.5</option><option value='8'>8</option><option value='8`'>8`</option>
      </select>
      </strong></td>
    <td align="left" valign="bottom" nowrap>
      <select name="pricetype" id="pricetype" style="width:80px">
        <option value="GrossPurchasesInvPrice" >GrossPurchasesInvPrice</option>
        <option value="StdSellingPrice"  selected>Standard</option>
        <option value="Min1" >Min 1</option>
        <option value="Volumn3" >Min 3</option>
        <option value="Volumn5" >Min 5</option>
        <option value="Volumn10" >Min 10</option>
        <option value="Project" >Project</option>
        <option value="WholeSales" >WholeSale</option>
        <option value="WholeSalesMin" >WholeS.Min</option>
      </select></td>
  </tr>
  </form>
  
  <tr> 
    <td colspan="11"><strong> </strong><strong> </strong><strong> </strong><strong> 
      </strong><strong> </strong><strong> </strong> <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#E6E6E6">
        <tr> 
          <td colspan="12" align="right"><strong>Page</strong> <font color="3366ff"> 1 </font>of <font color="3366ff"> 80 </font>: 
            <font color=#000000><b>1</b></font>  <a href='searchboxStockInItems.asp?num=50&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>2</a>  <a href='searchboxStockInItems.asp?num=100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>3</a>  <a href='searchboxStockInItems.asp?num=150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>4</a>  <a href='searchboxStockInItems.asp?num=200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>5</a>  <a href='searchboxStockInItems.asp?num=250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>6</a>  <a href='searchboxStockInItems.asp?num=300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>7</a>  <a href='searchboxStockInItems.asp?num=350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>8</a>  <a href='searchboxStockInItems.asp?num=400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>9</a>  <a href='searchboxStockInItems.asp?num=450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>10</a>  <a href='searchboxStockInItems.asp?num=500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>11</a>  <a href='searchboxStockInItems.asp?num=550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>12</a>  <a href='searchboxStockInItems.asp?num=600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>13</a>  <a href='searchboxStockInItems.asp?num=650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>14</a>  <a href='searchboxStockInItems.asp?num=700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>15</a>  <a href='searchboxStockInItems.asp?num=750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>16</a>  <a href='searchboxStockInItems.asp?num=800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>17</a>  <a href='searchboxStockInItems.asp?num=850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>18</a>  <a href='searchboxStockInItems.asp?num=900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>19</a>  <a href='searchboxStockInItems.asp?num=950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>20</a>  <a href='searchboxStockInItems.asp?num=1000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>21</a>  <a href='searchboxStockInItems.asp?num=1050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>22</a>  <a href='searchboxStockInItems.asp?num=1100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>23</a>  <a href='searchboxStockInItems.asp?num=1150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>24</a>  <a href='searchboxStockInItems.asp?num=1200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>25</a>  <a href='searchboxStockInItems.asp?num=1250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>26</a>  <a href='searchboxStockInItems.asp?num=1300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>27</a>  <a href='searchboxStockInItems.asp?num=1350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>28</a>  <a href='searchboxStockInItems.asp?num=1400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>29</a>  <a href='searchboxStockInItems.asp?num=1450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>30</a>  <a href='searchboxStockInItems.asp?num=1500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>31</a>  <a href='searchboxStockInItems.asp?num=1550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>32</a>  <a href='searchboxStockInItems.asp?num=1600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>33</a>  <a href='searchboxStockInItems.asp?num=1650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>34</a>  <a href='searchboxStockInItems.asp?num=1700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>35</a>  <a href='searchboxStockInItems.asp?num=1750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>36</a>  <a href='searchboxStockInItems.asp?num=1800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>37</a>  <a href='searchboxStockInItems.asp?num=1850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>38</a>  <a href='searchboxStockInItems.asp?num=1900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>39</a>  <a href='searchboxStockInItems.asp?num=1950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>40</a>  <a href='searchboxStockInItems.asp?num=2000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>41</a>  <a href='searchboxStockInItems.asp?num=2050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>42</a>  <a href='searchboxStockInItems.asp?num=2100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>43</a>  <a href='searchboxStockInItems.asp?num=2150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>44</a>  <a href='searchboxStockInItems.asp?num=2200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>45</a>  <a href='searchboxStockInItems.asp?num=2250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>46</a>  <a href='searchboxStockInItems.asp?num=2300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>47</a>  <a href='searchboxStockInItems.asp?num=2350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>48</a>  <a href='searchboxStockInItems.asp?num=2400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>49</a>  <a href='searchboxStockInItems.asp?num=2450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>50</a>  <a href='searchboxStockInItems.asp?num=2500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>51</a>  <a href='searchboxStockInItems.asp?num=2550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>52</a>  <a href='searchboxStockInItems.asp?num=2600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>53</a>  <a href='searchboxStockInItems.asp?num=2650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>54</a>  <a href='searchboxStockInItems.asp?num=2700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>55</a>  <a href='searchboxStockInItems.asp?num=2750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>56</a>  <a href='searchboxStockInItems.asp?num=2800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>57</a>  <a href='searchboxStockInItems.asp?num=2850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>58</a>  <a href='searchboxStockInItems.asp?num=2900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>59</a>  <a href='searchboxStockInItems.asp?num=2950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>60</a>  <a href='searchboxStockInItems.asp?num=3000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>61</a>  <a href='searchboxStockInItems.asp?num=3050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>62</a>  <a href='searchboxStockInItems.asp?num=3100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>63</a>  <a href='searchboxStockInItems.asp?num=3150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>64</a>  <a href='searchboxStockInItems.asp?num=3200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>65</a>  <a href='searchboxStockInItems.asp?num=3250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>66</a>  <a href='searchboxStockInItems.asp?num=3300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>67</a>  <a href='searchboxStockInItems.asp?num=3350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>68</a>  <a href='searchboxStockInItems.asp?num=3400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>69</a>  <a href='searchboxStockInItems.asp?num=3450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>70</a>  <a href='searchboxStockInItems.asp?num=3500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>71</a>  <a href='searchboxStockInItems.asp?num=3550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>72</a>  <a href='searchboxStockInItems.asp?num=3600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>73</a>  <a href='searchboxStockInItems.asp?num=3650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>74</a>  <a href='searchboxStockInItems.asp?num=3700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>75</a>  <a href='searchboxStockInItems.asp?num=3750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>76</a>  <a href='searchboxStockInItems.asp?num=3800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>77</a>  <a href='searchboxStockInItems.asp?num=3850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>78</a>  <a href='searchboxStockInItems.asp?num=3900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>79</a>  <a href='searchboxStockInItems.asp?num=3950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>80</a><a href='searchboxStockInItems.asp?num=50&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='> Next >></a> </td>
        </tr>
        <tr> 
          <th width="29" bgcolor="#CCCCCC"><strong><font color="#000000">No.</font></strong> 
          </th>
          <th align="left" nowrap bgcolor="#CCCCCC"><strong> <font color="#000000">Group</font></strong> 
          </th>
          <td width="84" bgcolor="#CCCCCC"><strong><font color="#000000">Code</font></strong></td>
          <td width="95" bgcolor="#CCCCCC"><strong><font color="#000000">Description</font></strong></td>
          <td width="95" bgcolor="#CCCCCC"> <strong><font color="#000000">Brand</font></strong></td>
          <td width="95" bgcolor="#CCCCCC"> <strong><font color="#000000">Type</font></strong></td>
          <td width="95" bgcolor="#CCCCCC"> <strong><font color="#000000">Model</font></strong></td>
          <td width="97" bgcolor="#CCCCCC"><strong><font color="#000000">Sub-Type1 
            / <br>
            Sub-Type2</font></strong></td>
          <td bgcolor="#CCCCCC"><font color="#000000"><strong> Capacity</strong></font></td>
          <td align="right" nowrap bgcolor="#CCCCCC"><font color="#000000"><strong>IN Stock</strong></font></td>
          <td align="right" nowrap bgcolor="#CCCCCC"><strong>On Order</strong></td>
          <td align="right" nowrap bgcolor="#CCCCCC"><strong>Free Stock</strong></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">1.</font></td>
          <td nowrap><font size="1"> AC</font></td>
          <td><font size="1">
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='10159';parent.document.forms['form4'].StockCode.value='';parent.document.forms['form4'].DESCRIPTION.value='';parent.document.forms['form4'].UNIT_PRICE.value='0';parent.document.forms['form4'].NET_AMOUNT.value='0';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='0';">
           
            </font></td>
          <td><font size="1"> </font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS410i</font></td>
          <td><font size="1"></font></td>
          <td><font size="1">WM / OD</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">-2</font></td>
          <td align="right" nowrap><font size="1"></font></td>
          <td align="right" nowrap><font size="1"></font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">2.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0012445
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6165';parent.document.forms['form4'].StockCode.value='0012445';parent.document.forms['form4'].DESCRIPTION.value='+Pan machine screw (brass)';parent.document.forms['form4'].UNIT_PRICE.value='2';parent.document.forms['form4'].NET_AMOUNT.value='2';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='2';">
           
            </font></td>
          <td><font size="1"> +Pan machine screw (brass)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FD10KY1</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">3.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0014395
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4920';parent.document.forms['form4'].StockCode.value='0014395';parent.document.forms['form4'].DESCRIPTION.value='SEALER, DRAIN HOSE';parent.document.forms['form4'].UNIT_PRICE.value='4';parent.document.forms['form4'].NET_AMOUNT.value='4';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='4';">
           
            </font></td>
          <td><font size="1"> SEALER, DRAIN HOSE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FHC48NUV1</font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">4.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0024758
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1173';parent.document.forms['form4'].StockCode.value='0024758';parent.document.forms['form4'].DESCRIPTION.value='THERMISTOR (FOR AIR)';parent.document.forms['form4'].UNIT_PRICE.value='14';parent.document.forms['form4'].NET_AMOUNT.value='14';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='14';">
           
            </font></td>
          <td><font size="1"> THERMISTOR (FOR AIR)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1"></font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">5.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0028882
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='3978';parent.document.forms['form4'].StockCode.value='0028882';parent.document.forms['form4'].DESCRIPTION.value='TUBE CLAMP';parent.document.forms['form4'].UNIT_PRICE.value='5';parent.document.forms['form4'].NET_AMOUNT.value='5';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='5';">
           
            </font></td>
          <td><font size="1"> TUBE CLAMP</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FDBG35AVE</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">6.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0030380
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='10031';parent.document.forms['form4'].StockCode.value='0030380';parent.document.forms['form4'].DESCRIPTION.value='Liquid Thermistor';parent.document.forms['form4'].UNIT_PRICE.value='0';parent.document.forms['form4'].NET_AMOUNT.value='0';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='0';">
           
            </font></td>
          <td><font size="1"> Liquid Thermistor</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> VRV</font></td>
          <td><font size="1"></font></td>
          <td><font size="1">0 / Wireless</font></td>
          <td nowrap><font size="1"></font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">7.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0030652
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4922';parent.document.forms['form4'].StockCode.value='0030652';parent.document.forms['form4'].DESCRIPTION.value='FUSE';parent.document.forms['form4'].UNIT_PRICE.value='4';parent.document.forms['form4'].NET_AMOUNT.value='4';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='4';">
           
            </font></td>
          <td><font size="1"> FUSE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS410</font></td>
          <td><font size="1">RUR10NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">8.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0031561
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2514';parent.document.forms['form4'].StockCode.value='0031561';parent.document.forms['form4'].DESCRIPTION.value='VIBRATION ISOLATOR';parent.document.forms['form4'].UNIT_PRICE.value='12';parent.document.forms['form4'].NET_AMOUNT.value='12';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='12';">
           
            </font></td>
          <td><font size="1"> VIBRATION ISOLATOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU08KUY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">8.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">9.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">003156J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2574';parent.document.forms['form4'].StockCode.value='003156J';parent.document.forms['form4'].DESCRIPTION.value='VIBRATION ISOLATOR';parent.document.forms['form4'].UNIT_PRICE.value='12';parent.document.forms['form4'].NET_AMOUNT.value='12';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='12';">
           
            </font></td>
          <td><font size="1"> VIBRATION ISOLATOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R56NUY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">6.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">10.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0032247
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6839';parent.document.forms['form4'].StockCode.value='0032247';parent.document.forms['form4'].DESCRIPTION.value='THERMISTOR';parent.document.forms['form4'].UNIT_PRICE.value='28';parent.document.forms['form4'].NET_AMOUNT.value='28';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='28';">
           
            </font></td>
          <td><font size="1"> THERMISTOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FH50BVE</font></td>
          <td><font size="1">CE / ID</font></td>
          <td nowrap><font size="1">2.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">11.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0033288
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='5437';parent.document.forms['form4'].StockCode.value='0033288';parent.document.forms['form4'].DESCRIPTION.value='NOISE ABSORBER, COMPRESSOR';parent.document.forms['form4'].UNIT_PRICE.value='66';parent.document.forms['form4'].NET_AMOUNT.value='66';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='66';">
           
            </font></td>
          <td><font size="1"> NOISE ABSORBER, COMPRESSOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R50BV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">12.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0033334
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4345';parent.document.forms['form4'].StockCode.value='0033334';parent.document.forms['form4'].DESCRIPTION.value='AIR DISCHARGE GRILLE ASS`Y';parent.document.forms['form4'].UNIT_PRICE.value='131';parent.document.forms['form4'].NET_AMOUNT.value='131';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='131';">
           
            </font></td>
          <td><font size="1"> AIR DISCHARGE GRILLE ASS'Y</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R60BV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">13.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0038827
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='3905';parent.document.forms['form4'].StockCode.value='0038827';parent.document.forms['form4'].DESCRIPTION.value='BLIND CAP. SERVICING PORT';parent.document.forms['form4'].UNIT_PRICE.value='6';parent.document.forms['form4'].NET_AMOUNT.value='6';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='6';">
           
            </font></td>
          <td><font size="1"> BLIND CAP. SERVICING PORT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU10KUY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">14.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0038973
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4597';parent.document.forms['form4'].StockCode.value='0038973';parent.document.forms['form4'].DESCRIPTION.value='LOW PRESSURE SWITCH';parent.document.forms['form4'].UNIT_PRICE.value='92';parent.document.forms['form4'].NET_AMOUNT.value='92';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='92';">
           
            </font></td>
          <td><font size="1"> LOW PRESSURE SWITCH</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU20NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">20.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">15.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0072238
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6372';parent.document.forms['form4'].StockCode.value='0072238';parent.document.forms['form4'].DESCRIPTION.value='DRAIN HOSE ASS`Y';parent.document.forms['form4'].UNIT_PRICE.value='43';parent.document.forms['form4'].NET_AMOUNT.value='43';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='43';">
           
            </font></td>
          <td><font size="1"> DRAIN HOSE ASS'Y</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FHC35KVE</font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">-1</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">16.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">007223J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4892';parent.document.forms['form4'].StockCode.value='007223J';parent.document.forms['form4'].DESCRIPTION.value='DRAIN HOSE ASS`Y';parent.document.forms['form4'].UNIT_PRICE.value='5';parent.document.forms['form4'].NET_AMOUNT.value='5';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='5';">
           
            </font></td>
          <td><font size="1"> DRAIN HOSE ASS'Y</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FHC48NUV1</font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">17.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0074328
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='5196';parent.document.forms['form4'].StockCode.value='0074328';parent.document.forms['form4'].DESCRIPTION.value='RETAINER. THERMISTOR';parent.document.forms['form4'].UNIT_PRICE.value='3';parent.document.forms['form4'].NET_AMOUNT.value='3';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='3';">
           
            </font></td>
          <td><font size="1"> RETAINER. THERMISTOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FH50BVE</font></td>
          <td><font size="1">CE / ID</font></td>
          <td nowrap><font size="1">2.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">18.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">007432J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6128';parent.document.forms['form4'].StockCode.value='007432J';parent.document.forms['form4'].DESCRIPTION.value='RETAINER. THERMISTOR';parent.document.forms['form4'].UNIT_PRICE.value='3';parent.document.forms['form4'].NET_AMOUNT.value='3';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='3';">
           
            </font></td>
          <td><font size="1"> RETAINER. THERMISTOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FH48NUV1</font></td>
          <td><font size="1">CE / ID</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">19.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0079866
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4560';parent.document.forms['form4'].StockCode.value='0079866';parent.document.forms['form4'].DESCRIPTION.value='PARTITION PLATE  R60AV1';parent.document.forms['form4'].UNIT_PRICE.value='95';parent.document.forms['form4'].NET_AMOUNT.value='95';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='95';">
           
            </font></td>
          <td><font size="1"> PARTITION PLATE  R60AV1</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R21NUV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">20.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0079929
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='7295';parent.document.forms['form4'].StockCode.value='0079929';parent.document.forms['form4'].DESCRIPTION.value='STOP VALVE ASS`Y (GAS LINE)';parent.document.forms['form4'].UNIT_PRICE.value='22';parent.document.forms['form4'].NET_AMOUNT.value='22';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='22';">
           
            </font></td>
          <td><font size="1"> STOP VALVE ASS'Y (GAS LINE)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R21NUV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">21.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0079936
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1881';parent.document.forms['form4'].StockCode.value='0079936';parent.document.forms['form4'].DESCRIPTION.value='CAP. STOP VALVE';parent.document.forms['form4'].UNIT_PRICE.value='12';parent.document.forms['form4'].NET_AMOUNT.value='12';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='12';">
           
            </font></td>
          <td><font size="1"> CAP. STOP VALVE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKE35BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">22.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">007993J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2579';parent.document.forms['form4'].StockCode.value='007993J';parent.document.forms['form4'].DESCRIPTION.value='CAP. STOP VALVE';parent.document.forms['form4'].UNIT_PRICE.value='12';parent.document.forms['form4'].NET_AMOUNT.value='12';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='12';">
           
            </font></td>
          <td><font size="1"> CAP. STOP VALVE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKD71BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">3.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">23.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0079967
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='7229';parent.document.forms['form4'].StockCode.value='0079967';parent.document.forms['form4'].DESCRIPTION.value='TERMINAL BLOCK';parent.document.forms['form4'].UNIT_PRICE.value='24';parent.document.forms['form4'].NET_AMOUNT.value='24';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='24';">
           
            </font></td>
          <td><font size="1"> TERMINAL BLOCK</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R60BV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">24.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0080347
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1175';parent.document.forms['form4'].StockCode.value='0080347';parent.document.forms['form4'].DESCRIPTION.value='FLARE NUT';parent.document.forms['form4'].UNIT_PRICE.value='14';parent.document.forms['form4'].NET_AMOUNT.value='14';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='14';">
           
            </font></td>
          <td><font size="1"> FLARE NUT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKE25BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">25.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0080354
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='7297';parent.document.forms['form4'].StockCode.value='0080354';parent.document.forms['form4'].DESCRIPTION.value='OVER-LOAD RELAY';parent.document.forms['form4'].UNIT_PRICE.value='22';parent.document.forms['form4'].NET_AMOUNT.value='22';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='22';">
           
            </font></td>
          <td><font size="1"> OVER-LOAD RELAY</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R25DV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.0</font></td>
          <td align="right" nowrap><font size="1">1</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">1</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">26.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0082033
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6180';parent.document.forms['form4'].StockCode.value='0082033';parent.document.forms['form4'].DESCRIPTION.value='TRUSS HEAD TAPPING SCREW';parent.document.forms['form4'].UNIT_PRICE.value='2';parent.document.forms['form4'].NET_AMOUNT.value='2';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='2';">
           
            </font></td>
          <td><font size="1"> TRUSS HEAD TAPPING SCREW</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKD35DVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">27.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">008792J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1244';parent.document.forms['form4'].StockCode.value='008792J';parent.document.forms['form4'].DESCRIPTION.value='THERMISTOR (FOR AIR)';parent.document.forms['form4'].UNIT_PRICE.value='14';parent.document.forms['form4'].NET_AMOUNT.value='14';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='14';">
           
            </font></td>
          <td><font size="1"> THERMISTOR (FOR AIR)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FHC48NUV1</font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">28.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0094151
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4619';parent.document.forms['form4'].StockCode.value='0094151';parent.document.forms['form4'].DESCRIPTION.value='STOP VALVE ASS`Y (GAS LINE)';parent.document.forms['form4'].UNIT_PRICE.value='88';parent.document.forms['form4'].NET_AMOUNT.value='88';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='88';">
           
            </font></td>
          <td><font size="1"> STOP VALVE ASS'Y (GAS LINE)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R18NUV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">29.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0100676
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2591';parent.document.forms['form4'].StockCode.value='0100676';parent.document.forms['form4'].DESCRIPTION.value='HALF UNION JOINT';parent.document.forms['form4'].UNIT_PRICE.value='11';parent.document.forms['form4'].NET_AMOUNT.value='11';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='11';">
           
            </font></td>
          <td><font size="1"> HALF UNION JOINT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FDBG35AVE</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">30.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0106030
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6666';parent.document.forms['form4'].StockCode.value='0106030';parent.document.forms['form4'].DESCRIPTION.value='THERMISTOR (FOR OUTDOOR AIR)';parent.document.forms['form4'].UNIT_PRICE.value='32';parent.document.forms['form4'].NET_AMOUNT.value='32';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='32';">
           
            </font></td>
          <td><font size="1"> THERMISTOR (FOR OUTDOOR AIR)</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKD71BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">3.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">31.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0111854
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2598';parent.document.forms['form4'].StockCode.value='0111854';parent.document.forms['form4'].DESCRIPTION.value='PLUG';parent.document.forms['form4'].UNIT_PRICE.value='11';parent.document.forms['form4'].NET_AMOUNT.value='11';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='11';">
           
            </font></td>
          <td><font size="1"> PLUG</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FDR10NY1</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">32.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">011185J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2662';parent.document.forms['form4'].StockCode.value='011185J';parent.document.forms['form4'].DESCRIPTION.value='PLUG';parent.document.forms['form4'].UNIT_PRICE.value='11';parent.document.forms['form4'].NET_AMOUNT.value='11';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='11';">
           
            </font></td>
          <td><font size="1"> PLUG</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FDR08NY1</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">8.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">33.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0115825
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6200';parent.document.forms['form4'].StockCode.value='0115825';parent.document.forms['form4'].DESCRIPTION.value='HEX. SOCKET SCREW';parent.document.forms['form4'].UNIT_PRICE.value='2';parent.document.forms['form4'].NET_AMOUNT.value='2';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='2';">
           
            </font></td>
          <td><font size="1"> HEX. SOCKET SCREW</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FH50BVE</font></td>
          <td><font size="1">CE / ID</font></td>
          <td nowrap><font size="1">2.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">34.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0119137
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='3985';parent.document.forms['form4'].StockCode.value='0119137';parent.document.forms['form4'].DESCRIPTION.value='RETAINING SPRING/ THERMOSTAT';parent.document.forms['form4'].UNIT_PRICE.value='5';parent.document.forms['form4'].NET_AMOUNT.value='5';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='5';">
           
            </font></td>
          <td><font size="1"> RETAINING SPRING/ THERMOSTAT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS410</font></td>
          <td><font size="1">RUR05NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">35.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0120005
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4397';parent.document.forms['form4'].StockCode.value='0120005';parent.document.forms['form4'].DESCRIPTION.value='REFRIGERANT FILTER';parent.document.forms['form4'].UNIT_PRICE.value='120';parent.document.forms['form4'].NET_AMOUNT.value='120';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='120';">
           
            </font></td>
          <td><font size="1"> REFRIGERANT FILTER</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FD10KAY1</font></td>
          <td><font size="1">Ducted / ID</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">36.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0132279
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='6201';parent.document.forms['form4'].StockCode.value='0132279';parent.document.forms['form4'].DESCRIPTION.value='PACKING/ CONNECTING FLANGE';parent.document.forms['form4'].UNIT_PRICE.value='2';parent.document.forms['form4'].NET_AMOUNT.value='2';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='2';">
           
            </font></td>
          <td><font size="1"> PACKING/ CONNECTING FLANGE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU20NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">20.0</font></td>
          <td align="right" nowrap><font size="1">22</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">22</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">37.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0133001
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='5679';parent.document.forms['form4'].StockCode.value='0133001';parent.document.forms['form4'].DESCRIPTION.value='TERMINAL PLATE';parent.document.forms['form4'].UNIT_PRICE.value='48';parent.document.forms['form4'].NET_AMOUNT.value='48';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='48';">
           
            </font></td>
          <td><font size="1"> TERMINAL PLATE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU10KUY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">38.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0137856
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='927';parent.document.forms['form4'].StockCode.value='0137856';parent.document.forms['form4'].DESCRIPTION.value='FLARE NUT';parent.document.forms['form4'].UNIT_PRICE.value='16';parent.document.forms['form4'].NET_AMOUNT.value='16';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='16';">
           
            </font></td>
          <td><font size="1"> FLARE NUT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKE35BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">39.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">013785J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1090';parent.document.forms['form4'].StockCode.value='013785J';parent.document.forms['form4'].DESCRIPTION.value='FLARE NUT';parent.document.forms['form4'].UNIT_PRICE.value='16';parent.document.forms['form4'].NET_AMOUNT.value='16';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='16';">
           
            </font></td>
          <td><font size="1"> FLARE NUT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKD35DVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">1.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">40.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0141552
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='9892';parent.document.forms['form4'].StockCode.value='0141552';parent.document.forms['form4'].DESCRIPTION.value='Magnetic Switch';parent.document.forms['form4'].UNIT_PRICE.value='0.5';parent.document.forms['form4'].NET_AMOUNT.value='0.5';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='0.5';">
           
            </font></td>
          <td><font size="1"> Magnetic Switch</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> 0</font></td>
          <td><font size="1"></font></td>
          <td><font size="1">0 / 0</font></td>
          <td nowrap><font size="1">0.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">41.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0141576
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='3724';parent.document.forms['form4'].StockCode.value='0141576';parent.document.forms['form4'].DESCRIPTION.value='SET BAND, CAPACITOR';parent.document.forms['form4'].UNIT_PRICE.value='7';parent.document.forms['form4'].NET_AMOUNT.value='7';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='7';">
           
            </font></td>
          <td><font size="1"> SET BAND, CAPACITOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">R21NUV1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">42.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0144359
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4926';parent.document.forms['form4'].StockCode.value='0144359';parent.document.forms['form4'].DESCRIPTION.value='FUSE';parent.document.forms['form4'].UNIT_PRICE.value='4';parent.document.forms['form4'].NET_AMOUNT.value='4';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='4';">
           
            </font></td>
          <td><font size="1"> FUSE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU10KUY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">10.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">43.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0154457
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='9894';parent.document.forms['form4'].StockCode.value='0154457';parent.document.forms['form4'].DESCRIPTION.value='Printed Circuit Board';parent.document.forms['form4'].UNIT_PRICE.value='0.5';parent.document.forms['form4'].NET_AMOUNT.value='0.5';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='0.5';">
           
            </font></td>
          <td><font size="1"> Printed Circuit Board</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> </font></td>
          <td><font size="1"></font></td>
          <td><font size="1"> / </font></td>
          <td nowrap><font size="1"></font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">44.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0154844
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='3551';parent.document.forms['form4'].StockCode.value='0154844';parent.document.forms['form4'].DESCRIPTION.value='HIGH PRESSURE CONTROL VALVE';parent.document.forms['form4'].UNIT_PRICE.value='187';parent.document.forms['form4'].NET_AMOUNT.value='187';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='187';">
           
            </font></td>
          <td><font size="1"> HIGH PRESSURE CONTROL VALVE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">RU20NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">20.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">45.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0157711
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2874';parent.document.forms['form4'].StockCode.value='0157711';parent.document.forms['form4'].DESCRIPTION.value='TUBE CLAMP';parent.document.forms['form4'].UNIT_PRICE.value='8';parent.document.forms['form4'].NET_AMOUNT.value='8';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='8';">
           
            </font></td>
          <td><font size="1"> TUBE CLAMP</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS410</font></td>
          <td><font size="1">RUR13NY1</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">13.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">46.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0171173
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='4927';parent.document.forms['form4'].StockCode.value='0171173';parent.document.forms['form4'].DESCRIPTION.value='VIBRATION ISOLATOR';parent.document.forms['form4'].UNIT_PRICE.value='4';parent.document.forms['form4'].NET_AMOUNT.value='4';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='4';">
           
            </font></td>
          <td><font size="1"> VIBRATION ISOLATOR</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">RKD71BVM</font></td>
          <td><font size="1">OD / OD</font></td>
          <td nowrap><font size="1">3.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">47.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0198433
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='1122';parent.document.forms['form4'].StockCode.value='0198433';parent.document.forms['form4'].DESCRIPTION.value='UNION JOINT, LIQUID LINE';parent.document.forms['form4'].UNIT_PRICE.value='15';parent.document.forms['form4'].NET_AMOUNT.value='15';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='15';">
           
            </font></td>
          <td><font size="1"> UNION JOINT, LIQUID LINE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FHC48NUV1</font></td>
          <td><font size="1">CK / ID</font></td>
          <td nowrap><font size="1">5.0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">48.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0202262
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2706';parent.document.forms['form4'].StockCode.value='0202262';parent.document.forms['form4'].DESCRIPTION.value='UNION JOINT, LIQUID LINE';parent.document.forms['form4'].UNIT_PRICE.value='9';parent.document.forms['form4'].NET_AMOUNT.value='9';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='9';">
           
            </font></td>
          <td><font size="1"> UNION JOINT, LIQUID LINE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FT60BVM</font></td>
          <td><font size="1">WM / ID</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#F3F3F3"> 
          <td align="center"><font size="1">49.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">020226J
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='2848';parent.document.forms['form4'].StockCode.value='020226J';parent.document.forms['form4'].DESCRIPTION.value='UNION JOINT, LIQUID LINE';parent.document.forms['form4'].UNIT_PRICE.value='9';parent.document.forms['form4'].NET_AMOUNT.value='9';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='9';">
           
            </font></td>
          <td><font size="1"> UNION JOINT, LIQUID LINE</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22i</font></td>
          <td><font size="1">FTKD60FVM</font></td>
          <td><font size="1">WM / ID</font></td>
          <td nowrap><font size="1">2.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr bgcolor="#FFFFFF"> 
          <td align="center"><font size="1">50.</font></td>
          <td nowrap><font size="1"> Parts</font></td>
          <td><font size="1">0202370
            <input type="button" name="Button" value="get" onClick="javascript:parent.document.forms['form4'].stock_id.value='7230';parent.document.forms['form4'].StockCode.value='0202370';parent.document.forms['form4'].DESCRIPTION.value='UNION JOINT';parent.document.forms['form4'].UNIT_PRICE.value='24';parent.document.forms['form4'].NET_AMOUNT.value='24';parent.document.forms['form4'].QTY_ORDER.value='1';parent.document.forms['form4'].FULL_NET_AMOUNT.value='24';">
           
            </font></td>
          <td><font size="1"> UNION JOINT</font></td>
          <td><font size="1"> Daikin</font></td>
          <td><font size="1"> SS22</font></td>
          <td><font size="1">FH30NUV1</font></td>
          <td><font size="1">CE / ID</font></td>
          <td nowrap><font size="1">3.5</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
          <td align="right" nowrap><font size="1">0</font></td>
        </tr>
        
        <tr> 
          <td colspan="12" align="right"><strong>Page</strong> <font color="3366ff"> 1 </font>of <font color="3366ff"> 80 </font>: 
          <font color=#000000><b>1</b></font>  <a href='searchboxStockInItems.asp?num=50&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>2</a>  <a href='searchboxStockInItems.asp?num=100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>3</a>  <a href='searchboxStockInItems.asp?num=150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>4</a>  <a href='searchboxStockInItems.asp?num=200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>5</a>  <a href='searchboxStockInItems.asp?num=250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>6</a>  <a href='searchboxStockInItems.asp?num=300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>7</a>  <a href='searchboxStockInItems.asp?num=350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>8</a>  <a href='searchboxStockInItems.asp?num=400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>9</a>  <a href='searchboxStockInItems.asp?num=450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>10</a>  <a href='searchboxStockInItems.asp?num=500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>11</a>  <a href='searchboxStockInItems.asp?num=550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>12</a>  <a href='searchboxStockInItems.asp?num=600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>13</a>  <a href='searchboxStockInItems.asp?num=650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>14</a>  <a href='searchboxStockInItems.asp?num=700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>15</a>  <a href='searchboxStockInItems.asp?num=750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>16</a>  <a href='searchboxStockInItems.asp?num=800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>17</a>  <a href='searchboxStockInItems.asp?num=850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>18</a>  <a href='searchboxStockInItems.asp?num=900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>19</a>  <a href='searchboxStockInItems.asp?num=950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>20</a>  <a href='searchboxStockInItems.asp?num=1000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>21</a>  <a href='searchboxStockInItems.asp?num=1050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>22</a>  <a href='searchboxStockInItems.asp?num=1100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>23</a>  <a href='searchboxStockInItems.asp?num=1150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>24</a>  <a href='searchboxStockInItems.asp?num=1200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>25</a>  <a href='searchboxStockInItems.asp?num=1250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>26</a>  <a href='searchboxStockInItems.asp?num=1300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>27</a>  <a href='searchboxStockInItems.asp?num=1350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>28</a>  <a href='searchboxStockInItems.asp?num=1400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>29</a>  <a href='searchboxStockInItems.asp?num=1450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>30</a>  <a href='searchboxStockInItems.asp?num=1500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>31</a>  <a href='searchboxStockInItems.asp?num=1550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>32</a>  <a href='searchboxStockInItems.asp?num=1600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>33</a>  <a href='searchboxStockInItems.asp?num=1650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>34</a>  <a href='searchboxStockInItems.asp?num=1700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>35</a>  <a href='searchboxStockInItems.asp?num=1750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>36</a>  <a href='searchboxStockInItems.asp?num=1800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>37</a>  <a href='searchboxStockInItems.asp?num=1850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>38</a>  <a href='searchboxStockInItems.asp?num=1900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>39</a>  <a href='searchboxStockInItems.asp?num=1950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>40</a>  <a href='searchboxStockInItems.asp?num=2000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>41</a>  <a href='searchboxStockInItems.asp?num=2050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>42</a>  <a href='searchboxStockInItems.asp?num=2100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>43</a>  <a href='searchboxStockInItems.asp?num=2150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>44</a>  <a href='searchboxStockInItems.asp?num=2200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>45</a>  <a href='searchboxStockInItems.asp?num=2250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>46</a>  <a href='searchboxStockInItems.asp?num=2300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>47</a>  <a href='searchboxStockInItems.asp?num=2350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>48</a>  <a href='searchboxStockInItems.asp?num=2400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>49</a>  <a href='searchboxStockInItems.asp?num=2450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>50</a>  <a href='searchboxStockInItems.asp?num=2500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>51</a>  <a href='searchboxStockInItems.asp?num=2550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>52</a>  <a href='searchboxStockInItems.asp?num=2600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>53</a>  <a href='searchboxStockInItems.asp?num=2650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>54</a>  <a href='searchboxStockInItems.asp?num=2700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>55</a>  <a href='searchboxStockInItems.asp?num=2750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>56</a>  <a href='searchboxStockInItems.asp?num=2800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>57</a>  <a href='searchboxStockInItems.asp?num=2850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>58</a>  <a href='searchboxStockInItems.asp?num=2900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>59</a>  <a href='searchboxStockInItems.asp?num=2950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>60</a>  <a href='searchboxStockInItems.asp?num=3000&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>61</a>  <a href='searchboxStockInItems.asp?num=3050&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>62</a>  <a href='searchboxStockInItems.asp?num=3100&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>63</a>  <a href='searchboxStockInItems.asp?num=3150&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>64</a>  <a href='searchboxStockInItems.asp?num=3200&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>65</a>  <a href='searchboxStockInItems.asp?num=3250&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>66</a>  <a href='searchboxStockInItems.asp?num=3300&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>67</a>  <a href='searchboxStockInItems.asp?num=3350&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>68</a>  <a href='searchboxStockInItems.asp?num=3400&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>69</a>  <a href='searchboxStockInItems.asp?num=3450&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>70</a>  <a href='searchboxStockInItems.asp?num=3500&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>71</a>  <a href='searchboxStockInItems.asp?num=3550&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>72</a>  <a href='searchboxStockInItems.asp?num=3600&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>73</a>  <a href='searchboxStockInItems.asp?num=3650&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>74</a>  <a href='searchboxStockInItems.asp?num=3700&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>75</a>  <a href='searchboxStockInItems.asp?num=3750&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>76</a>  <a href='searchboxStockInItems.asp?num=3800&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>77</a>  <a href='searchboxStockInItems.asp?num=3850&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>78</a>  <a href='searchboxStockInItems.asp?num=3900&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>79</a>  <a href='searchboxStockInItems.asp?num=3950&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='>80</a><a href='searchboxStockInItems.asp?num=50&ItemGroup=&Code=&Description=&Brand=&ItemType=&ModelParts=&SubType1=&SubType2=&Capacity='> Next >></a> </td>
        </tr>
      </table>
 </td>
  </tr>
</table>
</body>
</html>
