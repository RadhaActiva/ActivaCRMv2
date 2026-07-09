<script language="JavaScript1.2" type="text/javascript">
<!--
sniffBrowsers();

menuItemBullet = new bulletPoint("images/menu_off.gif","images/menu_on.gif");
labelBullet = new bulletPoint("images/header_off.gif","images/header_on.gif");
subMenuBullet = new bulletPoint("images/sub_header_off.gif","images/sub_header_on.gif");

myTest = new menuBar('myTest',700, 'horizontal', '#000000', '#000000');
    myTest.addLabel('labelBullet', 'Dashboard', 1, 120, '#515A5A ', '#3d85c6', 'rm_home.asp', 'left');
    myTest.addLabel('labelBullet', 'Job Sheet', 2, 120, '#515A5A ', '#3d85c6', 'sales_jobsheet_view.asp? type = reset', 'left');
    myTest.addLabel('labelBullet', 'Stock Master', 3, 120, '#515A5A ', '#3d85c6', 'sales_stock_view.asp', 'left');
myTest.height = 38;

menus[1] = new menu(200, 'vertical', '#000000', '#000000');
menus[1].height = 38;
menus[1].writeMenu();


menus[1].align='left';
menus[2].align='left';
menus[3].align='left';
menus[4].align = 'left';

//-->
</script>