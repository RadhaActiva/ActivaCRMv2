<script language="JavaScript1.2" type="text/javascript">
<!--
sniffBrowsers();

menuItemBullet = new bulletPoint("images/menu_off.gif","images/menu_on.gif");
labelBullet = new bulletPoint("images/header_off.gif","images/header_on.gif");
subMenuBullet = new bulletPoint("images/sub_header_off.gif","images/sub_header_on.gif");

myTest = new menuBar('myTest',700, 'horizontal', '#000000', '#000000');
myTest.addLabel('labelBullet', 'Home', 1, 120, '#515A5A ', '#3d85c6', 'rm_home.asp', 'left');
myTest.addLabel('labelBullet', 'Master', 2, 120, '#515A5A ', '#3d85c6', '#', 'left');
myTest.addLabel('labelBullet', 'Transaction', 3, 120, '#515A5A ', '#3d85c6', '#', 'left');
myTest.addLabel('labelBullet', 'Field Services', 4, 120, '#515A5A ', '#3d85c6', '#', 'left');
myTest.addLabel('labelBullet', 'Finance', 5, 120, '#515A5A ', '#3d85c6', '#', 'left');
myTest.addLabel('labelBullet', 'Report', 6, 120, '#515A5A', '#3d85c6', 'rm_rpt.asp', 'left');
myTest.height = 38;

menus[1] = new menu(200, 'vertical', '#000000', '#000000');
menus[1].height = 38;
menus[1].writeMenu();


menus[2] = new menu(175, 'vertical', '#515A5A', '#000000');
menus[2].height = 38;
menus[2].addItem('menuItemBullet', 'Create Stock', null, 300, '#515A5A', '#3d85c6', 'rm_stock_create.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'View Stock Master', null, 300, '#515A5A', '#3d85c6', 'rm_stock_view.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'View Store', null, 300, '#515A5A', '#3d85c6', 'rm_warehouse_view.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'Create Store', null, 300, '#515A5A', '#3d85c6', 'rm_warehouse_new.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'View Technician', null, 300, '#515A5A', '#3d85c6', 'rm_contractor_view.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'Create Technician', null, 300, '#515A5A', '#3d85c6', 'rm_contractor_new.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'View Customer', null, 300, '#515A5A', '#3d85c6', 'rm_customer_view.asp?type=reset', 'left');
menus[2].addItem('menuItemBullet', 'Create Customer', null, 300, '#515A5A', '#3d85c6', 'rm_customer_new.asp?type=reset', 'left');
menus[2].writeMenu();


menus[3] = new menu(175, 'vertical', '#515A5A', '#000000');
menus[3].height = 38;
menus[3].addItem('menuItemBullet', 'Stock-in', null, 300, '#515A5A', '#3d85c6', 'rm_stockin_view.asp?type=reset', 'left');
menus[3].addItem('menuItemBullet', 'Stock-Out', null, 300, '#515A5A', '#3d85c6', 'rm_stockout_view.asp?type=reset', 'left');
menus[3].addItem('menuItemBullet', 'Stock-Transfer', null, 300, '#515A5A', '#3d85c6', 'rm_stockTfr_view.asp?type=reset', 'left');
menus[3].addItem('menuItemBullet', 'Stock-Adjustment', null, 300, '#515A5A', '#3d85c6', 'rm_stockAdj_view.asp?type=reset', 'left');
menus[3].addItem('menuItemBullet', 'Parts Order Plan', null, 300, '#515A5A', '#3d85c6', 'rm_por_edit.asp?type=reset', 'left');
menus[3].addItem('menuItemBullet', 'Parts Order Forecast', null, 300, '#515A5A', '#3d85c6', 'rm_por_reorder.asp?type=reset', 'left');
menus[3].writeMenu();

menus[4] = new menu(175, 'vertical', '#000000', '#000000');
menus[4].height = 38;
menus[4].addItem('menuItemBullet', 'View Job Sheet', null, 300, '#515A5A', '#3d85c6', 'rm_jobsheet_view.asp?type=reset', 'left');
menus[4].addItem('menuItemBullet', 'Create Job Sheet', null, 300, '#515A5A', '#3d85c6', 'rm_jobsheet.asp?type=reset', 'left');
menus[4].addItem('menuItemBullet', 'Schedule ', null, 300, '#515A5A', '#3d85c6', 'rm_schedule.asp', 'left');
menus[4].writeMenu()

menus[5] = new menu(175, 'vertical', '#000000', '#000000');
menus[5].height = 38;
menus[5].addItem('menuItemBullet', 'View Invoice', null, 300, '#515A5A', '#3d85c6', 'rm_invoice_view.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'Create Invoice', null, 300, '#515A5A', '#3d85c6', 'rm_invoice_new.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'View DO', null, 300, '#515A5A', '#3d85c6', 'rm_do_view.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'View RCN', null, 300, '#515A5A', '#3d85c6', 'rm_rcn_view.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'Create RCN', null, 300, '#515A5A', '#3d85c6', 'rm_rcn_new.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'View CN', null, 300, '#515A5A', '#3d85c6', 'rm_cn_view.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'Create CN', null, 300, '#515A5A', '#3d85c6', 'rm_cn_new.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'View Receipt', null, 300, '#515A5A', '#3d85c6', 'rm_receipt_view.asp?type=reset', 'left');
menus[5].writeMenu();

menus[6] = new menu(200, 'vertical', '#000000', '#000000');
menus[6].height = 38;
menus[6].writeMenu();

menus[1].align='left';
menus[2].align='left';
menus[3].align='left';
menus[4].align='left';
menus[5].align='left';
menus[6].align='left';

//-->
</script>