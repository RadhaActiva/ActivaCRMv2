<script language="JavaScript1.2" type="text/javascript">
<!--
sniffBrowsers();

menuItemBullet = new bulletPoint("images/menu_off.gif","images/menu_on.gif");
labelBullet = new bulletPoint("images/header_off.gif","images/header_on.gif");
subMenuBullet = new bulletPoint("images/sub_header_off.gif","images/sub_header_on.gif");

myTest = new menuBar('myTest',700, 'horizontal', '#000000', '#000000');
    myTest.addLabel('labelBullet', 'Dashboard', 1, 120, '#515A5A ', '#3d85c6', 'rm_home.asp', 'left');
    myTest.addLabel('labelBullet', 'Job Sheet', 2, 120, '#515A5A ', '#3d85c6', '#', 'left');
    myTest.addLabel('labelBullet', 'Spare Parts', 4, 120, '#515A5A ', '#3d85c6', '#', 'left');
    myTest.addLabel('labelBullet', 'Claims', 5, 120, '#515A5A ', '#3d85c6', '#', 'left');
myTest.height = 38;

menus[1] = new menu(200, 'vertical', '#000000', '#000000');
menus[1].height = 38;
menus[1].writeMenu();

    menus[2] = new menu(135, 'vertical', '#515A5A', '#000000');
menus[2].height = 38;
    menus[2].addItem('menuItemBullet', 'View Job Sheet', null, 300, '#515A5A ', '#3d85c6', 'rmtech_jobsheet_view.asp?type=reset', 'left');
    menus[2].addItem('menuItemBullet', 'View Job Schedule', null, 300, '#515A5A ', '#3d85c6', 'rm_schedule.asp', 'left');
menus[2].writeMenu();

menus[3] = new menu(200, 'vertical', '#000000', '#000000');
menus[3].height = 38;
menus[3].writeMenu();


menus[4] = new menu(175, 'vertical', '#000000', '#000000');
menus[4].height = 38;
menus[4].addItem('menuItemBullet', 'View Spare Parts Balance', null, 300, '#515A5A ', '#3d85c6', 'rm_contractor_stockbalance.asp?type=reset', 'left');
menus[4].writeMenu();

menus[5] = new menu(195, 'vertical', '#000000', '#000000');
menus[5].height = 38;
menus[5].addItem('menuItemBullet', 'Monthly Claims', null, 300, '#515A5A ', '#3d85c6', 'rmtech_claims.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'View Submitted Job Schedule', null, 300, '#515A5A ', '#3d85c6', 'rmtech_rpt_weekly_job_submission.asp?type=reset', 'left');
menus[5].addItem('menuItemBullet', 'Download Claims', null, 300, '#515A5A ', '#3d85c6', 'rmtech_view_claim_form.asp?type=reset', 'left');
menus[5].writeMenu();

menus[1].align='left';
menus[2].align='left';
menus[3].align='left';
menus[4].align = 'left';
menus[5].align = 'left';



//-->
</script>