function LoadSubMenu() {
   var subMenu = new Array(2)
 
	subMenu[0] = '<font size="2" class="BabyBlue" face="Verdana"><b>';    
	subMenu[0] += '<a class="LightBlue" href="report1.aspx" target="main">Report1</a> | ';
	subMenu[0] += '<a class="LightBlue" href="reports/Report2.aspx" target="main">Report2</a> | ';
	subMenu[0] += '<a class="LightBlue" href="Construction.htm" target="main">Report3</a> | ';
	subMenu[0] += '<a class="LightBlue" href="Construction.htm" target="main">Report4</a> | ';
	subMenu[0] += '<a class="LightBlue" href="Construction.htm" target="main">Report5</a>';
	subMenu[0] += '</b></font>';
	
	subMenu[1]='';
	
   return subMenu;
}

function ShowSubMenu(index){

	var subMenu = LoadSubMenu();
	var menu = document.getElementById("SubMenuHolder");
	if (index >= 0) {
		menu.innerHTML = subMenu[index];
	}
}