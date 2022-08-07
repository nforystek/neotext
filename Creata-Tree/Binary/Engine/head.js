
var bVersion = parseInt(navigator.appVersion);
var pmPercent = 50;
var trans2x2 = new Image();


// JavaScript code for setting up the tree

function menuObject(treeTop,treeLeft,itemHeight,staticTree,blankImage,stretchBullets,usePlusMinus,plusIconSrc,minusIconSrc,useTreelines,topJointIconSrc,midJointIconSrc,endJointIconSrc,hLineIconSrc,vLineIconSrc) {
// this function sets up the base object of the tree.
// all of these functions params are required.
// the height and width of the plus and minus icons are equal and set by
//  the equation (pmPercent of this.itemHeight), pmPercent stands for
//  Plus Minus Percent which is a variable declared below..

	trans2x2.src = blankImage;
	this.staticTree = staticTree;
	this.treeTop = treeTop;
	this.treeLeft = treeLeft;
		// Starting X and Y position of the tree for painting.

	this.itemHeight=itemHeight;
		// Height of each and all items in the tree.
		
	this.stretchBullets=stretchBullets;
		// Whether to enforce height and width on bullet icons.

	this.menuItems=new Array();
		// holds all the menu items of this objects tree, of type menuItemObject.

	this.usePlusMinus=usePlusMinus;
		// this variable set to true displays the icons below as treelines.
	if (usePlusMinus==true) {
		this.plusIcon = new Image();
		this.plusIcon.src=plusIconSrc;
			// the icon used to indicate a tree may be expanded.

		this.minusIcon = new Image();
		this.minusIcon.src=minusIconSrc;
			// the icon used to indicate a tree may be retracted.
	
	} else {
		this.plusIcon = trans2x2;

		this.minusIcon = trans2x2;
	
	}


	this.useTreelines=useTreelines;
	
	if (useTreelines==true) {
		this.topJointIcon = new Image();
		this.topJointIcon.src=topJointIconSrc;
			// the icon used for the item at the top of the tree if it's not expandable.

		this.midJointIcon = new Image();
		this.midJointIcon.src=midJointIconSrc;
			// the icon used for a non expandable item thats inbetween items on the tree.

		this.endJointIcon = new Image();
		this.endJointIcon.src=endJointIconSrc;
			// the icon used for a non expandable item at the end of a tree.

		this.hLineIcon = new Image();
		this.hLineIcon.src=hLineIconSrc;
			// the icon used for horizontal tree lines.

		this.vLineIcon = new Image();
		this.vLineIcon.src=vLineIconSrc;
			// the icon used for vertical tree lines.
	} else {
		this.topJointIcon = trans2x2;

		this.midJointIcon = trans2x2;

		this.endJointIcon = trans2x2;

		this.hLineIcon = trans2x2;

		this.vLineIcon = trans2x2;
	}
}
function menuItemObject(bulletIconSrc,expandedBulletSrc,menuPicSrc,menuPicOverSrc,menuText,menuTextFont,menuTextSize,menuTextColor,menuLink,menuLinkTarget,menuDescription,isExpanded) {
// this function sets up menu items of a tree object.
// bulletIconSrc can be set to "" to disable the bullet on the menuitem.
// if there are no submenus just leave the subMenu variable alone.
// either of menuPicSrc or menuText can be set to "" if both are set
//  they layer each other.
// menuPicOverSrc is the location of a pic for the menu to replace menuPicSrc when the mouse is over the pic.
// if menuPicOverSrc is set to "" then the normal
//  menuPicSrc is used when mouse over occurs.
// the menuLink can also be set to "" to disable the linking.


	this.subMenu=new Array();
		//submenu items an array of type menuItemObject.

	if (bulletIconSrc=="") {
		this.menuBullet=trans2x2;
	} else {
		this.menuBullet = new Image();
		this.menuBullet.src=bulletIconSrc;
	}

	if (expandedBulletSrc=="") {
		this.expandedBullet=trans2x2;
	} else {
		this.expandedBullet = new Image();
		this.expandedBullet.src=expandedBulletSrc;
	}
	
	if (menuPicSrc=="") {
		this.menuPic=trans2x2;
	} else {
		this.menuPic = new Image();
		this.menuPic.src=menuPicSrc;
	}
	
	
	if (menuPicOverSrc=="") {
		this.menuPicOver=trans2x2;
	} else {
		this.menuPicOver = new Image();
		this.menuPicOver.src=menuPicOverSrc;
	}		

	if (navigator.appName=="Netscape") {this.picPtr="null";}
		//Pointer to find onMouseOver pic only once in nutscrape

	this.menuText=menuText;
	this.menuTextFont=menuTextFont;
	this.menuTextSize=menuTextSize;
	this.menuTextColor=menuTextColor;
		//text for the context of the menu item.
	
	this.menuLink=menuLink;
	this.menuLinkTarget=menuLinkTarget;
		//the HREF to link to when item is clicked.

	this.menuDescription=menuDescription;
		//ALT text of the menu item

	this.expanded=isExpanded;
		//this is a variable the code uses to show which menuItems are expanded.
		//this variable does not effect root items or items with out subitems.
}

function initImage(theImg,theSrc) {
	if (loadAll==true) {
		eval(theImg+'=new Image();');
		eval(theImg+'.src = '+theSrc+';');
		tmpstr=tmeImg+'.src';
	} else {
		eval('var '+theImg+'_src = '+theSrc+';');
	}
}
function getImage(theImg) {
	getImage=(loadAll==true)?tmeImg+'.src':tmeImg+'_src';
}


function expandTreeItem(menuItems,tItem) {
	var cnt;
	var found = false;
	for (cnt in menuItems) {
		if (!found) {
			if (menuItems[cnt].menuLink == tItem+'.html') {
				found = true;
			} else {
				found = expandTreeItem(menuItems[cnt].subMenu,tItem);
				if (found) {
					menuItems[cnt].expanded=true;
				}
			}
		}
	
	}
	return found;
}
