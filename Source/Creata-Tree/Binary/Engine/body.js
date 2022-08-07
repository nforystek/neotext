
//*********************************************************************************************************
//Browsers 4.0 and above (Animated engine)
//*********************************************************************************************************

if ((bVersion > 3)&&(menuBase.staticTree==false)) {

	var itemNum=0;
	var itemsPtr=new Array();
	var cTop=0;
	var cLeft=0;
	var itemCode="";

	function reverseString(theStr) {
		tmpStr="";
		var cnt;
		for (cnt=theStr.length-1;cnt>=0;cnt--) {
			tmpStr=tmpStr+theStr.substring(cnt,cnt+1);
		}
		return tmpStr;
	}
	function pushNum(theNum,newNum) {
		theNum=theNum+"_"+newNum;
		return theNum;
	}
	function popNum(theNum) {
		theNum=reverseString(theNum);
		var start=theNum.indexOf("_",0);
		var end=theNum.length-1;
		theNum=theNum.substring(start,end);
		theNum=reverseString(theNum);
		return theNum;
	}
	function menuItemChange(itemNum,newPic) {
		if (navigator.appName=="Netscape") {
			if (itemsPtr[itemNum].picPtr=="null") {
				var cnt;
				var cnt2;
				var theImg;
				for(cnt=0;cnt<document.layers[0].document.layers.length;cnt++) {
					for(cnt2=0;cnt2<eval('document.layers[0].document.layers['+cnt+'].document.images.length');cnt2++) {
						if (eval('document.layers[0].document.layers['+cnt+'].document.images['+cnt2+'].name')==('imageNum'+itemNum)) {
							theImg=eval('document.layers[0].document.layers['+cnt+'].document.images[cnt2]');
						}
					}
				}
				itemsPtr[itemNum].picPtr=theImg;
			}
				itemsPtr[itemNum].picPtr.src=newPic.src;
		} else {
			eval('imageNum'+itemNum+'.src = newPic.src;');
		}
	}

	function menuItemClick(itemNum) {
		//The fallowing code thats commented out is for trees with one subitem, 
		//the code will unexpand all but the clicked item.  The line of code that
		//isn't commented out is the regular style.

		itemsPtr[itemNum].expanded= (!itemsPtr[itemNum].expanded);

		//var cnt;
		//for (cnt in itemsPtr) {
		//	if (itemNum==cnt) {
		//		itemsPtr[cnt].expanded= (!itemsPtr[cnt].expanded);
		//	} else {
		//		itemsPtr[cnt].expanded= false;
		//	}
		//}
		refreshTree();
	}
	function paintMenu(newCode) {
		if (document.getElementById)
		{
			A = document.getElementById("tree");
			A.innerHTML = '';
			A.innerHTML = newCode;
		}
		else if (document.all)
		{
			A = document.all["tree"];
			A.innerHTML = newCode;
		}
		else if (document.layers)
		{
			A = document.layers["tree"];
			text2 = '&lt;P CLASS="testclass"&gt;' + newCode + '&lt;/P&gt;';
			A.document.open();
			A.document.write(text2);
			A.document.close();
		}
	}
	function addLink(menuItem,forceLink,capCode) {
		var tmpCode="";
		tmpCode=tmpCode+'<A HREF="';

		if (menuItem.menuLink!="") {
			tmpCode=tmpCode+menuItem.menuLink;
		} else {
			tmpCode=tmpCode+'#';	
		}
		tmpCode=tmpCode+'"';
		
		if (menuItem.menuLinkTarget!="") {
			tmpCode=tmpCode+'TARGET="'+menuItem.menuLinkTarget+'"';
		}
		if (menuItem.menuPicOver.src!="") {
			tmpCode=tmpCode+' OnMouseOver="menuItemChange(';
			tmpCode=tmpCode+"'"+itemNum+"',itemsPtr["+itemNum+"].menuPicOver";
			tmpCode=tmpCode+');return true;" OnMouseOut="menuItemChange(';
			tmpCode=tmpCode+"'"+itemNum+"',itemsPtr["+itemNum+"].menuPic";
			tmpCode=tmpCode+');return true;"';
		}
		if (menuItem.subMenu.length>0) {
			if ((!forceLink)||(menuItem.menuLink=="")) {
				tmpCode=tmpCode+' OnClick="menuItemClick(';
				tmpCode=tmpCode+"'"+itemNum+"'";
				tmpCode=tmpCode+');return false;"';
			}
		}
		tmpCode=tmpCode+'>';
		
		tmpCode=tmpCode+capCode;

		tmpCode=tmpCode+'</A>';
		return tmpCode;

	}
	function addSpan(left,top,height,width,capCode) {
		var tmpWidth = (width==""?'':'width:'+width+'; ');
		capCode='<SPAN STYLE="position: absolute; left:'+left+'; top:'+top+'; '+tmpWidth+'height:'+height+'">'+capCode;
		capCode=capCode+'</SPAN>';
		return capCode;
	}
	function addBulletSpan(left,top,height,width,bullet,capCode) {
		var padTop = top;//+Math.round( ((height-bullet.height)/2) );
		var padLeft = left;//+Math.round( ((width-bullet.width)/2) );
		capCode='<SPAN STYLE="position: absolute; left:'+ padLeft +'; top:'+padTop+'; width:'+width+'; height:'+height+'">'+capCode;
		capCode=capCode+'</SPAN>';
		return capCode;
	}

	var plusSize=0;
	var lineSize=0;
	var headerCnt=0;
	if (menuBase.stretchBullets==true) {
	//	var bulletStyle='HEIGHT='+Math.round( (menuBase.itemHeight*.75) )+' WIDTH='+Math.round( (menuBase.itemHeight*.75) )+' ';
		var bulletStyle='HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+' ';
	} else {
		var bulletStyle='';
	}

	function refreshTree2(menuItem,headerLines) {
		var cnt;
		var cnt2;
		var tmpCode;
		var currentHeader;
		var tmpIndent;
		for (cnt in menuItem) {
			tmpCode="";
			currentHeader="";	
			if ((menuBase.usePlusMinus==true)||(menuBase.useTreelines==true)) {
				if ((menuItem[cnt].subMenu.length>0)&&(menuBase.usePlusMinus==true)) {
					//Item as plus/minus
					if (menuItem[cnt].expanded==true) {
						expIcon=menuBase.minusIcon.src;
					} else {
						expIcon=menuBase.plusIcon.src;
					}
					if ((cnt==0) && (itemNum==0)){
						currentHeader=currentHeader+addSpan(cLeft,cTop,lineSize,menuBase.itemHeight,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+menuBase.itemHeight+'>');
					} else {
						currentHeader=currentHeader+addSpan(cLeft,cTop,lineSize,menuBase.itemHeight,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'><IMG SRC="'+menuBase.vLineIcon.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+plusSize+'><IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'>');
					}
					currentHeader=currentHeader+addSpan(cLeft,cTop+lineSize,plusSize,menuBase.itemHeight,addLink(menuItem[cnt],false,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+plusSize+' WIDTH='+lineSize+'><IMG SRC="'+expIcon+'" BORDER=0 HEIGHT='+plusSize+' WIDTH='+plusSize+'><IMG SRC="'+menuBase.hLineIcon.src+'" BORDER=0 HEIGHT='+plusSize+' WIDTH='+lineSize+'>'));
					if (cnt==menuItem.length-1)  {
						currentHeader=currentHeader+addSpan(cLeft,cTop+lineSize+plusSize,lineSize,menuBase.itemHeight,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+menuBase.itemHeight+'>');
					} else {
						currentHeader=currentHeader+addSpan(cLeft,cTop+lineSize+plusSize,lineSize,menuBase.itemHeight,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'><IMG SRC="'+menuBase.vLineIcon.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+plusSize+'><IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'>');
					}
				} else {
					//Item does not have plus/minus

					if ((cnt==0) && (itemNum==0)) {
							if (menuItem.length>1) {
								currentHeader=currentHeader+addSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,'<IMG SRC="'+menuBase.topJointIcon.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+'>');
							} else {
								currentHeader=currentHeader+addSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,'<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+'>');
							}
					} else {
						if (cnt==menuItem.length-1) {
								currentHeader=currentHeader+addSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,'<IMG SRC="'+menuBase.endJointIcon.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+'>');
						} else {
								currentHeader=currentHeader+addSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,'<IMG SRC="'+menuBase.midJointIcon.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+'>');
						}
					}
				}
				for (cnt2=0;cnt2<headerCnt;cnt2++) {
					tmpCode=tmpCode+addSpan(menuBase.treeLeft+(cnt2*menuBase.itemHeight),cTop,menuBase.itemHeight,menuBase.itemHeight,headerLines[cnt2])
				}

				tmpCode=tmpCode+currentHeader;
				cLeft=cLeft+menuBase.itemHeight;
			}

			if (menuItem[cnt].menuBullet.src!=trans2x2.src) {
				if ((menuItem[cnt].expanded==true) && (menuItem[cnt].expandedBullet.src!=trans2x2.src)) {
					tmpCode=tmpCode+addBulletSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,menuItem[cnt].expandedBullet,addLink(menuItem[cnt],false,'<IMG '+bulletStyle+'SRC="'+menuItem[cnt].expandedBullet.src+'" BORDER=0>'));
				} else {
					tmpCode=tmpCode+addBulletSpan(cLeft,cTop,menuBase.itemHeight,menuBase.itemHeight,menuItem[cnt].menuBullet,addLink(menuItem[cnt],false,'<IMG '+bulletStyle+'SRC="'+menuItem[cnt].menuBullet.src+'" BORDER=0>'));
				}
				cLeft=cLeft+menuBase.itemHeight;
				
			}
			
			cLeft=cLeft+2;
			if (menuItem[cnt].menuPic.src!="") {
				tmpCode=tmpCode+addSpan(cLeft,cTop,menuBase.itemHeight,'',addLink(menuItem[cnt],true,'<IMG SRC="'+menuItem[cnt].menuPic.src+'" NAME="imageNum'+itemNum+'" BORDER=0>'));
				}
			if (menuItem[cnt].menuText!="") {
				tmpCode=tmpCode+addSpan(cLeft,cTop,menuBase.itemHeight,'',addLink(menuItem[cnt],true,'<NOBR><FONT STYLE="font-family: '+ menuItem[cnt].mnuTextFont+';" SIZE="'+menuItem[cnt].menuTextSize+'" COLOR="'+menuItem[cnt].menuTextColor+'">'+menuItem[cnt].menuText+'</FONT></NOBR>'));
				}
			if (menuItem[cnt].menuBullet.src!=trans2x2.src) {
				cLeft=cLeft-menuBase.itemHeight;
			}
			cLeft=cLeft-2;
			
			itemCode=itemCode+tmpCode;	
			if (navigator.appName=="Netscape") {menuItem[cnt].picPtr="null";}
			itemsPtr[itemNum]=menuItem[cnt];
			itemNum++;
			cTop=cTop+menuBase.itemHeight;		
			if ((menuBase.usePlusMinus==true)||(menuBase.useTreelines==true)) {cLeft=cLeft-menuBase.itemHeight};
			if ((menuItem[cnt].subMenu.length>0) && (menuItem[cnt].expanded==true)) {
				if ((menuBase.usePlusMinus==true)||(menuBase.useTreelines==true)) {
					tmpIndent="";
					if (cnt==menuItem.length-1)  {
						tmpIndent='<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+menuBase.itemHeight+'>';
					} else {
						tmpIndent='<IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'><IMG SRC="'+menuBase.vLineIcon.src+'" BORDER=0 HEIGHT='+menuBase.itemHeight+' WIDTH='+plusSize+'><IMG SRC="'+trans2x2.src+'" BORDER=0 HEIGHT='+lineSize+' WIDTH='+lineSize+'>';
					}
					headerLines[headerCnt] =tmpIndent;
					headerCnt++;
				}
				cLeft=cLeft+menuBase.itemHeight;
				refreshTree2(menuItem[cnt].subMenu,headerLines);
				if ((menuBase.usePlusMinus==true)||(menuBase.useTreelines==true)) {
					headerLines[headerCnt]="";
					headerCnt--;
				}
				cLeft=cLeft-menuBase.itemHeight;
			}
		}
	}

	function refreshTree() {
	//this function is called to refresh a tree initialized by paintTree().
		var headerLines=new Array();
		itemNum=0;
		cTop=menuBase.treeTop;
		cLeft=menuBase.treeLeft;
		plusSize = Math.round(((pmPercent /100) * menuBase.itemHeight));
		lineSize = Math.round( ((menuBase.itemHeight-plusSize)/2) );
		itemCode="";
	        refreshTree2(menuBase.menuItems,headerLines);
		paintMenu(itemCode);
	        itemCode="";
		itemNum=0;
		return true;
	}

	if (navigator.appName=="Netscape") {
		var nutScrape=1;
		function reTime() {
			if (nutScrape==1) {
				nutScrape=2;
				document.location.reload();
			} else {
				if (nutScrape==2) {nutScrape=1}
			}
			return false;
		}
	}




	document.write('<SPAN ID=tree STYLE="position: absolute; left:'+menuBase.treeLeft+'; top:'+menuBase.treeTop+'; width:100%; height:100%"> </SPAN>');

	//show refresh the tree

	window.onload=refreshTree;

	if (navigator.appName=="Netscape") {
		window.onResize=reTime;
	}

	//*********************************************************************************************************
	//Browsers 3.0 and below (Nonanimated engine) (Note: Not perfected but usefull)
	//*********************************************************************************************************

} else {

	//*********************************************************************************************************
	//Browsers 3.0 and below (Nonanimated engine) (Note: Not perfected but usefull)
	//*********************************************************************************************************

	function naPaintTree(menuItems,indent) {
		var cnt;
		var nextLine;
		for (cnt in menuItems) {


			nextLine=indent+'<FONT';

			if (menuItems[cnt].menuTextColor!="") {
				nextLine+=' COLOR="'+menuItems[cnt].menuTextColor+'"';
			}
			if (menuItems[cnt].menuTextSize!="") {
				nextLine+=' SIZE="'+menuItems[cnt].menuTextSize+'"';
			}

			nextLine+='>';

			if (menuItems[cnt].menuLink!="") {
				nextLine+='<A HREF="'+menuItems[cnt].menuLink+'" TARGET="'+menuItems[cnt].menuLinkTarget+'">';
			}

			if (menuItems[cnt].menuBullet.src!="") {
				nextLine+='<IMG SRC="'+menuItems[cnt].menuBullet.src+'" BORDER=0>&nbsp;';
			}

			if (menuItems[cnt].menuPic.src!="") {
				nextLine+='<IMG SRC="'+menuItems[cnt].menuPic.src+'" BORDER=0>';
			} else {
				nextLine+=menuItems[cnt].menuText;
			}

			
			if (menuItems[cnt].menuLink!="") {
				nextLine+='</A>';}


			nextLine+='</FONT><BR>';

			document.write(nextLine);

			naPaintTree(menuItems[cnt].subMenu,'&nbsp;&nbsp;&nbsp;&nbsp;'+indent);
		}
	}

	naPaintTree(menuBase.menuItems,"");

}
