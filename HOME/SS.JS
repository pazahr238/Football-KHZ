//----------DHTML Menu Created using AllWebMenus ver 2.0-#456---------------
var awmMenuName='ss';
var awmLibraryPath='/awmData-ss';
var awmImagesPath='/awmData-ss';
var awmSupported=(navigator.appName + navigator.appVersion.substring(0,1)=="Netscape5" || document.all || document.layers)?1:0;
if (awmAltUrl!='' && !awmSupported) window.location.replace(awmAltUrl);
if (awmSupported){
var awmMenuPath;
if (document.all) mpi=document.all['awmMenuPathImg-ss'].src;
if (document.layers) mpi=document.images['awmMenuPathImg-ss'].src;
if (navigator.appName + navigator.appVersion.substring(0,1)=="Netscape5") mpi=document.getElementById('awmMenuPathImg-ss').src;
awmMenuPath=mpi.substring(0,mpi.length-16);
var nua=navigator.userAgent,scriptNo=(nua.indexOf('Gecko')>-1)?2:((document.layers)?3:((nua.indexOf('Opera')>-1)?4:((nua.indexOf('Mac')>-1)?5:1)));
document.write("<SCRIPT SRC='"+awmMenuPath+awmLibraryPath+"/awmlib"+scriptNo+".js'><\/SCRIPT>");
var n=null;
awmzindex=1000;
}

var awmSubmenusFrame='';
var awmSubmenusFrameOffset;
var awmOptimize=0;
function awmBuildMenu(){
if (awmSupported){
awmImagesColl=['home.jpg',120,32,'home_s.jpg',120,32,'sitem7-1.gif',115,23,'sitem7-2.gif',115,23,'sitem7-3.gif',115,23,'public.jpg',118,33,'public_s.jpg',118,33];
awmCreateCSS(1,2,1,n,n,n,'0mm sans-serif',n,'none',2,n,0,4)
awmCreateCSS(0,1,0,n,n,n,n,n,'none',2,n,0,0);
awmCreateCSS(1,2,1,'#000000',n,n,'4mm sans-serif',n,'none',1,n,0,1)
awmCreateCSS(0,2,1,'#FFFFFF',n,n,'4mm sans-serif',n,'none',1,n,0,1)
awmCreateCSS(0,2,1,'#FFFFFF','#408080',n,'4mm sans-serif',n,'none',1,n,0,1)
awmCreateCSS(0,2,1,'#FFFFFF','#000080',n,'4mm sans-serif',n,'none',1,n,0,1)
awmCreateCSS(1,2,1,n,n,n,'4mm sans-serif',n,'none',2,n,0,0)
awmCreateCSS(0,1,0,n,n,n,n,n,'outset',2,n,0,0);
awmCreateCSS(1,2,1,n,n,n,'4mm sans-serif',n,'none',1,n,0,1)
var s0=awmCreateMenu(0,0,0,0,1,0,1,2,1,0,0,0,0,1,0,n,"",n,1,0,1,0,n,n);
it=s0.addItemWithImages(2,3,3,"",n,n,"",0,1,n,0,0,0,n,n,n,"",n,n,n,n,"new");
var s1=it.addSubmenu(0,0,-4,0,0,0,0,7,6,"","",n,1,0,1,0,n,n);
it=s1.addItemWithImages(8,3,3,"",n,n,"",2,2,2,0,0,0,n,n,n,"",n,n,n,n,n);
it=s1.addItemWithImages(2,3,4,"",n,n,"",3,3,3,0,0,0,n,n,n,"",n,n,n,n,n);
it=s1.addItemWithImages(2,3,4,"",n,n,"",4,4,4,0,0,0,n,n,n,"",n,n,n,n,n);
it=s0.addItemWithImages(2,3,4,"",n,n,"",5,6,n,0,0,0,n,n,n,"",n,n,n,n,n);
it=s0.addItem(2,5,4,"",n,n,"","",n,n,n,"http://www.likno.com/?source=app","new");
it=s0.addItem(2,5,4,"",n,n,"","",n,n,n,"http://www.likno.com/?source=app","new");
s0.pm.buildMenu();
}}
