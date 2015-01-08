<html>
<head>
<title>实现图片无缝垂直滚动的JS代码 - www.webdm.cn</title>
</head>
<body>
<div id=ddmm style=overflow:hidden;height:500;width:180;>  
<div id=ddmm1>  
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >  
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
    <img src="http://www.webdm.cn/themes/pic/webdm_logo.gif" >
 </div>  
 <div id=ddmm2></div>  
 
<%
response.write "cname"
%>
 </div>  
   <script>  
   var speed=30  
   ddmm2.innerHTML=ddmm1.innerHTML
   function Marqueedd(){    
if(ddmm2.offsetTop-ddmm.scrollTop<=0)    
ddmm.scrollTop-=ddmm1.offsetHeight 
else{  
ddmm.scrollTop++  
   }  
   }  
   var MyMardd=setInterval(Marqueedd,speed)
   ddmm.onmouseover=function() {clearInterval(MyMardd)}  
   ddmm.onmouseout=function(){MyMardd=setInterval(Marqueedd,speed)}  
</script>
</body>
</html>
  <script language=javascript>

<!--

var index = 7



text = new Array(6);

text[0] ='images/ev1.jpg'

text[1] ='images/ev2.jpg'

text[2] ='images/ev3.jpg'

text[3] ='images/ev4.jpg'

text[4] ='images/ev5.jpg'

text[5] ='images/ev6.jpg'

text[6] ='images/ev7.jpg'



document.write ("<marquee scrollamount='1' scrolldelay='50' direction= 'up' width='180' height='500' style='display:inline;float:left;margin-left:10px; margin-right:-10px;'>");

 

for (i=0;i<index;i++){
document.write ("<img src='"+text[i]+"' style='margin-top:5px; border:2px solid #999;'>");

}

document.write ("</marquee>")

// -->

</script>
</body>
</html>
