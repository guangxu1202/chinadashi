<% 
Set Jpeg = Server.CreateObject("Persits.Jpeg") '调用组件 
Path = Server.MapPath("images") & "\clock.jpg" '待处理图片路径 
Jpeg.Open Path '打开图片 
'高与宽为原图片的1/2 
Jpeg.Width = Jpeg.OriginalWidth / 2 
Jpeg.Height = Jpeg.OriginalHeight / 2 
'保存图片 
Jpeg.Save Server.MapPath("images") & "\clock_small.jpg" 
%> 
<IMG src="images/clock_small.jpg" > 查看处理的图片 