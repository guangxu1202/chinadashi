<% 
Set Jpeg = Server.CreateObject("Persits.Jpeg") '������� 
Path = Server.MapPath("images") & "\clock.jpg" '������ͼƬ·�� 
Jpeg.Open Path '��ͼƬ 
'�����ΪԭͼƬ��1/2 
Jpeg.Width = Jpeg.OriginalWidth / 2 
Jpeg.Height = Jpeg.OriginalHeight / 2 
'����ͼƬ 
Jpeg.Save Server.MapPath("images") & "\clock_small.jpg" 
%> 
<IMG src="images/clock_small.jpg" > �鿴�����ͼƬ 