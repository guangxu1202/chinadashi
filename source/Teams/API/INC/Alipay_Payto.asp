<!--#include file="Alipay_md5.asp"-->
<%
Class creatAlipayItemURL
	'������յĹ���url
	Public Function creatAlipayItemURL(t1,t4,t5,service,partner,sign_type,subject,body,out_trade_no,price,show_url,quantity,seller_email,notify_url,return_url,key)
		Dim itemURL,count,mystr,i,minmax,minmaxSlot,j,mark,temp,Value,md5str,sign
		dim INTERFACE_URL,imgsrc,imgtitle
		'��ʼ������Ҫ����
		INTERFACE_URL	= t1	'֧���ӿ�
		imgsrc			= t4	'֧������ťͼƬ
		imgtitle		= t5	'��ť��ͣ˵��
		'Add by sunzhizhi 2006-5-10
		Count = 10
		mystr = Array("service="&service,"partner="&partner,"subject="&subject,"body="&body,"out_trade_no="&out_trade_no,"price="&price,"show_url="&show_url,"quantity="&quantity,"seller_email="&seller_email,"notify_url="&notify_url,"return_url="&return_url)
		Count=ubound(mystr)
		For i = Count TO 0 Step -1
			minmax = mystr( 0 )
			minmaxSlot = 0
			For j = 1 To i
				mark = (mystr( j ) > minmax)
				If mark Then 
					minmax = mystr( j )
					minmaxSlot = j
				End If
			Next
			If minmaxSlot <> i Then		   
				temp = mystr( minmaxSlot )
				mystr( minmaxSlot ) = mystr( i )
				mystr( i ) = temp
			End If
		Next
		For j = 0 To Count Step 1
			value = SPLIT(mystr( j ), "=")
			If  value(1)<>"" then
				If j=Count Then
					md5str= md5str&mystr( j )
				Else 
					md5str= md5str&mystr( j )&"&"
				End If 
			End If 
		Next
		md5str=md5str&key
		sign=md5(md5str)
		itemURL	= itemURL&INTERFACE_URL 
		For j = 0 To Count Step 1
			value = SPLIT(mystr( j ), "=")
			If  value(1)<>"" then
			itemURL= itemURL&mystr( j )&"&"
			End If 	     
		Next
		itemURL	= itemURL&"sign="&sign&"&sign_type="&sign_type
		creatAlipayItemURL	= get_AlipayButtonURL	(itemURL,imgtitle,imgsrc)
	End Function
	Public Function get_AlipayButtonURL(itemURL,imgtitle,imgsrc)
		dim responseText1
		responseText1	= itemURL
		get_AlipayButtonURL=responseText1
	End Function
End Class
%>