<!--#include file="../Conn.asp"-->
<!--#include file="Const.asp"-->
<!--#include file="INC/Alipay_payto.asp"-->
<%
dim service,partner,sign_type,subject,body,out_trade_no,price,show_url,quantity,seller_email,notify_url,return_url,key
dim t1,t4,t5,ExtCredits
dim AlipayObj,itemUrl
ExtCredits		=	Split(team.Club_Class(21),"|")
t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'֧���ӿ�
t4				=	"images/alipay_bwrx.gif"						'֧������ťͼƬ
t5				=	"team��̳�Ƽ�ʹ��֧��������"						'��ť��ͣ˵��
service         =   "create_digital_goods_trade_p"
partner			=	team.Forum_setting(103)						'partner�������ID�����ֶ�
sign_type       =   "MD5"
subject			=	Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)	'subject��Ʒ����
body			=	""'body��Ʒ����
out_trade_no    =	Replace(Now(),"-","")
out_trade_no    =   Trim(Replace(out_trade_no,":",""))
price		    =	HRF(2,2,"price")					'price��Ʒ����0.01��50000.00
show_url        =   "http://www.alipay.com"				'��Ʒչʾ��ַ
quantity        =   "1"									'��Ʒ����
seller_email    =   team.Forum_setting(101)				'�����˻�
key             =   team.Forum_setting(102)				'֧������ȫУ����
notify_url      =   "http://liuzhuo/alipay/Alipay_Notify.asp"
return_url      =   "http://liuzhuo/alipay/return_Alipay_Notify.asp"
Set AlipayObj	= New creatAlipayItemURL
itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,partner,sign_type,subject,body,out_trade_no,price,show_url,quantity,seller_email,notify_url,return_url,key)
Response.Redirect itemUrl
%>