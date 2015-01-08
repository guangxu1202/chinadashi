<!--#include file="../Conn.asp"-->
<!--#include file="Const.asp"-->
<!--#include file="INC/Alipay_payto.asp"-->
<%
dim service,partner,sign_type,subject,body,out_trade_no,price,show_url,quantity,seller_email,notify_url,return_url,key
dim t1,t4,t5,ExtCredits
dim AlipayObj,itemUrl
ExtCredits		=	Split(team.Club_Class(21),"|")
t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'支付接口
t4				=	"images/alipay_bwrx.gif"						'支付宝按钮图片
t5				=	"team论坛推荐使用支付宝付款"						'按钮悬停说明
service         =   "create_digital_goods_trade_p"
partner			=	team.Forum_setting(103)						'partner合作伙伴ID保留字段
sign_type       =   "MD5"
subject			=	Split(ExtCredits(Cid(team.Forum_setting(99))),",")(0)	'subject商品名称
body			=	""'body商品描述
out_trade_no    =	Replace(Now(),"-","")
out_trade_no    =   Trim(Replace(out_trade_no,":",""))
price		    =	HRF(2,2,"price")					'price商品单价0.01～50000.00
show_url        =   "http://www.alipay.com"				'商品展示地址
quantity        =   "1"									'商品数量
seller_email    =   team.Forum_setting(101)				'卖家账户
key             =   team.Forum_setting(102)				'支付宝安全校验码
notify_url      =   "http://liuzhuo/alipay/Alipay_Notify.asp"
return_url      =   "http://liuzhuo/alipay/return_Alipay_Notify.asp"
Set AlipayObj	= New creatAlipayItemURL
itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,partner,sign_type,subject,body,out_trade_no,price,show_url,quantity,seller_email,notify_url,return_url,key)
Response.Redirect itemUrl
%>