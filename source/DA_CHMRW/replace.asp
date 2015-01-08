<%
'格式化字符串
function replace_h(text_in)
        'if isnull(text_in) then text_in=""
        'text_in=replace(text_in,chr(32),"&amp;nbsp;")
		text_in=replace(text_in,"'","‘")
		text_in=replace(text_in,chr(32),"&nbsp;")
        text_in=replace(text_in,chr(13),"<br>")
        replace_h=text_in
end function

'反格式化字符串
function replace_t(text_in)
        'if isnull(text_in) then text_in=""
        'text_in=replace(text_in,"&amp;nbsp;",chr(32))
		text_in=replace(text_in,"&nbsp;",chr(32))
        text_in=replace(text_in,"<br>",chr(13))
        replace_t=text_in
end function
%>