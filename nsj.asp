<%@Language="VBScript" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <title>NSJSoft</title>
    <link rel="stylesheet" type="text/css" href="css/joint.css" />
    <script src="js/jquery.js"></script>
    <script src="js/lodash.js"></script>
    <script src="js/backbone.js"></script>
    <script src="js/joint.js"></script>
</head>
<body>
<%
    Response.CharSet="UTF-8"
    Response.Codepage = 65001

    Dim patternText : patternText = "[*/+-]"
    Dim str_cal : str_cal = "L1+L2-L3*L4/L5"
    Dim regex, ResultReg, tmp_arr

    Set regex = New RegExp

    regex.Pattern = patternText
    regex.IgnoreCase = True
    regex.Global = True

    Set Matches = regex.Execute(str_cal)
    tmp_arr = Split(regEx.Replace(str_cal, ","), ",")
    min_arr = tmp_arr
    Dim right_side, left_side, pre_operand
    Response.Write "min_arr=>" & UBound(min_arr) & "<br>"

    For i = 0 To Ubound(min_arr) - 1
        Response.Write "i=>" & i & "<br>"
    Next

    Response.Write "For End i=>" & i & "<br>"
    Response.Write "=============================<br>"

    pre_operand = ""
    Response.Write "Matches.Count=>" & Matches.Count & "<br>"

    IF Matches.Count <> 0 THEN
        For i = 0 To  Matches.Count - 1   
            left_side  = tmp_arr(i)
            right_side = tmp_arr(i+1)   
            IF pre_operand = "" THEN              
               pre_operand = "+"
            ELSE
               'Response.Write "Matches(i-1).Value=>" & i - 1 & ", " & Matches(i-1).Value & "<br>"
               pre_operand = Matches(i-1).Value
            End IF

            Response.Write "pre_operand=>" & pre_operand & "<br>" 
        Next
    End IF

    Dim aTest : aTest = Array("1", "2")
    Response.Write "UBound(aTest)=>" & UBound(aTest) & "<br>"

    Dim li_cal, li_cal2
    Dim ls_tmp
    
    ls_tmp = "0.0+1.0+0.0+1.0-0.258*5.0-0.572*5.0"
    execute("li_cal  = CDBL(FormatNumber(" & ls_tmp & " + 0, 5))")
    li_cal2 = CDBL(FormatNumber(li_cal,  1))

    Response.Write "li_cal=>" & li_cal & "<br>"
    Response.Write "li_cal2=>" & li_cal2 & "<br>"

    Dim X

    X = 8/2*(2+2)
    '<font size="5" color="red" face="돋움"><strong> 또는 <b>
    'css => font-style font-variant font-weight font-size/line-height font-family
    'font-style : normal, italic
    'font-variant : normal과 small-caps
    'font-weight : normal, lighter(100), normal(400), bold(600), bolder(900)
    'font-size : px, em
    'font-family : font face속성과 동일
    'line-height : font-size 바로 다음에 와야하며 "/"로 구분. 예:16px/3px
    Response.Write "<br><font color='red' size='5'><b>X=>" & X & "</b></font><br>"

%>
</body>
</html>
