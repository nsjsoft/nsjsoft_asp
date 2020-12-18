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
<!-- #include Virtual = "/dBHelper.inc.asp" -->
<!-- #include Virtual = "/jsonObject.inc.asp" -->
<%
 
  Response.CharSet="UTF-8"
  Response.Codepage = 65001
  'Session.codepage="65001"
  'Response.codepage="65001"
  'Response.ContentType="text/html;charset=utf-8"

Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
If IsObject(conn) Then
    Response.Write("객체입니다.<br>")
Else
    Response.Write("객체가 아닙니다.<br>")
End If
'Provider=SQLOLEDB;User ID=아이디;Password=패스워드;Initial Catalog=데이터베이스이름;Data Source=데이터베이스서버이름"
str = "Provider=SQLOLEDB;User ID=uwizconv2;Password=upwd20#@3;Initial Catalog=OrpWiselection;Data Source=wfdbsvr.database.windows.net"
conn.open str

Response.Cookies("__ORP_COOKIE_INFO__")("COOKIE_P_IDX")="1"

Dim tmp : tmp = Request.Cookies("__ORP_COOKIE_INFO__")("COOKIE_P_IDX")

Response.Write(tmp & "<br>")

'Dim SESSION_COMPANY_IDX : SESSION_COMPANY_IDX = Request.ServerVariables("HTTP_SITECOMPANYIDX")

'Response.Write(SESSION_COMPANY_IDX & "<br>")

Function getParameter(m, s)
    if m = "" or IsNull(m) then
       getParameter = Trim(s)
    else
       getParameter = Trim(m)
    end if
End Function

Public Function Params(strKey, defStr)
    Params = getParameter(objParamsDict.item(strKey), defStr)
End Function

Set DBHelper = new clsDBHelper

Dim objRs
Dim paramInfo()
Dim statusCode : statusCode = "00"
Dim msg:msg = ""

ReDim paramInfo(3)
paramInfo(0)   = DBHelper.MakeParam("@P_IDX"  ,    adInteger, adParamInput,   ,   "84" )
paramInfo(1)   = DBHelper.MakeParam("@PN_IDX" ,    adInteger, adParamInput,   ,   "107")
paramInfo(2)   = DBHelper.MakeParam("@PS_IDX" ,    adInteger, adParamInput,   ,   "309")
paramInfo(3)   = DBHelper.MakeParam("@PWTU_SN",    adVarWChar, adParamInput, 20,  "01-000411")

Set objRs = DBHelper.ExecSPReturnRS("usp_Web_eTest_Exam_Info_List", paramInfo, Nothing)

IF (objRs.BOF OR objRs.EOF) Then ''
    statusCode="01"
    msg="Not Exists User Exam Information!(get page)"
    Response.Write(statusCode & "<br>" & msg & "<br>")    
Else
    Response.Write(objRs("PRE_SEQ") & "<br>" & objRs("SEQ") & "<br>" & objRs("NEXT_SEQ") & "<br>")    
End IF

Set objRs=Nothing

Dim larr_opt : larr_opt = Array("CRT","STC","CAL","RAT","TOT","CRB","HMP")
Response.Write Lbound(larr_opt) & "<br>" & Ubound(larr_opt) & "<br>" 
For i = 0 To Ubound(larr_opt)
    Response.Write larr_opt(i) & "<br>"
Next

Dim li_cnt : li_cnt = 0
ReDim paramInfo(2) 'Preserve
Dim objDic : Set objDic = Server.CreateObject("Scripting.Dictionary")

paramInfo(0) = DBHelper.MakeParam("@P_IDX" , adInteger, adParamInput, , "103")
paramInfo(1) = DBHelper.MakeParam("@PN_IDX", adInteger, adParamInput, , "147")
paramInfo(2) = DBHelper.MakeParam("@PS_IDX", adInteger, adParamInput, , "479")

Set objRs = DBHelper.ExecSPReturnRS("usp_Web_eTest_Get_ToolsCode", paramInfo, Nothing)

If Not (objRs.BOF OR objRs.EOF) Then
    While Not objRs.EOF
        objDic.Add getParameter(objRs("PST_IDX"), "0"), Array(getParameter(objRs("TW_CODE")          ,""), _
                                                              getParameter(objRs("PWOT_EXAM_TYPE")   ,""), _
                                                              getParameter(objRs("TW_PECENTILE_CD")  ,""), _
                                                              getParameter(objRs("TW_TOTAL_SCORE_CD"),""))
        objRs.MoveNext
    Wend
End IF

Dim arrKey : arrKey = objDic.Keys
Dim arrItem : arrItem = objDic.Items
Dim tmpArr
For i = 0 To objDic.Count - 1
    'Response.Write arrKey(i) & " : " & arrItem(i) & "<br>" & Chr(13)
    tmpArr = objDic.Item(arrKey(i))
    Response.Write tmpArr(0) & "=" & Lbound(tmpArr) & "=" & Ubound(tmpArr) & "<br>"
Next

'Response.Write objDic.Count & "<br>"
objRs.Close()

Dim objCrt_Dic : Set objCrt_Dic = Server.CreateObject("Scripting.Dictionary")
objCrt_Dic.Add 1, "nsjsoft"
'objCrt_Dic.Add 1, "nsj"
arrKey = objCrt_Dic.Keys
Response.Write arrKey(0) & "<br>"
Response.Write objCrt_Dic.Item(1) & "<br>"
Response.Write "varType : " & varType(arrKey(0)) & "<br>"

objCrt_Dic.Key(1) = "1"
arrKey = objCrt_Dic.Keys
Response.Write arrKey(0) & "<br>"
Response.Write objCrt_Dic.Item("1") & "<br>"
Response.Write "varType : " & varType(arrKey(0)) & "<br>"

Response.Write Right("00" & hour(time()), 2) & "<br>"
Response.Write Year(date()) & Right("00" & month(date()), 2) & Right("00" & day(date()), 2)

Dim patternText : patternText = "[+-/*//]"
Dim content : content = "L1+L2-L3*L4/L5"
Dim regex, ResultReg
Set regex = New RegExp

regex.Pattern = patternText
regex.IgnoreCase = False
regex.Global = True

Set ResultReg = regex.Execute(content)
'Response.Write "<br><br>"
'Response.Write ResultReg.Count & "<br>"
if ResultReg.Count <> 0 then
    For Each Match In ResultReg
        resultString = Match.Value
        'Response.Write resultString & "<br>"
    Next
    
    'LBound(ResultReg)
    'Dim i
    'For i=LBound(ResultReg) To UBound(ResultReg)
    '    resultString = Match(i)
    'Next
end if
Response.Write "<br><br>"
Dim aTest : aTest = Array("L01+L03+L04+L05+0.297*L78+0.242*L79", 0, 0, 2.3, 5.6, 1.2, 2.3)
Dim aVal : aVal = Array("L23+L24+L25+L26+L27+0.188*L79", 0, 0, 3.3, 6.6, 2.7, 3.8)
For i = 0 To UBound(aTest)
    aTest(i) = aTest(i) +  aVal(i)
Next

For i = 0 To UBound(aTest)    
    'Response.Write aTest(i) & "<br>"
    'Response.Write i & "<br>"
Next

Dim ls_tmp, ls_el, i, larr_chr, li_s, ls_fe
Dim larr_sChar : larr_sChar = Array(Array("L", 3), Array("K",3), Array("H",3), Array("J",3))

ls_el = ""
ls_tmp = UCase("L01+L03+L04+L05+0.297*K78+0.242*L79L23+L24+J25+L26+H27+0.188*L79")

For i = 0 To UBound(larr_sChar)
    larr_chr = larr_sChar(i)
    Response.Write larr_chr(0)
    Response.Write larr_chr(1) & "<br>"
    li_s = InStr(ls_tmp, larr_chr(0))
    Response.Write li_s & "<br>"
    While li_s > 0
        ls_fe = Mid(ls_tmp, li_s, larr_chr(1))
        'Response.Write ls_fe & "<br>"
        ls_tmp = Replace(ls_tmp, ls_fe, "")
        'Response.Write ls_tmp & "<br>"
        ls_el = ls_el & ls_fe & ","
        'Response.Write ls_el & "<br>"
        li_s = InStr(ls_tmp, larr_chr(0))
    Wend
Next

Response.Write ls_el & "<br>"
IF ls_el <> "" Then 
   ls_el = Left(ls_el, len(ls_el) - 1)
   Response.Write ls_el & "<br>"
End If

ReDim paramInfo(2)
paramInfo(0) = DBHelper.MakeParam("@SLOT_IDX", adInteger , adParamInput,  , as_slotIdx)
paramInfo(1) = DBHelper.MakeParam("@TW_CODE" , adVarWChar, adParamInput, 5, as_twcode)
paramInfo(2) = DBHelper.MakeParam("@OPT"     , adVarWChar, adParamInput, 1, "S") 'TW_SEQ:고유번호'

Dim ls_sys, idx
ls_sys = "L048:1//////"
ls_sys = Split(ls_sys, "/")

Response.write LBound(ls_sys) & "-" & UBound(ls_sys) & "<br>"

For idx = 0 To UBound(ls_sys)
    Response.write "ls_sys=>" & ls_sys(idx) & "<br>"
Next

ls_tmp = (2*-11.9)--4.6+-12.0

execute("li_cal  = CDBL(FormatNumber(" & ls_tmp & " + 0, 5))")
Response.write  "li_cal=>" & li_cal & "<br>"

%>

<script>
    var foo = { key : 'value'};
    var bar = _.clone(foo);
    foo.key = 'other value';
    
    console.log(foo);
    console.log(bar);

</script>


</body>
</html>
