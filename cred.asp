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
<%
    Dim objCredibility_Dic : Set objCredibility_Dic = Server.CreateObject("Scripting.Dictionary")
    
    Dim TW_CODE, TWCC_GBN_CD, TWCC_EQ, TWCC_SEQ, TWQ_SEQ1, TWQ_SEQ2
    Dim TWCC_POINT, TWQ_SORT_NUM1, TWQ_SORT_NUM2
    Dim TWQ_ANSWER_TYPE_CD1, TWQ_ANSWER_TYPE_CD2
    Dim TWQ_CORRECT_ANSWER1_1, TWQ_CORRECT_ANSWER2_1, TWQ_CORRECT_ANSWER3_1, TWQ_CORRECT_ANSWER4_1, TWQ_CORRECT_ANSWER5_1
    Dim TWQ_CORRECT_ANSWER1_2, TWQ_CORRECT_ANSWER2_2, TWQ_CORRECT_ANSWER3_2, TWQ_CORRECT_ANSWER4_2, TWQ_CORRECT_ANSWER5_2
    Dim TWCC_NEQ_POINT, TWCC_SYS ,li_mask, i

    Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
    str = "Provider=SQLOLEDB;User ID=uwizconv2;Password=upwd20#@3;Initial Catalog=OrpWiselection;Data Source=wfdbsvr.database.windows.net"
    conn.open str

    Set DBHelper = new clsDBHelper

    Dim objRs
    Dim paramInfo(4)
    Dim statusCode : statusCode = "00"
    Dim msg:msg = ""

    paramInfo(0)   = DBHelper.MakeParam("@P_IDX"    , adInteger, adParamInput,  , 284)
    paramInfo(1)   = DBHelper.MakeParam("@PN_IDX"   , adInteger, adParamInput,  , 548)
    paramInfo(2)   = DBHelper.MakeParam("@PS_IDX"   , adInteger, adParamInput,  , 2153)
    paramInfo(3)   = DBHelper.MakeParam("@PST_IDX"  , adInteger, adParamInput,  , 1781)
    paramInfo(4)   = DBHelper.MakeParam("@DATA_KIND", adVarChar, adParamInput, 3, "CRB")

    Set objRs = DBHelper.ExecSPReturnRS("usp_Web_eTest_Opt_ScoringInfo_List", paramInfo, Nothing)
    Dim li_point : li_point = 0
    While Not objRs.Eof
        TW_CODE                    = getParameter(objRs("TW_CODE")              ,"")
        TWCC_GBN_CD                = getParameter(objRs("TWCC_GBN_CD")          ,"")
        TWCC_EQ                    = getParameter(objRs("TWCC_EQ")              ,"")
        TWCC_SEQ                   = getParameter(objRs("TWCC_SEQ")             ,"0")
        
        TWQ_SEQ1                   = getParameter(objRs("TWQ_SEQ1")             ,"0")
        TWQ_SEQ2                   = getParameter(objRs("TWQ_SEQ2")             ,"0")
        
        TWCC_POINT                 = getParameter(objRs("TWCC_POINT")           ,"0")
        TWCC_NEQ_POINT             = getParameter(objRs("TWCC_NEQ_POINT")       ,"0")
        TWCC_SYS                   = getParameter(objRs("TWCC_SYS")             ,"0")
        
        TWQ_SORT_NUM1              = getParameter(objRs("TWQ_SORT_NUM1")        ,"0")
        TWQ_SORT_NUM2              = getParameter(objRs("TWQ_SORT_NUM2")        ,"0")
        
        TWQ_ANSWER_TYPE_CD1        = getParameter(objRs("TWQ_ANSWER_TYPE_CD1")  ,"")
        TWQ_ANSWER_TYPE_CD2        = getParameter(objRs("TWQ_ANSWER_TYPE_CD2")  ,"")
        
        TWQ_CORRECT_ANSWER1_1      = getParameter(objRs("TWQ_CORRECT_ANSWER1_1"),"")
        TWQ_CORRECT_ANSWER2_1      = getParameter(objRs("TWQ_CORRECT_ANSWER2_1"),"")
        TWQ_CORRECT_ANSWER3_1      = getParameter(objRs("TWQ_CORRECT_ANSWER3_1"),"")
        TWQ_CORRECT_ANSWER4_1      = getParameter(objRs("TWQ_CORRECT_ANSWER4_1"),"")
        TWQ_CORRECT_ANSWER5_1      = getParameter(objRs("TWQ_CORRECT_ANSWER5_1"),"")

        TWQ_CORRECT_ANSWER1_2      = getParameter(objRs("TWQ_CORRECT_ANSWER1_2"),"")
        TWQ_CORRECT_ANSWER2_2      = getParameter(objRs("TWQ_CORRECT_ANSWER2_2"),"")
        TWQ_CORRECT_ANSWER3_2      = getParameter(objRs("TWQ_CORRECT_ANSWER3_2"),"")
        TWQ_CORRECT_ANSWER4_2      = getParameter(objRs("TWQ_CORRECT_ANSWER4_2"),"")
        TWQ_CORRECT_ANSWER5_2      = getParameter(objRs("TWQ_CORRECT_ANSWER5_2"),"")               

        '신뢰도 정보
        li_mask = 0
        For i = 1 To 5
            IF Trim(objRs("TWQ_CORRECT_ANSWER" & i & "_1")) <> "" Then
                li_mask = li_mask + (2 ^ (i))
            End IF
        Next

        'objCredibility_Dic.Add TWCC_EQ & "_" & TWCC_SEQ, 
        '                        Array(Array(TWCC_GBN_CD, TWCC_EQ), _ '신뢰도구분코드, 신뢰도기준지수
        '                              Array(li_mask, TWCC_POINT, TWCC_NEQ_POINT, TWCC_SYS), _
        '                              Array(TWQ_ANSWER_TYPE_CD1, TWQ_ANSWER_TYPE_CD2), _
        '                              Array(TWQ_SEQ1, TWQ_SEQ2))  
        IF Trim(TWCC_EQ) = "DF" THEN
           Response.Write "=================================" & "<br>"
           IF li_mask AND CDBL(2 ^ 2) Then
               li_point = li_point + TWCC_POINT
               Response.Write "TWCC_POINT=>" & TWCC_POINT & " li_point=>" & li_point & "<br>"
           End IF

           Response.Write "TWQ_SEQ1=>[" & TWQ_SEQ1 & "] li_mask=>" & li_mask & "<br>"
        End IF

        objRs.MoveNext

    WEnd

    '================================================================================================
    ' Name : printDict
    ' Description : Dictionary data Print
    '================================================================================================
    Function printDict(ByRef adic_a, ByVal as_str)
        Dim ls_key

        For Each ls_key In adic_a
            Response.Write ls_key & ":" & adic_a.Item(ls_key) & as_str
        Next
    End Function

    Function getParameter(m, s)
        if m = "" or IsNull(m) then
        getParameter = Trim(s)
        else
        getParameter = Trim(m)
        end if
    End Function
%>
</body>
</html>                