<!--METADATA NAME="Microsoft ActiveX Data Objects 2.5 Library" TYPE="TypeLib" UUID="{00000205-0000-0010-8000-00AA006D2EA4}"-->
<%
'====================================================================================================================================================
'프로젝트   : wiselecationLibrary
'파일명     : /global/DBHeper.inc.asp
'작성자		: 이명수(MyungSu.Lee@gmail.com)
'기  능		: DB Helper 설정
'수정일		:
'설정내역	:
'====================================================================================================================================================

Dim cmd
Dim rs

Dim SmsConnString
Dim ProcConnString

Class clsDBHelper
    Private DefaultConnString
    Private DefaultConnection

    private sub Class_Initialize()

        'Dim objShell
        'Set objShell = CreateObject("WScript.Shell")

        'Dim connStr : connStr = "MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZgMC AgDABAgFujOCkjztnwQQ64SvzO8108wjRPiKaByitASBqNu4vavJg14Ds6qe63ny VUVTL9s+8TPrg93BW/r4OLN12XXKtsbawYjSMdag9jmmUwU7rDdXGyhBtpbYc9DJ pf7H3PrxVVSbcG8pIppp0xTdSSowzcp455kGr0r8TvY6xLuiQ6jC+RIO2RF2aL20 5tRBpZ2n0oMoSWP+3Sz0qigqUESWKVhqOtqRC4FkWK5wcveIlFyKNiRFzp5RkswO ybnnE08LpylI2w=="

        'If (fn_IsWorkName() <> "real") Then
          'connStr = "MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZgMC AgDABAj9H7Xqp1V1tQQQy73p6xnkGTNQbFtRiWw95QSBqC0mpaY3qMHnRrKOh+za 9oy5JUB+2vUNBa+cInXaJ6oAvv3zmLiidBhqKvrq8Uk14spINl0MlMnyk12ziOYq kUrmJQu9G1/Q6xj+7BqMxb4teZI2ZxIiGPuqNm2kUi60X8ME1pr0MxMiFPYcpiYb JSQ455ow9ImF10Pg+Xmi6R8WVkUgQfMNg5wAsyFZLqHTxcNDAlLMB5+bWameJ5vT E8TaQ3Or4VmhwA=="
        'End IF

        'Dim strDefaultConn : strDefaultConn = objShell.RegRead(fnDecrypt(connStr))
        DefaultConnString = "Provider=SQLOLEDB;User ID=uwizconv2;Password=upwd20#@3;Initial Catalog=OrpWiselection;Data Source=wfdbsvr.database.windows.net"

        Set DefaultConnection = Nothing

        '//Write Proc DB connect '
        'Dim procConnStr
        'If (fn_IsWorkName() <> "real") Then
            '//테스트 사용'
        '    procConnStr="MIH3BgkrBgEEAYI3WAOggekwgeYGCisGAQQBgjdYAwGggdcwgdQCAwIAAQICZgMC AgDABAhTocw5vCTdMAQQGs7V0of8NDS7RcfDM5vkagSBqNx3s1V9isS5ADUyQ9rh +Uyhz28xuLsXtXNnEC4RuGKfrgHHIZTgVpGhCCwhLVBUAOFUCyRSO5bir5p000n6 2FU8zoLNyEmIrNSJQB+xCJXaeyPy2KqxAdKW36q/eRTVPkwev2EXeFRvK1MlrLuX /k2+IaFJJ/fquWKwDuDZ6ZpJKXZ+aDIuncYEa+9fkbogOVtwujJZKKJOMsHN//ub Obm7r6T1obKs/w=="
        'Else
            '//운영 사용'
        '    procConnStr="MIH/BgkrBgEEAYI3WAOggfEwge4GCisGAQQBgjdYAwGggd8wgdwCAwIAAQICZgMC AgDABAj6YBWWgVyaeQQQTAEMOKGHEAYEem1xPHiFcASBsPz+MDN47BHRcW/YRA49 P4jtNI/wc0uub53DzDv0JwS3P09OXE4OkY0E7Sk0pPcPU5pxtkACDKRu+HDm1wPq PIR0Sph5bZCpa7+9v1qj8F5kxmu1DWxJjcJZVC7vbdFmCRnwkqpPNCtUHStjNJ5C Ii8FXcn301AI0XfCRSzyQFPWNQj7qNw/+RIlFYXkmQ14xCYpBZlYtxd1oLCDF271 He1bsZ9cgAZwbVSWULNmNVP/"
        'End IF

        'Dim strProcConn:strProcConn=objShell.RegRead(fnDecrypt(procConnStr))
        'ProcConnString=fnDecrypt(strProcConn)

        'Dim smsConnStr : smsConnStr = "MIHfBgkrBgEEAYI3WAOggdEwgc4GCisGAQQBgjdYAwGggb8wgbwCAwIAAQICZgMC AgDABAhXSNDXrGJf4wQQviD6a5QThoR4qXRiO9CQvQSBkMjeyJ8Mpyy74dziiqVk +p8K3E1ts+Pw2/aCqJL60y/JSyzQ8oiYF4PzKq6bJSGpxNpr4Vw/jY1WtIIQL6Kf AHQfBLb0DOoSN7KtNtGJlWYBzAlxvDzoWTl6kce+zF2o38lf9yX1fd4fFGQy3kWl UoH6VOr2EzNE4zC1+pvFkH+0vA1Yo3/GIvvCb72dglTyeA=="
        'Dim strSmsConn : strSmsConn = objShell.RegRead(fnDecrypt(smsConnStr))
        'SmsConnString = fnDecrypt(strSmsConn)

        Set objShell = Nothing
    End Sub

'---------------------------------------------------
' SP를 실행하고, RecordSet을 반환한다.
'---------------------------------------------------
Public Function ExecSPReturnRS(spName, params, connectionString)
  Dim i2

  If IsObject(connectionString) Then
    If connectionString is Nothing Then
      If DefaultConnection is Nothing Then
        Set DefaultConnection = CreateObject("ADODB.Connection")
        DefaultConnection.Open DefaultConnString
      End If
      Set connectionString = DefaultConnection
    End If
  End If

  Set rs = CreateObject("ADODB.RecordSet")
  Set cmd = CreateObject("ADODB.Command")


  cmd.ActiveConnection = connectionString
  cmd.CommandText = spName
  cmd.CommandType = adCmdStoredProc
  cmd.CommandTimeout =300 'for query continuee by whpark MAX 5MIN'
  Set cmd = collectParams(cmd, params)
  'cmd.Parameters.Refresh

  rs.CursorLocation = adUseClient
  rs.Open cmd, ,adOpenStatic, adLockReadOnly

  For i2 = 0 To cmd.Parameters.Count - 1
    If cmd.Parameters(i2).Direction = adParamOutput OR cmd.Parameters(i2).Direction = adParamInputOutput OR cmd.Parameters(i2).Direction = adParamReturnValue Then
      If IsObject(params) Then
        If params is Nothing Then
          Exit For
        End If
      Else
        params(i2)(4) = cmd.Parameters(i2).Value

        'Response.write  params(i2)(4) &"__"& cmd.Parameters(i2).Value & "<br/>"

      End If
    End If
  Next

  Set cmd.ActiveConnection = Nothing
  Set cmd = Nothing
  'Set rs.ActiveConnection = Nothing

' Response.write  rs.ActiveConnection & "<br/>"

    if Not rs.ActiveConnection is Nothing then Set rs.ActiveConnection = Nothing

  Set ExecSPReturnRS = rs
End Function

'---------------------------------------------------
' SQL Query를 실행하고, RecordSet을 반환한다.
'---------------------------------------------------
Public Function ExecSQLReturnRS(strSQL, params, connectionString)
  If IsObject(connectionString) Then
    If connectionString is Nothing Then
      If DefaultConnection is Nothing Then
        Set DefaultConnection = CreateObject("ADODB.Connection")
        DefaultConnection.Open DefaultConnString
      End If
      Set connectionString = DefaultConnection
    End If
  End If

    Set rs = CreateObject("ADODB.RecordSet")
    Set cmd = CreateObject("ADODB.Command")


    cmd.ActiveConnection = connectionString
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    'cmd.CommandTimeout = 0 'for query continuee by whpark'
    Set cmd = collectParams(cmd, params)

    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing

    Set ExecSQLReturnRS = rs
End Function

'---------------------------------------------------
' SP를 실행한다.(RecordSet 반환없음)
'---------------------------------------------------
Public Sub ExecSP(strSP,params,connectionString)

  If IsObject(connectionString) Then
    If connectionString is Nothing Then
      If DefaultConnection is Nothing Then
        Set DefaultConnection = CreateObject("ADODB.Connection")
        DefaultConnection.Open DefaultConnString
      End If
      Set connectionString = DefaultConnection
    End If
  End If

    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = connectionString
  cmd.CommandText = strSP
  cmd.CommandType = adCmdStoredProc
    Set cmd = collectParams(cmd, params)

    cmd.Execute , , adExecuteNoRecords
      Dim i2
    For i2 = 0 To cmd.Parameters.Count - 1
      If cmd.Parameters(i2).Direction = adParamOutput OR cmd.Parameters(i2).Direction = adParamInputOutput OR cmd.Parameters(i2).Direction = adParamReturnValue Then
        If IsObject(params) Then
          If params is Nothing Then
            Exit For
          End If
        Else
          params(i2)(4) = cmd.Parameters(i2).Value
        End If
      End If
    Next

    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
End Sub

'---------------------------------------------------
' SP를 실행한다.(RecordSet 반환없음)
'---------------------------------------------------
Public Sub ExecSQL(strSQL,params,connectionString)
  If IsObject(connectionString) Then
    If connectionString is Nothing Then
      If DefaultConnection is Nothing Then
        Set DefaultConnection = CreateObject("ADODB.Connection")
        DefaultConnection.Open DefaultConnString
      End If
      Set connectionString = DefaultConnection
    End If
  End If

    Set cmd = CreateObject("ADODB.Command")

    cmd.ActiveConnection = connectionString
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    Set cmd = collectParams(cmd, params)

    cmd.Execute , , adExecuteNoRecords

    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
End Sub

'---------------------------------------------------
' 트랜잭션을 시작하고, Connetion 개체를 반환한다.
'---------------------------------------------------
Public Function BeginTrans(connectionString)
  If IsObject(connectionString) Then
    If connectionString is Nothing Then
      connectionString = DefaultConnString
    End If
  End If

  Dim conn : Set conn = Server.CreateObject("ADODB.Connection")
  conn.Open connectionString
  conn.BeginTrans
  Set BeginTrans = conn
End Function

'---------------------------------------------------
' 활성화된 트랜잭션을 커밋한다.
'---------------------------------------------------
Public Sub CommitTrans(connectionObj)
  If Not connectionObj Is Nothing Then
    connectionObj.CommitTrans
    connectionObj.Close
    Set ConnectionObj = Nothing
  End If
End Sub

'---------------------------------------------------
' 활성화된 트랜잭션을 롤백한다.
'---------------------------------------------------
Public Sub RollbackTrans(connectionObj)
  If Not connectionObj Is Nothing Then
    connectionObj.RollbackTrans
    connectionObj.Close
    Set ConnectionObj = Nothing
  End If
End Sub

'---------------------------------------------------
' 배열로 매개변수를 만든다.
'---------------------------------------------------
Public Function MakeParam(PName,PType,PDirection,PSize,PValue)
  MakeParam = Array(PName, PType, PDirection, PSize, PValue)
End Function

'---------------------------------------------------
' 매개변수 배열 내에서 지정된 이름의 매개변수 값을 반환한다.
'---------------------------------------------------
Public Function GetValue(params, paramName)
Dim param
  For Each param in params
    If param(0) = paramName Then
      GetValue = param(4)
      Exit Function
    End If
  Next
End Function

Public Sub Dispose
    if (Not DefaultConnection is Nothing) Then
        if (DefaultConnection.State = adStateOpen) Then DefaultConnection.Close
        Set DefaultConnection = Nothing
    End if
End Sub

'---------------------------------------------------------------------------
'Array로 넘겨오는 파라메터를 Parsing 하여 Parameter 객체를
'생성하여 Command 객체에 추가한다.
'---------------------------------------------------------------------------
Private Function collectParams(cmd,argparams)
    Dim params
    Dim i2
    Dim l
    Dim u
    Dim v
    If VarType(argparams) = 8192 or VarType(argparams) = 8204 or VarType(argparams) = 8209 then
        params = argparams
        For i2 = LBound(params) To UBound(params)
            l = LBound(params(i2))
            u = UBound(params(i2))

            ' Check for nulls.
            If u - l = 4 Then
                'Response.Write i2 & " : " & VarType(params(i2)(4)) & "_" & params(i2)(0) & "_" & params(i2)(1) & "_" & params(i2)(2) & "_" & v & chr(13)	'	디버깅용 -
                If VarType(params(i2)(4)) = vbString Then
                    If params(i2)(4) = "" Then
                        v = ""
                    Else
                        v = params(i2)(4)
                    End If
                Else
                    v = params(i2)(4)
                End If
                cmd.Parameters.Append cmd.CreateParameter(params(i2)(0), params(i2)(1), params(i2)(2), params(i2)(3), v)
                IF params(i2)(1) = adDecimal Then'wohho.park
                      cmd.Parameters(params(i2)(0)).NumericScale=30
                      cmd.Parameters(params(i2)(0)).Precision=38
                End IF
            End If
        Next

        Set collectParams = cmd
        Exit Function
    Else
        Set collectParams = cmd
    End If
End Function

End Class
%>
