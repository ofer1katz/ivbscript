Option Explicit
On Error Resume Next
Dim oInterpreter

Class Interpreter
    Private ForReading, ForWriting, ForAppending
    Private invokeKindPropertyGet, invokeKindFunction, invokeKindPropertyPut, invokeKindPropertyPutRef
    Private cmdFilePath, retFilePath
    Private cmdFile, retFile, logFile, debugPath
    Private fso, WshShell, typeDetails

    Private Sub Class_Initialize()
        ForReading = 1
        ForWriting = 2
        ForAppending = 8
        invokeKindPropertyGet = 0
        invokeKindFunction = 1
        invokeKindPropertyPut = 2
        invokeKindPropertyPutRef = 4
        Set WshShell = CreateObject("WScript.Shell")
        debugPath = WshShell.ExpandEnvironmentStrings("%IVBS_DEBUG_PATH%")
        cmdFilePath = WshShell.ExpandEnvironmentStrings("%IVBS_CMD_PATH%")
        retFilePath = WshShell.ExpandEnvironmentStrings("%IVBS_RET_PATH%")
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set logFile = fso.OpenTextFile(debugPath, ForWriting, True)
        Set typeDetails = CreateObject("Scripting.Dictionary")
        ' Add all values.
        typeDetails.add vbEmpty, "vbEmpty (uninitialized variable)" ' ; =0
        typeDetails.add vbNull, "vbNull (value unknown)" ' ; =1
        typeDetails.add vbInteger, "vbInteger" ' Short? ; =2
        typeDetails.add vbLong, "vbLong" ' Integer? ; =3
        typeDetails.add vbSingle, "vbSingle" ' ; =4
        typeDetails.add vbDouble, "vbDouble" ' ; =5
        typeDetails.add vbCurrency, "vbCurrency" ' ; =6
        typeDetails.add vbDate, "vbDate" ' ; =7
        typeDetails.add vbString, "vbString" ' ; =8
        typeDetails.add vbObject, "vbObject" ' ; =9
        typeDetails.add 10, "Exception" ' ; =10
        typeDetails.add vbBoolean, "vbBoolean" ' ; =11
        typeDetails.add vbVariant, "vbVariant" ' ; =12
        typeDetails.add 13, "DataObject" ' ; =13
        typeDetails.add vbDecimal, "vbDecimal" ' ; =14
        typeDetails.add vbByte, "vbByte" ' ; =17
        typeDetails.add 18, "vbChar" ' ; =18
        typeDetails.add 19, "ULong" ' ; =19
        typeDetails.add 20, "Long" ' really Long? ; =20
        typeDetails.add 24, "(void)" ' ; =24
        typeDetails.add 36, "UserDefinedType" ' ; =36

        logFile.WriteLine cmdFilePath
        logFile.WriteLine retFilePath
    End Sub

    Private Sub Class_Terminate()
        Set logFile = Nothing
        Set fso = Nothing
        Set WshShell = Nothing
        Set typeDetails = Nothing
    End Sub

    Public Sub Run()
        On Error Resume Next
        Dim cmd, response, stderr, inspect
        Do
            stderr = ""
            While fso.FileExists(cmdFilePath) = False
                WScript.Sleep(500)
            Wend
            Set cmdFile = fso.OpenTextFile(cmdFilePath, ForReading, False)
            cmd = cmdFile.ReadAll()
            cmdFile.Close()
            Set cmdFile = Nothing
            fso.DeleteFile cmdFilePath, True
            inspect = False
            cmd = RTrim(cmd)
            logFile.WriteLine cmd
            logFile.WriteLine Mid(cmd, Len(cmd), 1)
            If Mid(cmd, Len(cmd), 1) = "?" Then
                inspect = True
                cmd = "oInterpreter.HandleInspect " & Mid(cmd, 1, Len(cmd) - 1)
                logFile.WriteLine cmd
            End If
            Err.Clear()
            ExecuteGlobal cmd
            If Err.Number <> 0 Then
                stderr = "Err.Description: " & Err.Description & "." & vbNewLine & "Err.Number: " & Err.Number
                Err.Clear()
            End If
            logFile.WriteLine stderr
            Set retFile = fso.OpenTextFile(retFilePath, ForWriting, True)
            retFile.Write(stderr)
            retFile.Close()
            Set retFile = Nothing
        Loop
    End Sub

    Private Function GetVarTypeName(var)
        If typeDetails.Exists(var) Then
            GetVarTypeName = typeDetails(var)
        ' Why 8192?
        ElseIf var > 8192 Then
            GetVarTypeName = "vbArray"
        Else
            GetVarTypeName = "Unknown Type " & var
        End If
    End Function

    Public Sub HandleInspect(cmd)
        Dim result
        result = GetObjectInfo(cmd)
        logFile.WriteLine result
        WScript.Echo result
    End Sub

    Private Function GetObjectInfo(object)
        Dim typeLibInfo, typeInfo
        Dim member, memberInfo
        Dim parameterList, parameter, i, length
        Dim result: result = ""
        ' Object is empty.
        If IsEmpty(object) Or IsNull(object) Then
            result = GetVarTypeName(VarType(object))
        ' Object is invalid.
        ElseIf TypeName(object) = "Nothing" Then
            result = "Nothing (The Invalid Object)"
        ' Object if from regular type.
        ElseIf Not IsObject(object) And Not IsArray(object) Then
            result = GetVarTypeName(VarType(object)) & ", Value: " & object
        ' Object is Array.
        ElseIf IsArray(object) Then
            length = UBound(object)
            result = GetVarTypeName(VarType(object)) & " Len: " & length + 1 & vbNewLine
            For i = 0 To length
                result = result & "  " & i & "): " & GetObjectInfo(object(i)) & vbNewLine
            Next
        ' Object is unknown type (using reflection).
        Else
            Set typeLibInfo = CreateObject("TLI.TLIApplication")
            result = "Object " & TypeName(object)
            ' Get type information of the object.
            Err.Clear()
            Set typeInfo = typeLibInfo.InterfaceInfoFromObject(object)
            If Err.Number <> 0 Then

                result = result & "; Error: Failed to get type information of the object - Stopping! " & _
                "Err.Description: " & Err.Description & "." & vbNewLine & "Err.Number: " & Err.Number
                Err.Clear()
                Exit Function
            End If
            ' Get members of the object.
            For Each member In typeInfo.Members
                memberInfo = ""
                ' Build member information by its type.
                Select Case member.InvokeKind
                    ' Member is a Function/Sub.
                    Case InvokeKindFunction
                        ' Get function's type.
                        If member.ReturnType.VarType <> 24 Then
                            memberInfo = " Function " & GetVarTypeName(member.ReturnType.VarType)
                        Else
                            memberInfo = " Sub"
                        End If
                        ' Get function's signature.
                        memberInfo = memberInfo & " " & member.Name
                        parameterList = Array()
                        For Each parameter In member.Parameters
                            ReDim Preserve parameterList(UBound(parameterList) + 1)
                            parameterList(UBound(parameterList)) = parameter.Name
                        Next
                        memberInfo = memberInfo & "(" & Join(parameterList, ", ") & ")"
                    ' Member is a property.
                    Case InvokeKindPropertyGet
                        memberInfo = " Property " & member.Name
                    ' Member is a get/set property.
                    Case InvokeKindPropertyPut
                        memberInfo = " Property (set/get) " & member.Name
                    ' Member is a ref/set property.
                    Case InvokeKindPropertyPutRef
                        memberInfo = " Property (set ref/get) " & member.Name
                    ' Member is from unknown type.
                    Case Else
                        memberInfo = " Unknown member, InvokeKind " & member.InvokeKind
                End Select
                result = result & vbNewLine & memberInfo
            Next
            Set typeInfo = Nothing
            Set typeLibInfo = Nothing
        End If
        GetObjectInfo = result
    End Function
End Class

Set oInterpreter = New Interpreter
oInterpreter.Run()

Set oInterpreter = Nothing
WScript.Quit
