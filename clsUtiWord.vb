Option Strict On

Public Module modDatabase
    'keperluan retrieve value dari database sql, karena dbnullvalue kalau diconvert to string, jadi exception

    Public Function getString(ByVal iObject As Object) As String
        Dim aRetval As String

        If DBNull.Value.Equals(iObject) Then
            aRetval = ""
        Else
            Try
                aRetval = CStr(iObject)
            Catch ex As Exception
                aRetval = ""
            End Try
        End If

        Return aRetval
    End Function

    Public Function getBoolean(ByVal iObject As Object) As Boolean
        Dim aRetval = False
        If DBNull.Value.Equals(iObject) Then
            aRetval = False
        Else
            aRetval = CBool(iObject)
        End If
        Return aRetval
    End Function

    Public Function getLong(ByVal iObject As Object) As Long
        Try
            Return CLng(iObject)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function getInt(ByVal iObject As Object) As Integer
        Return getInt(iObject, -1)
    End Function

    Public Function getInt(ByVal iObject As Object, iDefaultNumber As Integer) As Integer
        Try
            Return CInt(iObject)
        Catch ex As Exception
            Return iDefaultNumber
        End Try
    End Function

    Public Function myMsg(ByVal iString As String(), Optional ByVal iStyle As MsgBoxStyle = MsgBoxStyle.OkOnly Or MsgBoxStyle.Information) As MsgBoxResult
        Dim aString = Strings.Join(iString)
        Dim aHasil = MsgBox(aString, iStyle)
        Return aHasil
    End Function

    Public Function getsingle(iObject As Object) As Single
        Try
            Return CSng(iObject)
        Catch ex As Exception
            Return -1.0!
        End Try
    End Function

    Public Function getDate(ByVal iObject As Object) As Date
        If DBNull.Value.Equals(iObject) Then
            Return Nothing
        ElseIf "".Equals(iObject) Then
            Return Nothing
        Else
            Return CDate(iObject)
        End If
    End Function

    Public Function getDateTime(ByVal iObject As Object) As DateTime
        If DBNull.Value.Equals(iObject) Then
            Return Nothing
        ElseIf "".Equals(iObject) Then
            Return Nothing
        Else
            Return CDate(iObject)
        End If
    End Function

    Public Function getDouble(iObject As Object, iDefaultValue As Double) As Double
        Dim aRetval As Double

        If DBNull.Value.Equals(iObject) OrElse "".Equals(iObject) Then
            aRetval = iDefaultValue
        Else
            Try
                aRetval = CDbl(iObject)
            Catch ex As Exception
                aRetval = iDefaultValue
            End Try
        End If

        Return aRetval
    End Function

    Public Function getDouble(ByVal iObject As Object) As Double
        Return getDouble(iObject, -1)
    End Function

    Public Function stringDariDate(ByVal iDate As Object, Optional ByVal iFormat As String = "dd MMM yyyy") As String
        Dim aNewDate07 As New DateTime
        Dim aNewDate As New Date
        aNewDate07 = aNewDate07.AddHours(7)

        If DBNull.Value.Equals(iDate) OrElse
            (aNewDate).Equals(iDate) OrElse
            (aNewDate07).Equals(iDate) Then

            Return ""
        Else
            Return Strings.Format(iDate, iFormat)
        End If
    End Function

End Module