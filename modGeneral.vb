Option Strict On

Public Module modGeneral

    'Public Function paddingBerKoma(iDouble As Double, iKarakterKoma As String, iJumlahAngkaSetelahKoma As Integer) As String
    '    Dim aString As String
    '    Dim aDummy() As String
    '    Dim aThousand As String
    '    Dim aDecim As String

    '    aThousand = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyGroupSeparator
    '    aDecim = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
    '    If iDouble = 0 Then
    '        aString = "0."
    '    Else
    '        aString = Format(iDouble, "###" & aThousand & "###" & aThousand & "###" & aDecim & Strings.StrDup(iJumlahAngkaSetelahKoma, "#")) & iKarakterKoma
    '    End If

    '    aDummy = Strings.Split(aString, iKarakterKoma)

    '    Return aDummy(0) & iKarakterKoma & Strings.Left(aDummy(1) & Strings.StrDup(iJumlahAngkaSetelahKoma, "0"), iJumlahAngkaSetelahKoma)

    'End Function

    Public Function addQueryStringParam(iStringAwal As String, iNamaParam As String, iValueParam As String) As String
        Dim aRetval As String
        Dim aDelta As String

        aDelta = iNamaParam & "=" & iValueParam
        If Strings.InStr(iStringAwal, "?", CompareMethod.Text) <> 0 Then
            aRetval = iStringAwal & "&" & aDelta
        Else
            aRetval = iStringAwal & "?" & aDelta
        End If
        Return aRetval
    End Function

    Public Function codeToHtml(ByVal iStringSumber As String) As String
        Dim VM As String = "ý"
        Dim SVM As String = "ü"
        Dim aRetval As String
        Dim aArray() As String = Split(iStringSumber, VM)
        Dim aInt As Integer
        Dim aParameter() As String = Nothing
        If aArray.Length > 1 Then
            aParameter = Split(aArray(1), SVM)
        End If

        aRetval = aArray(0)
        aRetval = Strings.Replace(aRetval, "%tanggal%", Format(CDate(Now()), "dd MMM yyyy"), Compare:=CompareMethod.Text)
        aRetval = Strings.Replace(aRetval, "%jam%", Format(CDate(Now()), "HH:mm:ss"), Compare:=CompareMethod.Text)

        If aParameter IsNot Nothing Then
            For aInt = 0 To aParameter.Length - 1
                aRetval = Strings.Replace(aRetval, "#par#" & CStr(aInt + 1) & "#", aParameter(aInt), Compare:=CompareMethod.Text)
            Next
        End If
        Return aRetval
    End Function

    Public Function dateFromYYYYmmDD(ByVal iTGLpoSAP As String) As Date
        Return New Date(CInt(Strings.Left(iTGLpoSAP, 4)), CInt(Strings.Mid(iTGLpoSAP, 5, 2)), CInt(Strings.Mid(iTGLpoSAP, 7, 2)))
    End Function

    Public Function ext(ByVal iValidPathFileName As String) As String
        Dim aPosisi As Integer
        aPosisi = InStrRev(iValidPathFileName, ".")
        Return Strings.Mid(iValidPathFileName, aPosisi + 1)
    End Function

    Public Function namaExt(ByVal iString As String) As String
        Dim aPosisiSlash As Integer
        Dim aPanjang As Integer
        aPosisiSlash = Math.Max(InStrRev(iString, "/"), InStrRev(iString, "\"))
        If aPosisiSlash = 0 Then
            Return iString
        Else
            aPanjang = iString.Length - aPosisiSlash
            Return iString.Substring(aPosisiSlash, aPanjang)
        End If
    End Function

    Public Function sandi(ByVal iKata As String) As String
        Dim aRetval As String
        Dim aTemp As String
        Dim aInt As Integer
        aRetval = ""
        For Each aChar As Char In iKata.ToCharArray
            aInt = aInt + 1

            aTemp = Hex(Asc(aChar) + aInt)
            If aRetval = "" Then
                aRetval = aTemp
            Else
                aRetval = aRetval & ";" & aTemp
            End If
        Next
        Return aRetval
    End Function

    Public Function unSandi(ByVal iKata As String) As String
        Dim aRetval As String
        Dim aInt As Integer

        If iKata = "" Then
            Return ""
        End If

        aRetval = ""
        For Each aString As String In Split(iKata, ";")
            aInt += 1
            aRetval = aRetval & Chr(CInt("&H" & aString) - aInt)
        Next
        Return aRetval
    End Function

    Public Function antiPadded(ByVal iSumber As String, ByVal iCharacter As String, ByVal iKiri As Boolean) As String
        Dim aRetval As String

        aRetval = ""
        If iKiri Then
            For aInt As Integer = 1 To Len(iSumber)
                If Mid(iSumber, aInt, 1) <> iCharacter Then
                    aRetval = Mid(iSumber, aInt)
                    Exit For
                End If
            Next
        Else
            For aInt As Integer = Len(iSumber) To 1 Step -1
                If Mid(iSumber, aInt, 1) <> iCharacter Then
                    aRetval = Left(iSumber, aInt)
                    Exit For
                End If
            Next
        End If

        Return aRetval
    End Function

    Public Function antiBackSlashEnd(ByVal iString As String) As String
        Do While Right(iString, 1) = "\"
            iString = Left(iString, Len(iString) - 1)
        Loop
        Return iString
    End Function

    Public Function normalizeDouble(ByVal iDouble As Double, ByVal iFormat As String, ByVal iJumlahAngkaBelakangKoma As Integer) As String
        Dim aRetval As String
        Dim aDecSeparator As String
        Dim aPosisiDecSep As Integer

        aDecSeparator = Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
        aRetval = Format(iDouble, iFormat)
        aPosisiDecSep = InStr(aRetval, aDecSeparator)
        If iDouble = 0 Then
            Return "0" & aDecSeparator & StrDup(iJumlahAngkaBelakangKoma, "0")
        End If
        If aPosisiDecSep = 0 Then
            aRetval = aRetval & aDecSeparator & StrDup(iJumlahAngkaBelakangKoma, "0")
        Else
            aRetval = aRetval & StrDup(iJumlahAngkaBelakangKoma - (Len(aRetval) - aPosisiDecSep), "0")
        End If
        Return aRetval
    End Function

    Public Function sumXOR(ByVal iInput As String, ByVal iBerapaKarakter As Integer) As String
        Dim aString As String
        Dim aSample As String
        Dim aHasil As String
        Dim aHuruf As String
        Dim aTempHasil As String
        aString = iInput
        aHasil = Strings.StrDup(iBerapaKarakter, Chr(0))

        Do
            Try
                aSample = Strings.Left(aString, iBerapaKarakter)
                aTempHasil = ""
                For aInt As Integer = 1 To iBerapaKarakter
                    aHuruf = Strings.Mid(aHasil, aInt, 1)
                    aTempHasil = aTempHasil & Chr(Asc(aHuruf) Xor Asc(Mid(aSample, aInt, 1)))
                Next
                aHasil = aTempHasil
                aString = Strings.Right(aString, Len(aString) - iBerapaKarakter)
            Catch ex As Exception
                Exit Do
            End Try
        Loop
        Return aHasil
    End Function

    Public Function HashPassword(ByVal password As String) As String
        Dim hashedPassword As String
        Dim hashProvider As New System.Security.Cryptography.SHA256Managed
        Try
            Dim passwordBytes() As Byte
            passwordBytes = System.Text.Encoding.Unicode.GetBytes(password)
            hashProvider.Initialize()
            passwordBytes = hashProvider.ComputeHash(passwordBytes)
            hashedPassword = System.Convert.ToBase64String(passwordBytes)
        Finally
            If Not hashProvider Is Nothing Then
                hashProvider.Clear()
                hashProvider = Nothing
            End If
        End Try
        Return hashedPassword

    End Function

    Public Function potongKanan(ByVal iSumber As String, ByVal iJumlah As Integer) As String
        Return Strings.Left(iSumber, Len(iSumber) - iJumlah)
    End Function

    Public Function potongKiri(iSumber As String, iJumlah As Integer) As String
        Return Strings.Right(iSumber, Len(iSumber) - iJumlah)
    End Function

    Public Function getKataInThis(iSumber As String, iPengapitAwal As String, iPengapitAkhir As String) As String
        Dim aPosAwal As Integer
        Dim aPosAkhir As Integer
        aPosAwal = Strings.InStr(1, iSumber, "<")
        aPosAkhir = Strings.InStr(aPosAwal, iSumber, ">")
        Return Strings.Mid(iSumber, aPosAwal, aPosAkhir - aPosAwal + 1)
    End Function

    '''' <summary>
    '''' ini hanya untuk keperluan mendebug saja.. mengubah mvitem supaya mudah dilihat mata
    '''' </summary>
    '''' <param name="iMvItem"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function stringDariMvitem(ByVal iMvItem As BlueFinity.mvNET.CoreObjects.mvItem) As String
    '    Dim aInt As Integer = 1
    '    Dim aLoop As Integer = 1
    '    Dim aTempstring As String = ""

    '    aInt = BlueFinity.mvNET.CoreObjects.DCount(iMvItem.Contents, BlueFinity.mvNET.CoreObjects.AM)
    '    For aLoop = 1 To aInt
    '        aTempstring &= CStr(aLoop).PadLeft(4, CChar("0")) & "==>" & CStr(iMvItem(aLoop)) & vbCrLf
    '    Next
    '    Return aTempstring
    'End Function

    Public Function isPalindrom(ByRef iString As String) As Boolean
        Return iString = Strings.StrReverse(iString)
    End Function

    Public Function breakDownMinutes(iMinutes As Integer) As clsMenitJamHari
        Dim aBreakDownMenit As Integer
        Dim aBreakDownJam As Integer
        Dim aTemp As Integer
        Dim aBreakDownHari As Integer
        Dim aRetval As New clsMenitJamHari

        aBreakDownMenit = iMinutes Mod 60
        aBreakDownJam = CInt((iMinutes - aBreakDownMenit) / 60)
        aTemp = aBreakDownJam Mod 24
        aBreakDownHari = CInt((aBreakDownJam - aTemp) / 24)
        aBreakDownJam = aTemp

        aRetval.jam = aBreakDownJam
        aRetval.menit = aBreakDownMenit
        aRetval.hari = aBreakDownHari

        Return aRetval
    End Function

End Module