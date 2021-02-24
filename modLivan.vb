Option Strict On

Imports System.Text
Imports System.Security.Cryptography
Imports System.IO


Namespace livan

    Namespace extension

        Namespace fileInfo

            Public Module modFileInfo

                <System.Runtime.CompilerServices.Extension()> _
                Public Function antiBackSlashEnd(ByVal iString As String) As String
                    Do While Right(iString, 1) = "\"
                        iString = Left(iString, Len(iString) - 1)
                    Loop
                    Return iString
                End Function

                <System.Runtime.CompilerServices.Extension()> _
                Public Function namaExt(ByVal iSuatuPathFileName As String) As String
                    Dim aInfo() As String
                    aInfo = Strings.Split("\\\\\" & iSuatuPathFileName, "\")
                    Return aInfo(aInfo.Length - 1)
                End Function

            End Module


        End Namespace

    End Namespace

    Namespace enkripdekrip

        Public Module modGeneral

            Private key() As Byte = {}
            Private IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}

            <System.Runtime.CompilerServices.Extension()> _
            Public Function lakukanEnkrip(ByVal stringToEncrypt As String, ByVal SEncryptionKey As String) As String
                Dim aReturn As String = String.Empty
                Try
                    key = System.Text.Encoding.UTF8.GetBytes(Left(SEncryptionKey, 8))
                    Dim des As New DESCryptoServiceProvider()
                    Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(stringToEncrypt)
                    Dim ms As New MemoryStream()
                    Dim cs As New CryptoStream(ms, des.CreateEncryptor(key, IV), CryptoStreamMode.Write)

                    cs.Write(inputByteArray, 0, inputByteArray.Length)
                    cs.FlushFinalBlock()
                    aReturn = Convert.ToBase64String(ms.ToArray())

                Catch e As Exception
                    aReturn = e.Message
                End Try
                Return aReturn

            End Function

            <System.Runtime.CompilerServices.Extension()> _
            Public Function lakukanDekrip(ByVal stringToDecrypt As String, ByVal sEncryptionKey As String) As String
                If stringToDecrypt = "" Then
                    Return ""
                End If
                Dim inputByteArray(stringToDecrypt.Length) As Byte
                Try
                    key = System.Text.Encoding.UTF8.GetBytes(Left(sEncryptionKey, 8))
                    Dim des As New DESCryptoServiceProvider()
                    inputByteArray = Convert.FromBase64String(stringToDecrypt)
                    Dim ms As New MemoryStream()
                    Dim cs As New CryptoStream(ms, des.CreateDecryptor(key, IV), _
                        CryptoStreamMode.Write)
                    cs.Write(inputByteArray, 0, inputByteArray.Length)
                    cs.FlushFinalBlock()
                    Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
                    Return encoding.GetString(ms.ToArray())
                Catch e As Exception
                    Return e.Message
                End Try

            End Function

            <System.Runtime.CompilerServices.Extension()> _
            Public Function dekrip(ByVal iCypher As String, ByVal iKey As String) As String
                Dim aTemp As New uti_word.Encryption.Symmetric(uti_word.Encryption.Symmetric.Provider.TripleDES)
                Dim aHasil As String
                aHasil = aTemp.Decrypt(New uti_word.Encryption.Data(iCypher), New uti_word.Encryption.Data(iKey)).Text
                aTemp = Nothing
                Return aHasil
            End Function

            <System.Runtime.CompilerServices.Extension()> _
            Public Function enkrip(ByVal iCoba As String, ByVal iKey As String) As String
                Dim aTemp As New uti_word.Encryption.Symmetric(uti_word.Encryption.Symmetric.Provider.TripleDES)
                Dim aHasil As String
                aHasil = aTemp.Encrypt(New uti_word.Encryption.Data(iCoba), New uti_word.Encryption.Data(iKey)).Text
                aTemp = Nothing
                Return aHasil
            End Function

            <System.Runtime.CompilerServices.Extension()> _
            Public Function getMd5Hash(ByVal input As String) As String
                ' Create a new instance of the MD5 object.
                Dim md5Hasher As MD5 = MD5.Create()

                ' Convert the input string to a byte array and compute the hash.
                Dim data As Byte() = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input))

                ' Create a new Stringbuilder to collect the bytes
                ' and create a string.
                Dim sBuilder As New StringBuilder()

                ' Loop through each byte of the hashed data 
                ' and format each one as a hexadecimal string.
                Dim i As Integer
                For i = 0 To data.Length - 1
                    sBuilder.Append(data(i).ToString("x2"))
                Next i

                ' Return the hexadecimal string.
                Return sBuilder.ToString()

            End Function

            Public Function passwordHash(ByVal iString As String) As String
                Dim aRetval As String
                Dim aToHash As String
                aToHash = "ini" & iString & "untuk password"
                aRetval = getMd5Hash(aToHash)
                Return aRetval
            End Function
        End Module
    End Namespace


    Public Module modWord

        <System.Runtime.CompilerServices.Extension()> _
        Public Function potongKanan(ByVal iSumber As String, ByVal iJumlah As Integer) As String
            Return Strings.Left(iSumber, Len(iSumber) - iJumlah)
        End Function

        <System.Runtime.CompilerServices.Extension()> _
        Public Function potongKiri(ByVal iSumber As String, ByVal iJumlah As Integer) As String
            Dim aRetval As String

            Try
                aRetval = Strings.Right(iSumber, Len(iSumber) - iJumlah)
            Catch ex As Exception
                aRetval = ""
            End Try
            Return aRetval
        End Function

        Public Function divide(ByVal iKalimat As String, ByVal iDivider As String) As String()
            Dim aRetval(0 To 1) As String
            Dim aPosisi As Integer

            aPosisi = CInt(Strings.InStr(iKalimat, iDivider))
            aRetval(0) = Strings.Left(iKalimat, aPosisi - 1)
            aRetval(1) = Strings.Mid(iKalimat, aPosisi + Len(iDivider))
            Return aRetval
        End Function

        Public Function talkElite(ByRef strin As String) As String
            Dim inptxt As String
            Dim aLength As Integer
            Dim numspc As Integer
            Dim nextchr As String
            Dim nextchrr As String
            Dim aNewsent As String

            Dim aCrapp As Integer

            'Returns the string in elite
            inptxt = strin
            aLength = Len(inptxt)

            aNewsent = ""
            Do While numspc <= aLength

                numspc = numspc + 1
                nextchr = Mid$(inptxt, numspc, 1)
                nextchrr = Mid$(inptxt, numspc, 2)

                If nextchrr = "ae" Then
                    nextchrr = "æ"
                    aNewsent = aNewsent + nextchrr
                    aCrapp = 2
                    GoTo Greed
                ElseIf nextchrr = "AE" Then
                    nextchrr = "Æ"
                    aNewsent = aNewsent + nextchrr
                    aCrapp = 2
                    GoTo Greed
                ElseIf nextchrr = "oe" Then
                    nextchrr = "œ"
                    aNewsent = aNewsent + nextchrr
                    aCrapp = 2
                    GoTo Greed
                ElseIf nextchrr = "OE" Then
                    nextchrr = "Œ"
                    aNewsent = aNewsent + nextchrr
                    aCrapp = 2
                    GoTo Greed
                End If

                If aCrapp > 0 Then GoTo Greed

                Select Case nextchr
                    Case "A"
                        nextchr = "Å"
                    Case "a"
                        nextchr = "å"
                    Case "B"
                        nextchr = "ß"
                    Case "C"
                        nextchr = "Ç"
                    Case "c"
                        nextchr = "¢"
                    Case "D"
                        nextchr = "Ð"
                    Case "d"
                        nextchr = "ð"
                    Case "E"
                        nextchr = "Ê"
                    Case "e"
                        nextchr = "è"
                    Case "f"
                        nextchr = "ƒ"
                    Case "H"
                        nextchr = "h"
                    Case "I"
                        nextchr = "‡"
                    Case "i"
                        nextchr = "î"
                    Case "L"
                        nextchr = "£"
                    Case "M"
                        nextchr = "/\/\"
                    Case "m"
                        nextchr = "rn"
                    Case "N"
                        nextchr = "/\/"
                    Case "n"
                        nextchr = "ñ"
                    Case "O"
                        nextchr = "Ø"
                    Case "o"
                        nextchr = "ö"
                    Case "P"
                        nextchr = "¶"
                    Case "p"
                        nextchr = "Þ"
                    Case "r"
                        nextchr = "®"
                    Case "S"
                        nextchr = "§"
                    Case "s"
                        nextchr = "$"
                    Case "t"
                        nextchr = "†"
                    Case "U"
                        nextchr = "Ú"
                    Case "u"
                        nextchr = "µ"
                    Case "V"
                        nextchr = "\/"
                    Case "W"
                        nextchr = "VV"
                    Case "w"
                        nextchr = "vv"
                    Case "X"
                        nextchr = "><"
                    Case "x"
                        nextchr = "×"
                    Case "Y"
                        nextchr = "¥"
                    Case "y"
                        nextchr = "ý"
                    Case "!"
                        nextchr = "¡"
                    Case "?"
                        nextchr = "¿"
                    Case "."
                        nextchr = "…"
                    Case ","
                        nextchr = "‚"
                    Case "1"
                        nextchr = "¹"
                    Case "%"
                        nextchr = "‰"
                    Case "2"
                        nextchr = "²"
                    Case "3"
                        nextchr = "³"
                    Case "_"
                        nextchr = "¯"
                    Case "-"
                        nextchr = "—"
                    Case " "
                        nextchr = "  "
                    Case "<"
                        nextchr = "«"
                    Case ">"
                        nextchr = "»"
                    Case "*"
                        nextchr = "¤"
                    Case "`"
                        nextchr = """"
                    Case "'"
                        nextchr = """"
                    Case "0"
                        nextchr = "º"
                End Select
                aNewsent = aNewsent + nextchr

Greed:
                If aCrapp > 0 Then aCrapp = aCrapp - 1

            Loop
            Return aNewsent

        End Function

        Public Function talkInNumber(ByVal talkWhat As String) As String
            Dim banyakHuruf As Integer
            Dim theHuruf As String
            Dim aRetval As String = ""
            Dim tambah As String
            Dim aBcd As String = "IiEeAaSsZzGJjBbgOo1"
            Dim a123 As String = "113344552267788900L"
            Dim aPosisi As Integer

            banyakHuruf = Len(talkWhat)
            For inta As Integer = 1 To banyakHuruf
                theHuruf = Mid(talkWhat, inta, 1)
                aPosisi = Strings.InStr(aBcd, theHuruf)
                If aPosisi = 0 Then
                    tambah = theHuruf
                Else
                    tambah = Mid(a123, aPosisi, 1)
                End If
                aRetval = aRetval & tambah
            Next
            Return aRetval
        End Function

        Public Function decimalDariHexa(ByRef iSesuatu As String) As Double
            Dim aSample As String
            Dim aHuruf As String
            Dim aBank As Double

            aSample = "0123456789ABCDEF"
            aBank = 0
            For aInt As Integer = 1 To Len(iSesuatu)
                aHuruf = Mid$(iSesuatu, Len(iSesuatu) - aInt + 1, 1)
                aBank = aBank + (InStr(aSample, aHuruf) - 1) * 16 ^ (aInt - 1)
            Next aInt
            Return aBank
        End Function

        Public Function pembagianPembulatanKeatas(ByVal angka As Long, ByVal pembagi As Long) As Double
            Return CDbl(IIf(angka Mod pembagi <> 0, angka \ pembagi + 1, angka \ pembagi))
        End Function

        Public Function padded(ByVal sumber As String, ByVal kodeAsciiChar As Integer, ByVal jumlah As Integer, ByVal kanan As Boolean) As String
            Dim aRetval As String
            Dim aDelta As String

            aDelta = CStr(Strings.StrDup(jumlah, kodeAsciiChar))
            If kanan Then
                aRetval = Strings.Left(sumber & aDelta, jumlah)
            Else
                aRetval = Strings.Right(aDelta & sumber, jumlah)
            End If
            Return aRetval
        End Function

        ''' <summary>
        ''' langsung mengevaluate pembagian yang ditulis di string
        ''' </summary>
        ''' <param name="sesuatu"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function numberDariString(ByRef sesuatu As String) As Double
            Dim aRetval As Double
            Dim aInfo() As String

            Select Case True
                Case (InStr(sesuatu, "/") <> 0), (InStr(sesuatu, "\") <> 0)
                    aInfo = Strings.Split(sesuatu, "/")
                    aRetval = CDbl(aInfo(0)) / CDbl(aInfo(1))
                Case Else
                    aRetval = CDbl(sesuatu)
            End Select

            Return aRetval
        End Function

        Public Function stringDariNumber(ByVal angka As Single) As String
            Dim a As Integer
            Dim Saa As String

            Saa = Str$(angka)
            a = Len(Saa)
            If angka >= 0 Then
                stringDariNumber = Right$(Saa, a - 1)
            Else
                stringDariNumber = Saa
            End If
        End Function

        'Public Function teksDariTPDU(ByRef teks As String, ByVal panjangMessageSebenarnya As Integer) As String
        '    'Livan made this
        '    Dim jumHur As Integer
        '    Dim looper As Integer
        '    Dim asciiHurufSaatIni As Integer
        '    Dim asciiHurufSebelumnya As Integer
        '    Dim jumBitAmbil As Integer
        '    Dim depan As Integer
        '    Dim belakang As Integer
        '    Dim hasil As String

        '    jumHur = Len(teks)
        '    jumBitAmbil = 8

        '    For looper = 1 To jumHur Step 2
        '        asciiHurufSaatIni = CInt(decimalDariHexa(Mid(teks, looper, 2)))
        '        asciiHurufSebelumnya = CInt(decimalDariHexa(Mid(teks, looper - 2, 2)))

        '        If Err.Number <> 0 Then
        '            asciiHurufSebelumnya = 0
        '            Err.Number = 0
        '        End If

        '        jumBitAmbil = jumBitAmbil - 1
        '        If jumBitAmbil = 0 Then
        '            hasil = hasil & Chr(asciiHurufSebelumnya \ 2)
        '            jumBitAmbil = 7
        '        End If

        '        depan = CInt((asciiHurufSaatIni And CInt(2 ^ jumBitAmbil) - 1) * 2 ^ (7 - jumBitAmbil))
        '        belakang = (asciiHurufSebelumnya And (Not CInt(2 ^ jumBitAmbil - 1))) \ CInt(2 ^ (jumBitAmbil + 1))
        '        hasil = hasil & Chr(depan Or belakang)
        '    Next
        '    hasil = hasil & Chr(asciiHurufSaatIni \ 2)
        '    hasil = Left(hasil, panjangMessageSebenarnya)
        '    Return asciiFormatDariGsm(hasil)
        'End Function

        Public Function twocharhex(ByVal angkaByte As Integer) As String
            Dim temp As String
            temp = Hex$(angkaByte)
            If Len(temp) = 1 Then
                Return "0" & temp
            Else
                Return temp
            End If
        End Function

        Public Function getkatainthis(ByVal sumber As String, ByVal char1 As String, ByVal char2 As String) As String
            Dim sta As Integer, sto As Integer
            sta = InStr(sumber, char1)
            If sta = 0 Then
                Return ""
            End If
            sta = sta + Len(char1)
            sto = InStr(sta + 1, sumber, char2)
            If sto = 0 Then
                Return ""
            End If
            sto = sto - 1
            If sto = -1 Then sto = Len(sumber)
            Return kalimatrefpos(sumber, sta, sto)
        End Function

        Public Function kalimatrefpos(ByVal sumber As String, ByVal awal As Integer, ByVal akhir As Integer) As String
            On Error Resume Next
            If sumber = "" Then Return ""
            Return Strings.Mid(sumber, awal, akhir - awal + 1)
        End Function

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

        Public Function getInt(ByVal iObject As Object) As Integer
            Try
                Return CInt(iObject)
            Catch ex As Exception
                Return -1
            End Try
        End Function

        Public Function getDate(ByVal iObject As Object) As Date
            Try
                Return CDate(iObject)
            Catch ex As Exception
                Return New Date
            End Try
        End Function

        Public Function getDouble(ByVal iObject As Object) As Double
            Try
                Return CDbl(iObject)
            Catch ex As Exception
                Return -1
            End Try
        End Function
    End Module

End Namespace