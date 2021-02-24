Option Strict On

Public Module clsMathInString

    Private Const _akurasi As Integer = 15
    Private Const _siKoma As String = "."
    Private _arrayHasil(0, 0) As Integer
    Private _varSisa As Integer

    Private Function zGetSisa() As Integer
        Dim aTemp As Integer
        aTemp = _varSisa
        _varSisa = 0
        Return aTemp
    End Function

    Private Sub zPutSisa(ByVal berapa As Integer)
        _varSisa = berapa
    End Sub

    Private Function zTanda(ByRef angkaString1 As String, angkaString2 As String) As String
        Dim retVal As String
        retVal = "-"

        If (InStr(angkaString1, "-") <> 0 And InStr(angkaString2, "-") <> 0) OrElse
           (InStr(angkaString1, "-") = 0 And InStr(angkaString2, "-") = 0) Then
            retVal = ""
        End If
        Return retVal
    End Function

    ''' <summary>
    ''' ini support perkalian integer yang tak terbatas.. 17291851986982691586668256192568215 * 5687659286592685928561925681925689
    ''' </summary>
    ''' <param name="string1"></param>
    ''' <param name="string2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function perkalianStringInteger(ByVal string1 As String, ByVal string2 As String) As String
        Dim indexAtas As Integer
        Dim indexBawah As Integer
        Dim angkaBawah As Integer
        Dim angkaAtas As Integer
        Dim hasilPerkalian As Integer
        Dim hasilPenambahan As Integer
        Dim satuan As Integer
        Dim len1 As Integer
        Dim len2 As Integer
        Dim kolomCell As Integer
        Dim barisCell As Integer
        Dim aTempString As String
        Dim tandaAkhir As String
        Dim aPuluhan As Integer

        aTempString = ""
        tandaAkhir = zTanda(string1, string2)

        string1 = Replace(Replace(string1, _siKoma, ""), "-", "")
        string2 = Replace(Replace(string2, _siKoma, ""), "-", "")

        len1 = Len(string1)
        len2 = Len(string2)

        ReDim _arrayHasil(len2, len1 + len2 + 1)

        For indexBawah = len2 To 1 Step -1
            angkaBawah = CInt(Mid(string2, indexBawah, 1))
            kolomCell = len1 + indexBawah
            For indexAtas = len1 To 1 Step -1
                angkaAtas = CInt(Mid(string1, indexAtas, 1))
                hasilPerkalian = zGetSisa() + angkaAtas * angkaBawah
                satuan = hasilPerkalian Mod 10
                aPuluhan = hasilPerkalian \ 10
                zPutSisa(aPuluhan)
                _arrayHasil(indexBawah, kolomCell) = satuan
                kolomCell = kolomCell - 1
            Next
            If _varSisa <> 0 Then
                _arrayHasil(indexBawah, kolomCell) = zGetSisa()
            End If
        Next

        For kolomCell = len1 + len2 To 1 Step -1
            hasilPenambahan = zGetSisa()
            For barisCell = len2 To 1 Step -1
                hasilPenambahan = hasilPenambahan + _arrayHasil(barisCell, kolomCell)
            Next
            satuan = hasilPenambahan Mod 10
            aPuluhan = hasilPenambahan \ 10
            zPutSisa(aPuluhan)
            aTempString = satuan & aTempString
        Next

        Dim inta As Integer
        For inta = 1 To Len(aTempString)
            If Mid(aTempString, inta, 1) <> "0" Then
                aTempString = Mid(aTempString, inta)
                Exit For
            End If
        Next

        Return tandaAkhir & aTempString
    End Function

    Public Function penambahanStringInteger(iInteger1 As String, iInteger2 As String) As String

        Dim aJumlahMax As Integer
        aJumlahMax = Math.Max(Len(iInteger1), Len(iInteger2))

        iInteger1 = iInteger1.PadLeft(aJumlahMax, "0"c)
        iInteger2 = iInteger2.PadLeft(aJumlahMax, "0"c)

        Dim aAtas As Integer
        Dim aBawah As Integer
        Dim aTempor As Integer
        Dim aHasil As String
        Dim aSisa As Integer
        aHasil = ""
        For aInt As Integer = aJumlahMax To 1 Step -1
            aAtas = CInt(Strings.Mid(iInteger1, aInt, 1))
            aBawah = CInt(Strings.Mid(iInteger2, aInt, 1))
            aTempor = (aAtas + aBawah) + aSisa
            aHasil = CStr(aTempor Mod 10) & aHasil
            aSisa = aTempor \ 10
        Next
        If aSisa = 0 Then
            Return aHasil
        Else
            Return aSisa & aHasil
        End If
    End Function

    Public Function xPangkatYinteger(iEx As String, iYe As Integer) As String
        Dim aHasil As String
        aHasil = "1"

        For aInt As Integer = 1 To iYe
            aHasil = lib_word.perkalianStringInteger(aHasil, iEx)
        Next

        Return aHasil
    End Function

End Module