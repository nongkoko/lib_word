Option Strict On

Public Class terbilang
    Private _decimalSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
    Private _angka As String() = {"", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas"}
    Private _level As String() = {"", "ribu", "juta", "miliar"}

    Public Function terbilang(ByVal iAngka As Double) As String
        Dim aRetval As String
        Dim aAngka As String
        Dim aKoma As String
        Dim aSplit() As String
        Dim aTotalLength As Integer

        aKoma = ""

        'sample 12 345 678
        aSplit = Split(CStr(iAngka), _decimalSeparator)
        aAngka = aSplit(0)
        If aSplit.Length > 1 Then
            aKoma = aSplit(1)
        End If

        aTotalLength = Len(aAngka)
        Dim aSpent As Integer
        Dim aLevel As Integer
        Dim aHasil As String
        aRetval = ""
        Do While aTotalLength > 0
            aSpent = Math.Min(aTotalLength, 3)
            aHasil = spell3char(Mid(aAngka, Math.Max(aTotalLength - 3 + 1, 1), aSpent))
            If aHasil <> "" Then
                If _level(aLevel) <> "" Then
                    aHasil = aHasil & " " & _level(aLevel)
                End If
                aHasil = Replace(aHasil, "satu ribu", "seribu")
                If aRetval = "" Then
                    aRetval = aHasil
                Else
                    aRetval = aHasil & " " & aRetval
                End If
            End If

            aTotalLength = aTotalLength - aSpent
            aLevel += 1
        Loop
        If aRetval = "" Then
            aRetval = "nol"
        End If
        If aKoma <> "" Then
            aRetval = aRetval & " koma " & spellKoma(aKoma)
        End If
        Return aRetval
    End Function

    Public Function spell3char(ByVal iAngka As String) As String
        Dim aRetval As String
        Dim aCharKedua As String
        Dim aCharPertama As String
        Dim aCharKetiga As String

        iAngka = Right("000" & iAngka, 3)
        aCharPertama = Mid(iAngka, 1, 1)
        aCharKedua = Mid(iAngka, 2, 1)
        aCharKetiga = Mid(iAngka, 3, 1)
        aRetval = ""

        If aCharPertama = "1" Then
            aRetval = "seratus "
        Else
            If aCharPertama <> "0" Then
                aRetval = _angka(CInt(aCharPertama)) & " ratus "
            End If
        End If
        If aCharKedua = "1" Then
            aRetval = aRetval & _angka(CInt(Mid(iAngka, 2, 2)))
        Else
            If aCharKedua <> "0" Then
                aRetval = aRetval & _angka(CInt(aCharKedua)) & " puluh "
            End If
            If aCharKetiga <> "0" Then
                aRetval = aRetval & _angka(CInt(aCharKetiga))
            End If
        End If
        Return aRetval
    End Function

    Public Function spellKoma(ByVal iAngka As String) As String
        Dim aRetval As String
        aRetval = ""
        For aInt As Integer = 1 To Len(iAngka)
            If aRetval = "" Then
                aRetval = angkaToUcapKoma(Mid(iAngka, aInt, 1))
            Else
                aRetval = aRetval & " " & angkaToUcapKoma(Mid(iAngka, aInt, 1))
            End If
        Next
        Return aRetval
    End Function

    Private Function angkaToUcapKoma(ByVal iString As String) As String
        If iString = "0" Then
            Return "kosong"
        Else
            Return _angka(CInt(iString))
        End If
    End Function

End Class