Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Module Module1
    Public con As New OleDbConnection("Provider=MICROSOFT.ACE.OLEDB.12.0; Data Source=|DataDirectory|/database.accdb")
    Public rs As New OleDbDataAdapter

    Function SafeSqlLiteral(ByVal strValue, ByVal intLevel) As String


        If Not IsDbNull(strValue) Then
            If intLevel > 0 Then
                strValue = replace(strValue, "'", "''") ' Most important one! This line alone can prevent most injection attacks
                strValue = replace(strValue, "--", "")
                strValue = replace(strValue, "[", "[[]")
                strValue = replace(strValue, "%", "[%]")
            End If

            If intLevel > 1 Then
                Dim myArray As Array
                myArray = Split("xp_ ;update ;insert ;select ;drop ;alter ;create ;rename ;delete ;replace ", ";")
                Dim i, i2, intLenghtLeft As Integer
                For i = LBound(myArray) To UBound(myArray)
                    Dim rx As New Regex(myArray(i), RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                    Dim matches As MatchCollection = rx.Matches(strValue)
                    i2 = 0
                    For Each match As Match In matches
                        Dim groups As GroupCollection = match.Groups
                        intLenghtLeft = groups.Item(0).Index + len(myArray(i)) + i2
                        strValue = Left(strValue, intLenghtLeft - 1) & "&nbsp;" & right(strValue, len(strValue) - intLenghtLeft)
                        i2 += 5
                    Next
                Next
            End If

            'strValue = replace(strValue, ";", ";&nbsp;")
            'strValue = replace(strValue, "_", "[_]")

            Return strValue
        Else
            Return strValue
        End If

    End Function


End Module
