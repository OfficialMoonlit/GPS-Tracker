Module Parser

    Dim gpsNextLine As String
    Dim gpsNextField As String
    Dim gpsNextChar As String
    Dim gpsCheckChar As String

    Dim gpsGGATime
    Dim gpsGGALatitude
    Dim gpsGGALatitudeNS
    Dim gpsGGALongitude
    Dim gpsGGALongitudeEW
    Dim gpsGGAFixQuality
    Dim gpsGGASatellites
    Dim gpsGGAHorizontalDilution
    Dim gpsGGAAltitude
    Dim gpsGGAAltUnits
    Dim gpsGGAGeoidHeight
    Dim gpsGGASinceLastUpdate
    Dim gpsGGADGPSStationID
    Dim gpsGGAChecksum

    Dim gpsGSADSelect
    Dim gpsGSA2D3D
    Dim gpsGSAPRN(11)
    Dim gpsGSAPDOP
    Dim gpsGSAHDOP
    Dim gpsGSAVDOP
    Dim gpsGSAChecksum

    Dim gpsGSVSentences
    Dim gpsGSVSentence
    Dim gpsGSVPRN
    Dim gpsGSVElevation
    Dim gpsGSVAzimuth
    Dim gpsGSVSNR(3)
    Dim gpsGSVChecksum

    Dim gpsVTGKMH
    Dim gpsVTGMPH

    Dim gpsDecLat
    Dim gpsDecLong

    Dim gpsGGA(16) As String
    Dim gpsGSA(18) As String
    Dim gpsGSV(21) As String
    Dim gpsVTG(8) As String

    Dim gpsGGASentence As String

    Dim XORString As String
    Dim XORSum As String
    Dim XORChar As String

    Public Const fixInvalid = 0
    Public Const fixSPS = 1
    Public Const fixDGPS = 2
    Public Const fixPPS = 3
    Public Const fixRTK = 4
    Public Const fixFRTK = 5
    Public Const fixEstimated = 6
    Public Const fixManual = 7
    Public Const fixSimulation = 8

    Public Function gpsChecksum(ByVal XORString As String)
        '//////////////////// Function to calculate GPS checksum
        XORSum = "00"
        XORChar = ""

        For i = 0 To Len(XORString) - 1
            XORSum = XORSum Xor Asc(Mid$(XORString, i + 1, 1))
        Next

        Return (Hex$(XORSum))
    End Function

    Public Function ParseGPS(ByVal Sentence As String, ByVal Request As String)
        '//////////////////// GPS sentence parser
        Dim sData As String
        gpsNextLine = Sentence

        '//////////////////// Parse Global Positioning System Fix Data (GGA)
        If Strings.Left(gpsNextLine, 6) = "$GPGGA" Then
            gpsGGA = Split(gpsNextLine, ",")
            gpsGGATime = gpsGGA(1) : gpsGGATime = Val(gpsGGATime)
            gpsGGALatitude = gpsGGA(2)
            gpsGGALatitudeNS = gpsGGA(3)
            gpsGGALongitude = gpsGGA(4)
            gpsGGALongitudeEW = gpsGGA(5)
            gpsGGAFixQuality = gpsGGA(6)
            gpsGGASatellites = gpsGGA(7)
            gpsGGAHorizontalDilution = gpsGGA(8)
            gpsGGAAltitude = gpsGGA(9)
            gpsGGAAltUnits = gpsGGA(10)
            gpsGGAGeoidHeight = gpsGGA(11)
            gpsGGAAltUnits = gpsGGA(12)
            gpsGGASinceLastUpdate = gpsGGA(13)
            gpsGGAChecksum = Strings.Right(gpsGGA(14), 3)

            '//////////////////// Assemble GGA Sentence
            gpsGGASentence = "GPGGA" & "," & gpsGGA(1) & "," & gpsGGA(2) & "," & gpsGGA(3) & "," & gpsGGA(4) & "," & gpsGGA(5) & "," & gpsGGA(6) & "," & gpsGGA(7) & "," & gpsGGA(8) & "," & gpsGGA(9) & "," & gpsGGA(10) & "," & gpsGGA(11) & "," & gpsGGA(12) & "," & gpsGGA(13) & ","

            '//////////////////// Return Requested Data
            Select Case Request
                Case "Time"
                    sData = Format(gpsGGATime, "00:00:00")
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "Long"
                    sData = LongToDecimal(gpsGGALongitude) & gpsGGALongitudeEW
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "Lat"
                    sData = LatToDecimal(gpsGGALatitude) & gpsGGALatitudeNS
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "Head"
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "Speed"
                    sData = gpsVTGMPH
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "Alt"
                    sData = gpsGGAAltitude & "m"
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
                Case "NMEA"
                    sData = gpsGGALongitude & gpsGGALongitudeEW & " " & gpsGGALatitude & gpsGGALatitudeNS
                    Return sData
            End Select

        End If

        '//////////////////// Parse Satellite Status (GSA)
        If Strings.Left(gpsNextLine, 6) = "$GPGSA" Then
            gpsGSA = Split(gpsNextLine, ",")
            gpsGSADSelect = gpsGSA(1)
            gpsGSA2D3D = gpsGSA(2)
            For gpsArrayCount = 0 To 11
                gpsGSAPRN(gpsArrayCount) = gpsGSA(gpsArrayCount + 3)
            Next
            gpsGSAPDOP = gpsGSA(15)
            gpsGSAHDOP = gpsGSA(16)
            gpsGSAChecksum = Strings.Right(gpsGSA(17), 3)
            gpsGSAVDOP = Strings.Left(gpsGSA(17), Len(gpsGSA(17)) - 4)
        End If

        '//////////////////// Parse Satellites In View (GSV)
        If Strings.Left(gpsNextLine, 6) = "$GPGSV" Then
            gpsGSV = Split(gpsNextLine, ",")

        End If

        '//////////////////// End Parse Velocity (VTG)
        If Strings.Left(gpsNextLine, 6) = "$GPVTG" Then
            gpsVTG = Split(gpsNextLine, ",")
            gpsVTGKMH = gpsVTG(7)
            gpsVTGMPH = Format(gpsVTGKMH / 1.609344, "##0.00")
            Select Case Request
                Case "Speed"
                    sData = gpsVTGMPH
                    sData = Replace(sData, vbCr, vbNullString)
                    sData = Replace(sData, vbLf, vbNullString)
                    sData = Replace(sData, vbCrLf, vbNullString)
                    sData = Trim(sData)
                    Return sData
            End Select
        End If
    End Function

    Public Function ParseGSM(ByVal Command As String, ByVal Password As String)
        '//////////////////// GSM modem stats parser
        If InStr(Command, "+CMTI: ", CompareMethod.Text) Then
            Dim aSMSIncoming
            Command = Replace(Command, Chr(34), vbNullString)
            Command = Replace(Command, "+CMT: ", vbNullString)
            aSMSIncoming = Split(Command, ",")
            Return "SEND:AT+CMGR=" & aSMSIncoming(1) & vbCrLf
        End If

        If InStr(Command, "+CMGR: ", CompareMethod.Text) Then
            Dim aSMSIncoming
            Dim sDate
            Dim sNumber
            Dim sText

            Command = Replace(Command, "+CMGR: ", vbNullString)
            Command = Replace(Command, Chr(34), vbNullString)
            Command = Replace(Command, Chr(10), vbNullString)
            Command = Replace(Command, Chr(13), vbNullString)
            sText = Strings.Right(Command, Len(Command) - InStrRev(Command, Chr(34)))
            aSMSIncoming = Split(Command, ",")
            sNumber = aSMSIncoming(1)
            sText = Replace(sText, "OK", vbNullString)
            sDate = aSMSIncoming(3)
            If InStr(Command, Password) Then
                Return "MSG:" & sNumber
            Else
                Return "PWI:" & sNumber
            End If
        End If

        If InStr(Command, ">", CompareMethod.Text) Then
            Return "SENDMSG"
        End If

        If InStr(Command, "+COPS: ", CompareMethod.Text) Then
            Dim aOperators As Array
            Dim aParams As Array
            Command = Strings.Replace(Command, "+COPS: ", vbNullString)
            Command = Strings.Replace(Command, "),", "|")
            Command = Strings.Replace(Command, "(", vbNullString)
            Command = Strings.Replace(Command, Chr(34), vbNullString)
            aOperators = Split(Command, "|")
            For o = 0 To UBound(aOperators) - 2
                aParams = Split(aOperators(o), ",")
                If Val(aParams(o)) = 2 Then
                    Console.WriteLine("")
                    Console.WriteLine("SMS Enabled, Powered by " & aParams(1) & ".")
                    Console.WriteLine("")
                End If
            Next
        End If
    End Function

    '////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////// Decimal to NMEA/NMEA to Decimal conversion, from http://www.tma.dk/gps/
    '////////////////////////////////////////////////////////////////////////////////////////////////////

    Public Enum enumLongLat
        Latitude = 1
        Longitude = 2
    End Enum

    Public Enum enumReturnformat
        WithSigns = 0
        NMEA = 1
    End Enum

    Private Function LongToDecimal(ByVal sLong As String) As Double
        Dim dLon As Double = Convert.ToDouble(sLong)
        dLon = dLon / 100
        Dim lon() As String = dLon.ToString().Split(".")
        Return lon(0).ToString() + "." + ((Convert.ToDouble(lon(1)) / 60)).ToString("#####")
    End Function

    Private Function LatToDecimal(ByVal sLat As String) As Double
        Dim dLat As Double = Convert.ToDouble(sLat)
        dLat = dLat / 100
        Dim lat() As String = dLat.ToString().Split(".")
        Return lat(0).ToString() + "." + ((Convert.ToDouble(lat(1)) / 60)).ToString("#####")
    End Function

    Public Function DecimalPosToDegrees(ByVal Decimalpos As Double, ByVal Type As enumLongLat, ByVal Outputformat As enumReturnformat, Optional ByVal SecondResolution As Integer = 2) As String
        Dim Deg As Integer = 0
        Dim Min As Double = 0
        Dim Sec As Double = 0
        Dim Dir As String = ""
        Dim tmpPos As Double = Decimalpos
        If tmpPos < 0 Then tmpPos = Decimalpos * -1 'Always do math on positive values

        Deg = CType(Math.Floor(tmpPos), Integer)
        Min = (tmpPos - Deg) * 60
        Sec = (Min - Math.Floor(Min)) * 60
        Min = Math.Floor(Min)
        Sec = Math.Round(Sec, SecondResolution)

        Select Case Type
            Case enumLongLat.Latitude '=1
                If Decimalpos < 0 Then
                    Dir = "S"
                Else
                    Dir = "N"
                End If
            Case enumLongLat.Longitude '=2
                If Decimalpos < 0 Then
                    Dir = "W"
                Else
                    Dir = "E"
                End If
        End Select
        Select Case Outputformat
            Case enumReturnformat.NMEA
                Return AddZeros(Deg, 3) & AddZeros(Min, 2) & AddZeros(Sec, 2)
            Case enumReturnformat.WithSigns
                Return Deg & "°" & Min & """" & Sec & "'" & Dir
            Case Else
                Return ""
        End Select
    End Function

    Public Function AddZeros(ByVal Value As Double, ByVal Zeros As Integer) As String
        If Math.Floor(Value).ToString.Length < Zeros Then
            Return Value.ToString.PadLeft(Zeros, CType("0", Char))
        Else
            Return Value.ToString
        End If
    End Function

    '////////////////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////// End of code from http://www.tma.dk/gps/
    '////////////////////////////////////////////////////////////////////////////////////////////////////
End Module
