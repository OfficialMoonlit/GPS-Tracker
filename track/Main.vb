Imports System.IO
Imports System.IO.Ports
Imports System.Threading

Module Main

    Public Class GPSClass
        Public Shared GPSData As String
    End Class

    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs

    Sub Main()
        Console.WriteLine("")
        Console.WriteLine("IcyEwe v1.0 Tracking System")
        Console.WriteLine("")

        GPSClass.GPSData = "Something"

        'Thunderbirds are go.
        Dim GPS As Thread
        GPS = New Thread(AddressOf ProcessGPS)
        GPS.Start()

        Dim GSM As Thread
        GSM = New Thread(AddressOf ProcessGSM)
        GSM.Start()
    End Sub

    Private Sub BadArgs()
        'You messed up, run it again, and this time do what it says.
        Console.WriteLine("Error: GPS and GSM modem ports are both required.")
        Console.WriteLine("Usage: track.exe /gps:com1 /gsm:com2")
        Console.WriteLine("Where COM1 and COM2 are your GPS and GSM modem ports respectively.")
        End
    End Sub

    Private Sub ProcessGPS()
        Dim gpsPort As New SerialPort                            'New Serial Port: GPS.
        Dim sNMEA As String
        Dim sData As String

        Try
            With gpsPort
                .PortName = "COM12" 'GPS port
                .BaudRate = 4800
                .Open()
            End With
            If gpsPort.IsOpen Then Console.WriteLine("GPS Online.")
        Catch x As Exception
            Console.WriteLine("GPS Failed. Could not access port. Check device and settings.")
            End
        End Try

        Do
            If gpsPort.BytesToRead > 0 And gpsPort.IsOpen = True Then
                Try
                    sNMEA = gpsPort.ReadLine
                    'Console.WriteLine(sNMEA)
                    If ParseGPS(sNMEA, "Time") <> "" Then
                        sData = "Time: " & ParseGPS(sNMEA, "Time") & " (GMT) Lat: " & ParseGPS(sNMEA, "Lat") & " Long: " & ParseGPS(sNMEA, "Long") & " Alt: " & ParseGPS(sNMEA, "Alt") & " Speed: " & ParseGPS(sNMEA, "Speed") & "MPH"
                        SyncLock GPSClass.GPSData
                            GPSClass.GPSData = sData
                        End SyncLock
                        'Console.WriteLine(sData)
                    End If
                Catch ex As Exception
                    Console.WriteLine("GPS Read Failed.")
                End Try
            End If
            System.Threading.Thread.Sleep(250)
        Loop
    End Sub

    Private Sub ProcessGSM(ByVal Port As String)
        Dim gsmPort As New SerialPort                            'New Serial Port: GSM.
        Dim sGSM
        Dim sGPSSend
        Dim bPassword As Boolean

        Try
            With gsmPort
                .PortName = "COM9" 'GSM port
                .BaudRate = 9600
                .Open()
            End With

            If gsmPort.IsOpen Then Console.WriteLine("GSM Online.")
        Catch x As Exception
            Console.WriteLine("GSM Failed. Could not access port. Check device and settings.")
            End
        End Try
        'Console.WriteLine("Testing modem...")
        gsmPort.WriteLine("AT" & vbCrLf)
        'Console.WriteLine("Sent AT.")
        Threading.Thread.Sleep(100)
        gsmPort.WriteLine("AT+CMGF=1" & vbCrLf)
        'Console.WriteLine("Switching SMS Mode...")
        Threading.Thread.Sleep(100)
        gsmPort.WriteLine("AT+CNMI=1,1,0,0,0" & vbCrLf)
        Threading.Thread.Sleep(100)
        gsmPort.WriteLine("AT+COPS=?" & vbCrLf)
        Threading.Thread.Sleep(100)

        Do
            If gsmPort.BytesToRead > 0 And gsmPort.IsOpen = True Then
                Try
                    sGSM = gsmPort.ReadExisting 'Grab incoming data from GSM modem
                    'Console.WriteLine(sGSM)
                    sGSM = ParseGSM(sGSM, "SENDGPS")
                    If InStr(sGSM, "SEND:") Then
                        sGSM = Replace(sGSM, "SEND:", "")
                        gsmPort.WriteLine(sGSM & vbCrLf)
                        Console.WriteLine("Request Received, Processing...")
                    End If
                    If InStr(sGSM, "MSG:") Then
                        bPassword = True
                        sGSM = Replace(sGSM, "MSG:", "")
                        gsmPort.WriteLine("AT+CMGS=" & Chr(34) & sGSM & Chr(34) & vbCrLf)
                    End If
                    If InStr(sGSM, "SENDMSG") Then
                        If bPassword = True Then
                            sGSM = Replace(sGSM, "SENDMSG", "")
                            SyncLock GPSClass.GPSData
                                sGPSSend = GPSClass.GPSData
                            End SyncLock
                            gsmPort.WriteLine("IcyEwe v1.0 Tracking System: " & sGPSSend & Chr(26))
                            Console.WriteLine("Success.")
                        Else
                            sGSM = Replace(sGSM, "SENDMSG", "")
                            sGPSSend = "You are not authorised to use this tracker. Request failed."
                            gsmPort.WriteLine("IcyEwe v1.0 Tracking System: " & sGPSSend & Chr(26))
                            Console.WriteLine("Bad Request.")
                        End If

                    End If
                    If InStr(sGSM, "PWI:") Then
                        bPassword = False
                        sGSM = Replace(sGSM, "PWI:", "")
                        gsmPort.WriteLine("AT+CMGS=" & Chr(34) & sGSM & Chr(34) & vbCrLf)
                    End If
                Catch ex As Exception
                    Console.WriteLine("GSM Read Failed.")
                End Try
            End If
            System.Threading.Thread.Sleep(100)
        Loop

    End Sub


End Module
