Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Dim HeosINIFile As String = "HeosBridge.ini"
Dim HeosTag As String = "Heos"
Dim HeosAddressRoot As String = "HeosBridge.Root"
Dim HeosAddressSource As String = "HeosBridge.Source"
Dim HeosAddressStation As String = "HeosBridge.Station"
Dim HeosAddressAlbum As String = "HeosBridge.Album"
Dim HeosAddressStatus As String = "HeosBridge.Status"
Dim HeosAddressVolume As String = "HeosBridge.Volume"
Dim HeosAddressArtist As String = "HeosBridge.Artist"
Dim HeosAddressTitle As String = "HeosBridge.Title"
Dim HeosAddressImage As String = "HeosBridge.Image"
Dim HeosAddressPosition As String = "HeosBridge.Position"
Dim HeosAddressDuration As String = "HeosBridge.Duration"

Dim valOFF As Integer = 0
Dim valOffReq As Integer = 1
Dim valONReq As Integer = 99
Dim valON As Integer = 100
Dim valStop As Integer = 1
Dim valPlay As Integer = 3
Dim valPause As Integer = 2
Dim valPrev As Integer = 6
Dim valNext As Integer = 7

Dim type As String = "Heos Btn"
Dim debug = False

' This will be set to the 'Type' value in the HeosBridge.ini section
'--------------------------------------------------------------
Dim HeosIniName As String = "Heos"
Dim HeosMonitor As String = "HeosMonitor.vb"

' Device Buttons functionality library script.  The call to this script and its functions is embedded into the auto-generated NuVo zone and master devices during their creation.  
' It serves As the code And functionality For Each button When it Is pressed. 
' A change here will change the behavior of the Heos device, without needing to recreate the buttons in the device.


' Root Device Buttons
' Also called by ControlsStatusCmd, and ControlHeos.vb on startup
'----------------
Sub HeosConnect(Input As Object)
    Dim dvRef As Integer = Input(0)
    Dim DeviceParm As String = Input(1)
    Dim HeosPowerOnString As String = ""
    Dim HeosSourceString As String = ""
    Dim HeosPowerOffString As String = ""

    If debug Then hs.writeLog(type, "HeosConnect called by a button with command '" & DeviceParm & "' for Device ID: " & dvRef)
    ' set configurations accordingly - Heos Player ID and Output Zone
    '-----------------------------------------------------------------
    Dim HeosDenonZoneID As String = hs.getINISetting(HeosIniName, "PlayerZone", "2", HeosINIFile)
    Dim HeosDenonPlayerID As String = hs.getINISetting(HeosIniName, "PlayerID", "2", HeosINIFile)
    HeosPowerOnString = "Z" & HeosDenonZoneID & "ON" & vbCRLF
    HeosSourceString = "Z" & HeosDenonZoneID & "NET" & vbCRLF
    HeosPowerOffString = " heos://player/set_volume?pid=" & HeosDenonPlayerID & "&level=0"

    Select Case Ucase(DeviceParm)
        Case "ONREQ", valONReq
            Dim HeosStatus = hs.DeviceValueEx(dvRef)
            If HeosStatus <> valON Then
                hs.SetDeviceValueByRef(dvRef, valONReq, True)
                If sendCommandtoHeos(HeosPowerOnString, False) = 0 Then Exit Sub ' Sets up the Heos Player for Heos, in case it's off
                If sendCommandtoHeos(HeosSourceString, False) = 0 Then Exit Sub ' Sets up the Heos Player on the given zone
                hs.runscriptfunc(HeosMonitor, "Main", "Start", False, True) 'Sets up Heos monitor, to poll for status changes
                hs.waitsecs(0.01)
            End If
        Case "OFFREQ", valOffReq
            If hs.isscriptrunning(HeosMonitor) Then
                hs.SetDeviceValueByRef(dvRef, valOffReq, True)
                sendCommandtoHeos(HeosPowerOffString, False) 'Sent in order to un-stick the telnet session, which is waiting for a response 
            Else
                HeosInitialize("CLEAR")
            End If
        Case Else
            If debug Then hs.writeLog(type, "HeosConnect: unknown command '" & DeviceParm & "' for Device ID: " & dvRef)
    End Select
End Sub

' Status Device Buttons
' We can abstract this at a time when we have another Heos Player handy
' For now, we will use "Room" variable to branch, on the theory that if we create a different Heos bridge, this value will differ
' Any Control Press is other than "stop" an implied "connect" for meta-data purposes
' Object parameter is an array:
' {
'   dvRef - Typically the calling HS Device Ref for simple calls, doubles as the 'sid' for browse command
'   DeviceParm - Either the button label pressed {stop, play, pause, prev, next, voldn, volup, mute} or a command sent by another process {register, metadata, sources, browse} 
'   cid (optional) - required only for 'browse' command
'   rangeStart (optional) - default = 0 if not sent. used only for 'browse' command
'   rangeEnd (optional) - default = 100. used only for 'browse' command
'---------------------------
Sub ControlsStatusCmd(Input As Object)
    Dim dvRef As Integer = Input(0) 'Not used, so We also can pass something that Heos cares about. A bit naughty.
    Dim DeviceParm As String = Input(1)
    Dim HeosCommandString As String = ""
    Dim HeosPlayerID As String = "#PID#"
    Dim cid As String = ""
    Dim rangeStart As String = "0"
    Dim rangeEnd As String = "100"
    If ucase(DeviceParm) = "BROWSE" Then
        If Input.length > 2 Then cid = Input(2)
        If Input.length = 5 Then
            rangeStart = Input(3)
            rangeEnd = Input(4)
        End If
    End If

    ' load the sid from the existing source.  Overraide it with the dvRef if the dvRef is 
    ' between a given range - a bit lazy/naughty, but it allows for scripts to call the function with a
    ' valid sid.  
    Dim sid As Integer = hs.DeviceValue(hs.deviceExistsAddress(HeosAddressSource, True))
    If ((1 <= dvRef) AndAlso (dvRef <= 17)) Or ((1024 <= dvRef) AndAlso (dvRef <= 1028)) Then sid = dvRef 'should also check to see if we are browsing, but my dvRefs are well above 1028

    If debug = True Then hs.WriteLog(type, "ControlsStatusCmd - Processing command for Heos: " & DeviceParm)
    ' Any button press is an impled connect request
    '-----------------------------------------------
    Dim dvRoot As Integer = hs.DeviceExistsAddress(HeosAddressRoot, True)
    Dim dvVal = hs.deviceValue(dvRoot)
    If debug Then hs.writelog(type, "Value for Root " & dvRoot & " is " & dvVal)
    If hs.deviceValue(dvRoot) <> valON Then HeosConnect({dvRoot, valONReq})

    ' determine Heos type and set configurations accordingly.
    '---------------------------------------------------------
    HeosPlayerID = hs.getINISetting(HeosIniName, "PlayerID", "", HeosINIFile)
    Dim WaitForResponse As Boolean = False
    Select Case Ucase(DeviceParm)
        Case "PLAY"
            HeosCommandString = "heos://player/set_play_state?pid=" & HeosPlayerID & "&state=play"
        Case "PAUSE"
            HeosCommandString = "heos://player/set_play_state?pid=" & HeosPlayerID & "&state=pause"
        Case "STOP"
            HeosCommandString = "heos://player/set_play_state?pid=" & HeosPlayerID & "&state=stop"
        Case "PREV"
            HeosCommandString = "heos://player/play_previous?pid=" & HeosPlayerID & ""
        Case "NEXT"
            HeosCommandString = "heos://player/play_next?pid=" & HeosPlayerID & ""
        Case "VOLUP"
            HeosCommandString = "heos://player/volume_up?pid=" & HeosPlayerID & "&step=5"
        Case "VOLDN"
            HeosCommandString = "heos://player/volume_down?pid=" & HeosPlayerID & "&step=5"
        Case "MUTE"
            HeosCommandString = "heos://player/toggle_mute?pid=" & HeosPlayerID & ""
        Case "REGISTER"
            HeosCommandString = "heos://player/register_for_change_events?enable=on"
        Case "METADATA"
            HeosCommandString = "heos://player/get_now_playing_media?pid=" & HeosPlayerID & ""
        Case "SOURCES"
            HeosCommandString = "heos://browse/get_music_sources"
            WaitForResponse = True
        Case "SOURCEINFO"
            HeosCommandString = "heos://browse/get_source_info&sid=" & sid
            WaitForResponse = True
        Case "BROWSE"
            HeosCommandString = "heos://browse/browse?sid=" & sid & "&cid=" & cid & "&range=" & rangeStart & "," & rangeEnd
            If debug Then Hs.writelog(type, "Sending to Heos: " & HeosCommandString)
            WaitForResponse = True
        Case "THUMBSUP"
            HeosCommandString = "heos://browse/set_service_option?sid=" & sid & "&option=11&pid=" & HeosPlayerID
        Case "THUMBSDOWN"
            HeosCommandString = "heos://browse/set_service_option?sid=" & sid & "&option=12&pid=" & HeosPlayerID
        Case "ADDFAV"
            HeosCommandString = "heos://browse/set_service_option?option=19&pid=" & HeosPlayerID
        Case "REMOVEFAV"
            HeosCommandString = "heos://browse/set_service_option?option=20&pid=" & HeosPlayerID
        Case Else
            hs.writelog(type, "Unable to send unknown Heos command: " & DeviceParm)
            Exit Sub
    End Select

    If sendCommandtoHeos(HeosCommandString, WaitForResponse) = 1 Then
        'change status - this is theoretically handled when the data is read from the Heos data stream
    End If

End Sub


' serves 2 roles - either sending data to the denon receiver, to set up the active Heos zone, or
' send a command to the heos control port.  Some commands require a data pull, so if WaitForAnswer is set,
' the script will wait for a response which contains payload.  For other commands, it is expected to be handled 
' by the HeosMonitor, which is looking for status changes.
'-----------------------------------------------------------------------------------------------------
Function sendCommandtoHeos(ByVal HeosCommandString As String, Optional ByVal WaitForAnswer As Boolean = False) As Integer
    ' Dim Buffernum As Integer
    'debug = True

    ' set all connection values - we will need to massage this later
    ' When alternate Heos servers are available
    '----------------------------
    Dim HeosIP As String = hs.getINISetting(HeosIniName, "ServerIP", "", HeosINIFile)
    Dim HeosCtlPort As String = hs.getINISetting(HeosIniName, "ControlPort", "23", HeosINIFile)
    Dim HeosPlayerID As String = hs.getINISetting(HeosIniName, "PlayerID", "", HeosINIFile)
    Dim HeosReturnval As String = ""
    'Set port based on command string
    '----------------------------------
    If instr(1, HeosCommandString, "heos://") Then HeosCtlPort = hs.getINISetting(HeosIniName, "HeosPort", "1255", HeosINIFile)

    ' Do a little error checking 
    ' We should probably verify that the host can be reached, but for
    ' now, lets assume it can
    '--------------------------------------------
    If HeosIP = "" Then
        hs.writeLog(type, "Heos: Server IP not configured. Check the config file '" & HeosINIFile & "' for the 'ServerIP' setting under [" & HeosIniName & "]")
        hs.SetDeviceValueByRef(hs.DeviceExistsAddress(HeosAddressRoot, True), valOFF, True)
        Return 0
    End If
    If HeosPlayerID = "" Then
        hs.writeLog(type, "Heos: Player ID not configured. Check the config file '" & HeosINIFile & "' for the 'PlayerID' setting under [" & HeosIniName & "]")
        hs.SetDeviceValueByRef(hs.DeviceExistsAddress(HeosAddressRoot, True), valOFF, True)
        Return 0
    End If

    'Clean up the command String, and insert PID
    '---------------------------------------------
    HeosCommandString = HeosCommandString.Replace(vbCR, "")
    HeosCommandString = HeosCommandString.Replace("#PID#", HeosPlayerID)
    If debug Then hs.WriteLog(type, "Processing String: " & HeosCommandString)

    Using TelnetClient As New TcpClient()
        Try
            If debug Then hs.WriteLog(type, "Attempting to connect to Denon at: " & HeosIP & ":" & HeosCtlPort)
            TelnetClient.Connect(IPAddress.Parse(HeosIP), HeosCtlPort)
            If debug Then hs.WriteLog(type, "Connected to Denon at: " & HeosIP & ":" & HeosCtlPort)
            hs.waitSecs(0.01)
            ' Send data out to Heos Server
            Using TelnetOutputStream As NetworkStream = TelnetClient.GetStream
                Dim Bytes As [Byte]() = Encoding.UTF8.GetBytes(HeosCommandString & vbCRlf)
                If TelnetOutputStream.CanTimeout Then TelnetOutputStream.WriteTimeout = 1000 * 5
                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)
                ' Receive data back from Heos, if requested (Used for browse and menu commands
                If WaitForAnswer Then
                    hs.saveVar("HeosReceiveFlag", False)
                    hs.saveVar("HeosReceiveString", "")
                    Using TelnetInputStream As NetworkStream = TelnetClient.GetStream
                        Dim DataFound As Boolean = False
                        Do Until DataFound
                            Dim Buffersize As Integer = TelnetClient.ReceiveBufferSize
                            Dim Data(Buffersize) As Byte
                            Dim BytesRead As Integer = 0
                            If TelnetInputStream.CanTimeout Then TelnetInputStream.ReadTimeout = 10000 ' wait 10 seconds for data. 
                            Try
                                BytesRead = TelnetInputStream.Read(Data, Net.Sockets.Socketflags.None, Buffersize)
                            Catch
                                'hs.WriteLog(type, "Heos Monitor timeout")
                            End Try
                            Dim CommandStr = Encoding.ASCII.GetString(Data, 0, BytesRead)
                            Dim HeosDataStream = CommandStr.Split(vbCRLF)
                            For Each HeosData As String In HeosDataStream
                                HeosData = HeosData.Replace(vbCR, "")
                                HeosData = HeosData.Replace(vbLF, "")
                                HeosData = HeosData.Replace(vbCRLF, "")
                                HeosData = HeosData.Replace("\" & chr(34), chr(34))
                                HeosReturnval = HeosReturnval & HeosData
                            Next
                            If instr(1, HeosReturnval, "command under process") = 0 Then
                                DataFound = True
                            Else
                                HeosReturnval = ""
                            End If
                        Loop
                    End Using
                End If
            End Using
        Catch ex As Exception
            hs.WriteLog(type, "Heos Send Error: " & ex.toString)
            TelnetClient.Close()
            Return 0
        End Try
        TelnetClient.Close()
        If debug Then hs.WriteLog(type, "Send Complete. Disconnected from " & HeosIP)
    End Using

    ' We need to pass the results to interested processes
    ' The Assumption is that the process will read the data, and
    ' then reset the HeosReceiveFlag to False
    '-----------------------------------------------
    If debug Then hs.writelog(type, HeosReturnval)
    If HeosReturnval <> "" Then
        hs.saveVar("HeosReceiveString", HeosReturnval)
        hs.saveVar("HeosReceiveFlag", True)
    End If
    Return 1
End Function

' Initialization functions
'--------------------------
Function HeosInitialize(ByVal Args As String) As Integer
    If hs.deviceExistsAddress(HeosAddressRoot, True) = -1 Then
        hs.writelog(type, "Heos not present. Skipping Heos initialization")
        Return 0
    End If
    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressRoot, True), valOFF, False)
    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressSource, True), valOFF, True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressStation, True), "", True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressAlbum, True), "", True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressTitle, True), "", True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressArtist, True), "", True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressImage, True), "", True)
    hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressPosition, True), 0, False)
    hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressDuration, True), 0, False)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressPosition, True), "00:00:00", True)
    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressDuration, True), "00:00:00", True)
    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressStatus, True), valStop, True)
    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressVolume, True), 0, True) ' perhaps should be default?

    Select Case Ucase(Args)
        Case "STARTUP"
            ' Create globals used for passing data results
            hs.writeLog(type, "Heos Monitor (c) 2020 xs10shl. Initializing . . . ")
            hs.CreateVar("HeosReceiveFlag")
            hs.SaveVar("HeosReceiveFlag", False)
            hs.createVar("HeosReceiveString")
            hs.saveVar("HeosReceiveString", "")
            If hs.getINISetting(HeosTag, "ConnectOnStartup", "0", HeosINIFile) = "1" Then HeosConnect({Hs.DeviceExistsAddress(HeosAddressRoot, True), valONReq})
        Case "CLEAR"

    End Select
    Return 1
End Function

' Helper Functions
'--------------------------------
' Called by HeosMenu.vb
Function HeosPlay(ByVal HeosCommandString As String) As Integer
    Return sendCommandtoHeos(HeosCommandString, False)
End Function
Function HeosGet(ByVal HeosCommandString As String) As Integer
    Return sendCommandtoHeos(HeosCommandString, True)
End Function




