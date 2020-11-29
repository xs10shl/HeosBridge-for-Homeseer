Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Collections.Generic

Dim HeosINIFile As String = "HeosBridge.ini"
Dim menuHEOSSOURCES As String = "Heos Sid"
Dim HeosControlPort As String = "ControlPort"
Dim HeosMonitorPort As String = "HeosPort"
Dim HeosTag As String = "Heos"
Dim HeosAddressRoot As String = "HeosBridge.Root"
Dim HeosAddressSource As String = "HeosBridge.Source"
Dim HeosAddressStation As String = "HeosBridge.Station"
Dim HeosAddressAlbum As String = "HeosBridge.Album"
Dim HeosAddressArtist As String = "HeosBridge.Artist"
Dim HeosAddressTitle As String = "HeosBridge.Title"
Dim HeosAddressImage As String = "HeosBridge.Image"
Dim HeosAddressPosition As String = "HeosBridge.Position"
Dim HeosAddressDuration As String = "HeosBridge.Duration"
Dim HeosAddressVolume As String = "HeosBridge.Volume"
Dim HeosAddressStatus As String = "HeosBridge.Status"
Dim type As String = "Heos Monitor"

Dim valOFF As Integer = 0
Dim valOFFReq As Integer = 1
Dim valOnReq As Integer = 99
Dim valON As Integer = 100
Dim valStop As Integer = 1
Dim valPlay As Integer = 3
Dim valPause As Integer = 2

Dim debug As Boolean = False

' Started by a runScript command, or automatically upon any Zone Source command issued 
' While Heos source Is selected the the active Zone (akin to 'auto-start')
'
' This script will Loop indefinately until HS Is restarted, an exception occurs, Or 
' if a "Disconnect" flag is set in the Heos root device.  Because we've disabled Telnet timeout,
' the only way we can "unstick" the syncronous loop is to send a "dummy" status command, which will 
' be collected by the inputstream, and processed.  The subsequent trip through the loop will cause the script to
' notice the disconnect request, close the connection, and exit.
' ------------------------------------------------------------------------------------------------
Sub Main(parms As Object)
    Dim HeosIP = hs.getINISetting(HeosTag, "ServerIP", "", HeosINIFile)
    Dim HeosCtlPort = hs.getINISetting(HeosTag, HeosControlPort, "23", HeosINIFile)
    Dim HeosMonitorPort = hs.getINISetting(HeosTag, HeosMonitorPort, "1255", HeosINIFile)
    Dim HeosPlayerID = hs.getINISetting(HeosTag, "PlayerID", "", HeosINIFile)
    Dim Buffernum As Integer

    'debug = True
    If debug Then hs.WriteLog(type, "Heos Monitor attempting to Connect to Heos at: " & HeosIP & ":" & HeosMonitorPort)

    ' Do a little error checking 
    ' We should probably verify that the host can be reached, but for
    ' now, lets assume it can
    '--------------------------------------------
    ' Get Device dvRefs 
    '------------------------
    Dim dvRootRef = hs.deviceExistsAddress(HeosAddressRoot, True)
    If dvRootRef < 0 Then
        Hs.writelog(type, "Error - unable to locate root device for HeosBridge: " & HeosAddressRoot)
        hs.SetDeviceValueByRef(dvRootRef, valOFF, False)
        Exit Sub
    End If

    If HeosIP = "" Then
        hs.writeLog(type, "Error in starting Heos Monitor: ServerIP is not configured for Heos.  Please check the configuration file '" & HeosINIFile & "'")
        hs.SetDeviceValueByRef(dvRootRef, valOFF, False)
        Exit Sub
    End If


    Using TelnetRecvClient As New TcpClient()

        Try
            TelnetRecvClient.Connect(IPAddress.Parse(HeosIP), HeosMonitorPort)
            If TelnetRecvClient.Connected Then hs.SetDeviceValueByRef(dvRootRef, valON, False)
            hs.WriteLog(type, "Heos Monitor Connected to Heos at: " & HeosIP & ":" & HeosMonitorPort)

            ' Pause to settle everything in HS
            hs.waitsecs(0.1)
            Using TelnetInputStream As NetworkStream = TelnetRecvClient.GetStream
                'Register as interested in data - heos-specific
                '------------------------------
                Dim HeosDenonCommandString As String = "heos://system/register_for_change_events?enable=on" & vbCRLF
                Dim Bytes As [Byte]() = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                Dim TelnetOutputStream As NetworkStream = TelnetRecvClient.GetStream
                If TelnetOutputStream.CanTimeout Then TelnetOutputStream.WriteTimeout = 1000 * 5
                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)

                ' Fill in Current metadata, Set Default Volume
                '----------------------------------------------
                HeosDenonCommandString = "heos://player/get_now_playing_media?pid=" & HeosPlayerID & vbCRLF
                Bytes = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)
                HeosDenonCommandString = "heos://player/get_play_state?pid=" & HeosPlayerID & vbCRLF
                Bytes = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)
                'HeosDenonCommandString = "heos://player/get_volume?pid=" & HeosPlayerID & vbCRLF
                HeosDenonCommandString = "heos://player/set_volume?pid=" & HeosPlayerID & "&level=" & hs.getVar("DefaultVolume") & vbCRLF
                Bytes = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)

                ' Loop continuously and monitor status
                '-----------------------------------------------------------
                Do While hs.DeviceValueEx(dvRootRef) = valON
                    hs.waitsecs(0.01) ' play nice with other scripts
                    'Buffernum = Buffernum + 1 ' For debugging purposes. (The loop does not run wild.)
                    Dim Buffersize As Integer = TelnetRecvClient.ReceiveBufferSize
                    Dim Data(Buffersize) As Byte
                    Dim BytesRead As Integer = 0

                    ' This code will protect us from an unknown server disconnect that we don't detect.  
                    ' We timeout every 30 seconds And re-loop. This code is 
                    ' in the place of using ascynronous coding, which is probably the better solution.
                    '-----------------------------------------------------------------------------
                    If TelnetInputStream.CanTimeout Then TelnetInputStream.ReadTimeout = 30000 ' wait 30 seconds for data. 
                    Try
                        BytesRead = TelnetInputStream.Read(Data, Net.Sockets.Socketflags.None, Buffersize)
                    Catch
                        ' Contine looping
                    End Try
                    Dim CommandStr = Encoding.ASCII.GetString(Data, 0, BytesRead)
                    'If debug Then hs.WriteLog(type, "Data received: " & Buffernum & ": " & CommandStr)

                    ' In case more than one upate is in the buffer, split the buffer into separate commands,
                    ' And proccess each one in the order reeived
                    '-------------------------------------------------------------------------------
                    Dim HeosDataStream = CommandStr.Split(vbCRLF)
                    For Each HeosData As String In HeosDataStream
                        HeosData = HeosData.Replace(vbCR, "")
                        HeosData = HeosData.Replace(vbLF, "")
                        HeosData = HeosData.Replace(vbCRLF, "")
                        HeosData = HeosData.Replace("\" & chr(34), chr(34))

                        ' These next tests are Heos-specific.  May need to abstract them to support other Players
                        '____________________________________________________________________________________
                        If HeosData.Length < 3 Then Continue For
                        If Left(HeosData, 9) <> "{" & chr(34) & "heos" & chr(34) & ": " Then Continue For ' check to make sure we have a Heos-formatted result, saves from having to check later

                        ' ready to process. Grab the header data, check pid,
                        ' and act based on the returned command
                        '-----------------------------------------------
                        'If debug Then hs.writelog(type, HeosData)
                        Dim HeosHeader As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
                        If Not HeosHeaderParse(HeosData, HeosHeader) Then
                            hs.writelog(type, "Unable to process Heos header: " & HeosData)
                            Continue For
                        End If
                        If Not HeosHeader.ContainsKey("pid") Then Continue For 'heos System-level commands, like registering for change events
                        ' at this point we know we have a properly formatted Heos Header
                        ' so it's simply a matter of branching based on the command, 
                        ' updating the device, and sending a notify that changes were made
                        ' in place of a notify, we could have registered a callback instead
                        '--------------------------------------------------------------
                        If HeosHeader("pid") <> HeosPlayerID Then Continue For 'We first make sure this is the playerID we are monitoring.  This will change for multi-pid
                        Select Case HeosHeader("command")
                            Case "event/player_state_changed", "player/get_play_state"
                                If debug Then hs.writelog(type, "Play State Retrieved: " & HeosHeader("state"))
                                HeosHeader("update_metadata") = "true"
                                Select Case HeosHeader("state")
                                    Case "stop"
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressStation, True), "", True)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressTitle, True), "", True)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressTitle, True), "", True)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressArtist, True), "", True)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressImage, True), "", True)
                                        hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressPosition, True), 0, False)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressPosition, True), "00:00:00", True)
                                        hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressDuration, True), 0, False)
                                        hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressDuration, True), "00:00:00", True)
                                        hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressStatus, True), valStop, True)
                                    Case "play"
                                        hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressStatus, True), valPlay, True)
                                    Case "pause"
                                        hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressStatus, True), valPause, True)
                                End Select
                            Case "event/player_now_playing_progress"
                                'if debug then hs.writelog(type, "Cur_pos: " & HeosHeader("cur_pos"))
                                If Not isnumeric(HeosHeader("cur_pos")) Then HeosHeader("cur_pos") = "0"
                                If Not isnumeric(HeosHeader("duration")) Then HeosHeader("duration") = "0" ' bug in AP2 that doesn't send duration on iPhone
                                Dim HeosPosTime As TimeSpan = New TimeSpan(0, 0, 0, 0, HeosHeader("cur_pos"))
                                Dim HeosDurTime As TimeSpan = New TimeSpan(0, 0, 0, 0, HeosHeader("duration"))
                                If debug Then hs.writelog(type, "Clock Position : " & HeosPosTime.toString() & " Duration: " & HeosDurTime.toString())
                                hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressPosition, True), CInt(HeosHeader("cur_pos")), False)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressPosition, True), HeosPosTime.toString(), True)
                                ' if the duration changes, then we will assume the song changed
                                ' and we need to flag to re-send song duration 
                                If HeosDurTime.ToString() <> hs.DeviceString(hs.deviceExistsAddress(HeosAddressDuration, True)) Then
                                    HeosHeader("update_metadata") = "true"
                                    hs.setDeviceValueByRef(hs.deviceExistsAddress(HeosAddressDuration, True), CInt(HeosHeader("duration")), False)
                                    hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressDuration, True), HeosDurTime.toString(), True)
                                End If
                            Case "event/player_now_playing_changed"
                                'Call to request metadata
                                HeosDenonCommandString = "heos://player/get_now_playing_media?pid=" & HeosPlayerID & vbCRLF
                                Bytes = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                                TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)
                            Case "event/player_volume_changed"
                                If debug Then hs.writelog(type, "Volume: " & HeosHeader("level") & " Mute:  " & HeosHeader("mute"))
                                If HeosHeader("mute") = "on" Then
                                    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressVolume, True), valOFF, True)
                                Else
                                    hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressVolume, True), HeosHeader("level"), True)
                                End If
                                ' Treat volume change as a command to change current zone
                                ' Implementation-specific to me
                                '--------------------------------------------------
                                'If HeosHeader("level") <> hs.GetVar("DefaultVolume") Then
                                '    hs.runscriptfunc("sendNuVoZoneVolume.vb", "Main", "Heos", False, True)
                                '    hs.waitsecs(0.1)
                                '  HeosDenonCommandString = "heos://player/set_volume?pid=" & HeosPlayerID & "&level=" & hs.getVar("DefaultVolume") & vbCRLF
                                '   Bytes = Encoding.UTF8.GetBytes(HeosDenonCommandString)
                                '   TelnetOutputStream.Write(Bytes, Net.Sockets.Socketflags.None, Bytes.Length)
                                'End If
                            Case "player/get_volume", "player/set_volume"
                                If debug Then hs.writelog(type, "Volume: " & HeosHeader("level"))
                                hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressVolume, True), HeosHeader("level"), True)
                                HeosHeader("update_metadata") = "true"
                            Case "player/get_now_playing_media"
                                ' Fill in Device Metadata
                                HeosData = jsonGetPayload(HeosData)
                                If HeosData = "" Then Continue For
                                Dim HeosPayload As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
                                ' If Debug Then hs.writelog(Type, src)
                                Dim cnt As Integer = jsonParse(HeosData, HeosPayload)
                                If cnt = 0 Then Continue For
                                If debug Then hs.Writelog(type, "sid: " & HeosPayload("sid") & "Artist: " & HeosPayload("artist") & " Song: " & HeosPayload("song") & " Image: " & HeosPayload("image_url") & " Album: " & HeosPayload("album"))
                                hs.SetDeviceValueByRef(hs.deviceExistsAddress(HeosAddressSource, True), CInt(HeosPayload("sid")), True)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressStation, True), HeosPayload("station"), True)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressAlbum, True), HeosPayload("album"), True)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressTitle, True), HeosPayload("song"), True)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressArtist, True), HeosPayload("artist"), True)
                                hs.SetDeviceString(hs.deviceExistsAddress(HeosAddressImage, True), HeosPayload("image_url"), True)
                                HeosHeader("update_metadata") = "true"
                            Case Else
                                ' Someone might want to see the data, if not handled above. Send it if the Send Flag isnt set.
                                ' If debug Then hs.WriteLog(type, "Data received.  Passing to HS via HeosData if the buffer is clear. " & Buffernum & ": " & HeosData)
                                ' If hs.getVar("HeosReceiveFlag") = False Then
                                ' hs.saveVar("HeosReceiveString", HeosData)
                                ' hs.SaveVar("HeosReceiveFlag", True)
                                ' End If
                        End Select
                        ' finish the loop with sending meta data if there's an update
                        If HeosHeader("update_metadata") = "true" Then
                            '-------------------------------------------------------------------------------
                            ' PLACE CODE/FUNCTION CALL HERE TO UPDATE ANY SCRIPTS OF METADATA UPDATES
                            ' Place whatever code here to facilitate a metadata update
                            ' example: hs.runscriptfunc("UpdateHomeMonitor.vb", "Main", "UPDATE", False, False)
                            '----------------------------------------------------------------------------------
                           ' hs.runscriptfunc("sendMetaData.vb", "Main", "Heos", False, False)
                        End If
                    Next
                Loop
            End Using
        Catch ex As Exception
            hs.WriteLog(type, "Heos Monitor Error: " & ex.toString)
        End Try

        TelnetRecvClient.Close()
        hs.waitsecs(0.02)
        hs.SetDeviceValueByRef(dvRootRef, valOFF, False)
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
        hs.WriteLog(type, "Heos Monitor Disconnected")
        Hs.waitsecs(0.2)
       ' hs.runscriptfunc("sendMetaData.vb", "Main", "Heos", False, False)
    End Using

End Sub

Function HeosHeaderParse(ByVal HeosData As String, JsonDict As Dictionary(Of String, String)) As Boolean
    JsonDict.Add("level", "")
    JsonDict.Add("mute", "")
    JsonDict.Add("cur_pos", "0")
    JsonDict.Add("duration", "0")
    JsonDict.Add("update_metadata", "false")
    HeosData = jsonGetHeader(HeosData)
    If HeosData = "" Then Return False
    Dim pairCount = jsonParse(HeosData, JsonDict)
    If pairCount = 0 Then Return False
    Dim MsgData() As String = split(JsonDict("message"), "&")
    For Each pairData As String In MsgData
        Dim Pair() As String = split(pairData, "=")
        If debug Then hs.writelog(type, "JSON Pair - Name: '" & Pair(0) & "' Value: '" & Pair(1) & "'")
        JsonDict(Pair(0)) = Pair(1)
    Next
    Return True
End Function

' json Functions
' Specific to Heos
'-------------------------------
Function jsonParse(ByVal HeosData As String, JsonDict As Dictionary(Of String, String)) As Integer
    Dim Count As Integer = 0
    HeosData = HeosData.Trim(" ", "[", "{", "]", "}")
    Dim HeosPair() As String = Split(HeosData, ", ") ' wont work for embedded ', ' 
    For Each dataSet As String In HeosPair
        'dataSet = dataSet.Trim(" ", chr(34))
        Dim pair() As String = split(dataSet, chr(34) & ": ") ' wont work for imbedded '": ' - highly unlikely
        If pair.count <> 2 Then Continue For ' give up on the pair if data is corrupted
        pair(0) = pair(0).Trim(" ", chr(34))
        If pair(0) = "" Then Continue For
        pair(1) = pair(1).Trim(" ")
        If left(pair(1), 1) = Chr(34) Then pair(1) = right(pair(1), Len(pair(1)) - 1)
        If right(pair(1), 1) = chr(34) Then pair(1) = left(pair(1), len(pair(1)) - 1)
        If debug Then hs.writelog(type, "JSON Pair - Name: '" & pair(0) & "' Value: '" & pair(1) & "'")
        JsonDict.Add(pair(0), pair(1))
        Count += 1
    Next
    If Not JsonDict.ContainsKey("station") Then JsonDict.Add("station", "")
    If Not JsonDict.ContainsKey("album") Then JsonDict.Add("album", "")
    Return Count
End Function

Function jsonGetPayload(ByVal HeosData As String) As String
    If debug Then hs.writelog(type, "HeosData: " & HeosData)

    Dim PayloadTag = instr(1, HeosData, chr(34) & "payload" & chr(34))
    If PayloadTag > 0 Then
        HeosData = right(HeosData, Len(HeosData) - PayloadTag)
        ' we need to see if the options tag is here.
        Dim optTag = instr(1, HeosData, ", " & chr(34) & "options" & chr(34))
        If optTag = 0 Then
            ' Heos monitor payloads are different from 'browse' payloads
            'HeosData = left(HeosData, instrrev(HeosData, "]") - 1)
        Else
            HeosData = left(HeosData, optTag - 1)
        End If
    Else
        Return ""
    End If
    If debug Then hs.writelog(type, "Payload: " & HeosData)
    Return HeosData
End Function

Function jsonGetHeader(ByVal HeosData As String) As String
    HeosData = right(HeosData, len(HeosData) - 9) 'remove heos: tag
    Dim x As Integer = instr(1, HeosData, ", " & chr(34) & "payload" & chr(34))
    If x > 0 Then HeosData = left(HeosData, x - 1)
    If debug Then hs.writelog(type, "Header: " & HeosData)
    Return HeosData
End Function
