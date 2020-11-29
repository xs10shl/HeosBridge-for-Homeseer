Public Const HeosINIFile As String = "HeosBridge.ini"
Dim menuHEOSSOURCES As String = "Heos Sid"
Public Const HeosImageDir As String = "\Images\HeosBridge\"
Public Const HeosImage As String = "Heos.png"

Dim HSType As String = "HeosBridge" ' This is used to look up the devices in other scripts, so if it's changed here, it has to be changed elsewehere
Dim HeosAddressRoot As String = "HeosBridge.Root"
Dim HeosAddressSource As String = "HeosBridge.Source"
Dim HeosAddressStation As String = "HeosBridge.Station"
Dim HeosAddressAlbum As String = "HeosBridge.Album"
Dim HeosAddressArtist As String = "HeosBridge.Artist"
Dim HeosAddressTitle As String = "HeosBridge.Title"
Dim HeosAddressImage As String = "HeosBridge.Image"
Dim HeosAddressDuration As String = "HeosBridge.Duration"
Dim HeosAddressPosition As String = "HeosBridge.Position"
Dim HeosAddressStatus As String = "HeosBridge.Status"
Dim HeosAddressVolume As String = "HeosBridge.Volume"

Public Sub Main(ByVal Parms As Object)
    Dim Type As String = "Create Device"

    ' Use this routine at your own risk.  This routine has code which alters the state of your Homeseer database 
    ' by creating and modifying devices in Homeseer.  It has only been sparingly tested, and there my be unknown 
    ' or untried or previously unknown cases which may cause this routine to corrupt your homeseer system.
    ' As a precaution, create a backup of your entire Homeseer system prior to executing this script.
    ' Do NOT use this software unless you feel confident that you can restore your system to the state it was in
    ' prior to running this routine.  

    ' Select images to represent your sources in the Images directory, or browse and select your 
    ' image of choice in any subdirectory in the /html subfolder within the Homeseer system
    ' These images will display on the device screen and on the HSTouch default screen
    ' You will want to select an appropriate image for the Source you are controlling, so it's displayed
    ' Properly.
    '
    '  |  |  |  |  |  |  |  |  |  |
    '  v  v  v  v  v  v  v  v  v  v
    '-----------------------------------------------------------------------------------
    Dim ImageDir = hs.getINISetting("Source Images", "ImageDir", HeosImageDir, HeosINIFile)
    Dim HeosTag As String = "Heos"
    Dim HeosType As String = hs.getINISetting(HeosTag, "Name", "Denon", HeosINIFile)
    Dim HeosFloor = HeosTag
    Dim HeosName As String = hs.getINISetting(HeosTag, "Name", "Heos", HeosINIFile)
    Dim HeosIP As String = hs.getINISetting(HeosTag, "ServerIP", "127.0.0.1", HeosINIFile)
    Dim HeosCtlPort As String = hs.getINISetting(HeosTag, "ControlPort", "23", HeosINIFile)
    Dim HeosHeosPort As String = hs.getINISetting(HeosTag, "HeosPort", "1255", HeosINIFile)
    Dim HeosPlayerID As String = hs.getINISetting(HeosTag, "PlayerID", "0", HeosINIFile)
    Dim HeosPlayZone As String = hs.getINISetting(HeosTag, "PlayZone", "3", HeosINIFile)
    Dim HeosControlScript As String = "HeosButtons.vb"

    ' Select default images for the Heos control system.
    ' If you are happy with the default images provided, there is no
    ' need to change these. The root directory is the Homeseer /html subdirectory.
    ' ---------------------------------------------------
    Dim ReadyImage As String = ImageDir & "power.png"
    Dim PlayPauseImage As String = ImageDir & "playpause.png"
    Dim PrevImage As String = ImageDir & "previous.png"
    Dim NextImage As String = ImageDir & "next.png"
    Dim PlayingImage As String = ImageDir & "play.png"
    Dim PauseImage As String = ImageDir & "pause.png"
    Dim StopImage As String = ImageDir & "stop.png"
    Dim VolZeroImage As String = ImageDir & "mute.png"
    Dim VolLowImage As String = ImageDir & "volumedown.png"
    Dim VolMedImage As String = ImageDir & "volumemed.png"
    Dim VolHighImage As String = ImageDir & "volumeup.png"

    '--------------------------------------------------------------------------------------
    '  ^  ^  ^  ^  ^  ^  ^  ^  ^  ^  
    '  |  |  |  |  |  |  |  |  |  |
    '
    ' **** YOU SHOULD NOT HAVE TO MODIFY ANY CODE BELOW THIS COMMENT ****
    ' Create the corresponding devices in Homeseer.  If the devices already exist, make any retrofits as needed.
    ' ------------------------------------------------
    hs.WriteLog(Type, "Creating Heos Controller:  " & HeosFloor & " " & HeosType & "  '" & HeosName & "'")

    ' Declare the variables we need to create new devices here
    '----------------------------------------------------------
    Dim dv As Scheduler.Classes.DeviceClass = Nothing
    Dim root_dv As Scheduler.Classes.DeviceClass = Nothing
    Dim dvRef As Integer = 0
    Dim BaseRef As Integer = 0
    Dim Pair As VSPair
    Dim GPair As VGPair
    Dim HSName As String = ""

    ' Reset all device refrences
    '----------------------------------
    dv = Nothing
    root_dv = Nothing
    BaseRef = 0

    ' Begin process of creating/modifying the devices.  First check to see if it exists, 
    ' Create from scratch if it doesnt, over-write it's appearance with the new formats if it does.
    ' This is a destructive process, all previous configurations of a zone of the same name will 
    ' be erased, but the DEvice Ref will remain the same, making updates of HS Events much easier 
    '--------------------------------------------------------
    HSName = "Root"
    dvRef = hs.DeviceExistsAddress(HeosAddressRoot, True)

    If dvRef <> -1 Then
        ' Root device is already created.  Instead of deleting it, we are going to delete and regenerate the dv pairs within the zone.  
        ' This will allow us to save the reference of the Zone, and simply re-define it's presentation...in theory
        '---------------------------------------------------------------------------------------------------------
        hs.WriteLog(Type, "Device " & HSName & " exists with Ref of " & dvRef & ". Reconfiguring . . .")
        hs.DeviceScriptButton_DeleteAll(dvRef)
        hs.DeviceVSP_ClearAll(dvRef, True)
        hs.DeviceVGP_ClearAll(dvRef, True)

        'Retrieve the actual pointer to the existing device, via it's reference number
        '-------------------------------------------------------------------------------
        dv = hs.getdevicebyref(dvRef)
    Else
        ' Starting from Scratch. Create a new device, and assign Floor and Room parameters
        hs.WriteLog(Type, "Creating Device - " & HSName & " for " & HeosType)
        dv = hs.NewDeviceEx(HSName)

        ' Assign Floor, Room and Type values here
        '--------------------------------
        dv.location2(hs) = HeosFloor
        dv.location(hs) = HeosType
    End If

    If dv IsNot Nothing Then
        ' Create all the control and Status values for the Device
        '-----------------------------------------------------------
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Disconnected"
        Pair.Value = 0
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Disconnecting . . ."
        Pair.Value = 1
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Connecting . . ."
        Pair.Value = 99
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Connected: " & HeosIP
        Pair.Value = 100
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 0
        GPair.Graphic = "" 'StopImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 1
        GPair.Graphic = "" 'StopImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 99
        GPair.Graphic = "" 'PlayingImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 100
        GPair.Graphic = "" 'PlayingImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        ' Create all the Graphics for the device
        '------------------------------------------
        ' Using default Homeseer Graphics for on/off

        ' Final device housekeeping
        '--------------------------------------------
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
        dv.Device_Type_String(hs) = HSType
        dv.Address(hs) = HeosAddressRoot
        hs.setdevicevaluebyref(dv.Ref(hs), 0, True)

        ' Create all the Button Code for the device
        '-----------------------------------------
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Off", 0, HeosControlScript, "HeosReceive", "Off", 0, 0, 0)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "On", 100, HeosControlScript, "HeosReceive", "On", 0, 0, 0)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Disconnect", 1, HeosControlScript, "HeosConnect", "OffReq", 1, 2, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Connect", 99, HeosControlScript, "HeosConnect", "OnReq", 1, 1, 1)

        'Mark this Device as the Root Device, and save its reference for later
        '----------------------------------------------
        root_dv = dv
        dv.Relationship(hs) = Enums.eRelationship.Parent_Root
        BaseRef = dv.Ref(hs)
    Else
        hs.WriteLog(Type, "Error in creating new dv for Heos: " & HeosType)
    End If

    ' Create the various Devices containing all the Track information
    ' Probably overkill, but we are doing it for conformity sake
    '--------------------------------------------------------------
    Dim HeosArray() As String = {"Source", "Station", "Album", "Title", "Artist", "Image", "Duration", "Position"}
    For Each HeosDv As String In HeosArray
        HSName = HeosDv
        dvRef = hs.DeviceExistsAddress(HSType & "." & HSName, True)

        If dvRef <> -1 Then
            ' Device is already created.  Instead of deleting it, we are going to delete and regenerate the dv pairs within the zone.  
            ' This will allow us to save the reference of the Zone, and simply re-define it's presentation...in theory
            '---------------------------------------------------------------------------------------------------------
            hs.WriteLog(Type, "Device " & HSName & " exists with Ref of " & dvRef & ". Reconfiguring . . .")
            hs.DeviceScriptButton_DeleteAll(dvRef)
            hs.DeviceVSP_ClearAll(dvRef, True)
            hs.DeviceVGP_ClearAll(dvRef, True)

            'Retrieve the actual pointer to the existing device, via it's reference number
            '-------------------------------------------------------------------------------
            dv = hs.getdevicebyref(dvRef)
        Else
            ' Starting from Scratch. Create a new device, and assign Floor and Room parameters
            hs.WriteLog(Type, "Creating Device - " & HSName & " for " & HeosType)
            dv = hs.NewDeviceEx(HSName)

            ' Assign Floor, Room and Type values here
            '--------------------------------
            dv.location2(hs) = HeosFloor
            dv.location(hs) = HeosType
        End If

        If dv IsNot Nothing Then
            ' Create all the control and Status values for the Device
            '-----------------------------------------------------------
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Status = ""
            Pair.Value = -1
            hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

            ' Souce needs 17 values, and 17 images
            If HeosDv = "Source" Then
                For src As Integer = 0 To 17
                    'Value
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = src
                    Pair.Status = HS.getINISetting(menuHEOSSOURCES, src.ToString(), "Source " & src.ToString(), HeosINIFile)
                    hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)
                    'image
                    GPair = New VGPair
                    GPair.PairType = VSVGPairType.SingleValue
                    GPair.set_value = src
                    GPair.Graphic = ImageDir & "source" & src.ToString() & ".png"
                    hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)
                Next
                ' Heos-specific sources
                For src As Integer = 1024 To 1028
                    'Value
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = src
                    Pair.Status = HS.getINISetting(menuHEOSSOURCES, src.ToString(), "Source " & src.ToString(), HeosINIFile)
                    hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)
                    'image
                    GPair = New VGPair
                    GPair.PairType = VSVGPairType.SingleValue
                    GPair.set_value = src
                    GPair.Graphic = ImageDir & "heos.png" ' could make separate png, depending on selection
                    hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)
                Next
            End If
            ' Station needs an "add/Remove Heos Favorite button
            If HeosDv = "Station" Then
                hs.DeviceScriptButton_AddButton(dv.ref(hs), "Add As Fav", 1, HeosControlScript, "ControlsStatusCmd", "AddFav", 1, 1, 1)
                hs.DeviceScriptButton_AddButton(dv.ref(hs), "Remove Fav", 2, HeosControlScript, "ControlsStatusCmd", "RemoveFav", 1, 2, 1)
            End If

            ' Title needs thumbs Up/Down button
            If HeosDv = "Title" Then
                hs.DeviceScriptButton_AddButton(dv.ref(hs), "Thumbs Up", 1, HeosControlScript, "ControlsStatusCmd", "ThumbsUp", 1, 1, 1)
                hs.DeviceScriptButton_AddButton(dv.ref(hs), "Thumbs Dn", 2, HeosControlScript, "ControlsStatusCmd", "ThumbsDown", 1, 2, 1)
            End If

            ' Duration and Position also need a zero value
            '----------------------------------------------
            If HeosDv = "Duration" Or HeosDv = "Position" Then
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Status = "00:00:00"
                Pair.Value = 0
                hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.status)
                Pair.PairType = VSVGPairType.Range
                Pair.RangeStart = 0
                Pair.RangeEnd = 10000000
                Pair.RangeStatusDecimals = 0
                Pair.Status = 0
                hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)
            End If

            ' Final device housekeeping
            '--------------------------------------------
            dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
            dv.Device_Type_String(hs) = HSType
            dv.Address(hs) = HSType & "." & HSName
            hs.setdevicevaluebyref(dv.Ref(hs), -1, True)

            ' Add to root_dv (Power) as a child 
            '-----------------------------------------
            root_dv.AssociatedDevice_Add(hs, dv.Ref(hs))
            dv.Relationship(hs) = Enums.eRelationship.Child
            dv.AssociatedDevice_Add(hs, BaseRef)
        Else
            hs.WriteLog(Type, "Error in creating new dv for Heos: " & HeosName & ": dv.NewDeviceEx failed for Device " & HSName)
        End If
    Next

    ' Create Controls device, and add as a child to Root (Power)
    '-----------------------------------------
    HSName = "Status"
    dvRef = hs.DeviceExistsAddress(HeosAddressStatus, True)

    If dvRef <> -1 Then
        ' Zone device is already created.  Instead of deleting it, we are going to delete and regenerate the dv pairs within the zone.  
        ' This will allow us to save the reference of the Zone, and simply re-define it's presentation...in theory
        '---------------------------------------------------------------------------------------------------------
        hs.WriteLog(Type, "Device " & HSName & " exists for " & HeosType & " with Ref of " & dvRef & " Reconfiguring . . .")
        hs.DeviceScriptButton_DeleteAll(dvRef)
        hs.DeviceVSP_ClearAll(dvRef, True)
        hs.DeviceVGP_ClearAll(dvRef, True)

        'Retrieve the actual pointer to the existing device, via it's reference number
        '-------------------------------------------------------------------------------
        dv = hs.getdevicebyref(dvRef)
    Else
        ' Starting from Scratch. Create a new device, and assign Floor and Room parameters
        hs.WriteLog(Type, "Creating Device - " & HSName & " for " & HeosType)
        dv = hs.NewDeviceEx(HSName)

        ' Assign Floor and Room values here
        '--------------------------------
        dv.location2(hs) = HeosFloor
        dv.location(hs) = HeosType
    End If

    If dv IsNot Nothing Then

        ' Assign Floor and Room values
        '--------------------------------
        dv.location2(hs) = HeosFloor
        dv.location(hs) = HeosType

        ' Create all the control and Status values for the Device
        '-----------------------------------------------------------
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Unknown"
        Pair.Value = -1
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Stop"
        Pair.Value = 1
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Idle"
        Pair.Value = 1
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Play"
        Pair.Value = 2
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Paused"
        Pair.Value = 2
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Pause"
        Pair.Value = 3
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Playing"
        Pair.Value = 3
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Go To Beginning"
        Pair.Value = 6
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Status = "Go To End"
        Pair.Value = 7
        'Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        'Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
        'Pair.PairType = VSVGPairType.SingleValue
        'Pair.Status = "Menu"
        'Pair.Value = 8
        'Pair.Render = Enums.CAPIControlType.Button
        'hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)


        ' Create all the Graphics for the device
        '------------------------------------------
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 0
        GPair.Graphic = ReadyImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 1
        GPair.Graphic = StopImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 2
        GPair.Graphic = PauseImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 3
        GPair.Graphic = PlayingImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 6
        GPair.Graphic = PrevImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 7
        GPair.Graphic = NextImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)



        ' Final device housekeeping
        '--------------------------------------------
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
        dv.Device_Type_String(hs) = HSType
        dv.Address(hs) = HeosAddressStatus
        hs.setdevicevaluebyref(dv.Ref(hs), 1, True)

        ' Create all the Button Code for the device
        '-----------------------------------------
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Stop", 1, HeosControlScript, "ControlsStatusCmd", "Stop", 1, 1, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Play", 3, HeosControlScript, "ControlsStatusCmd", "Play", 1, 2, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Pause", 2, HeosControlScript, "ControlsStatusCmd", "Pause", 1, 3, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Prev", 6, HeosControlScript, "ControlsStatusCmd", "Prev", 2, 1, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Next", 7, HeosControlScript, "ControlsStatusCmd", "Next", 2, 2, 1)

        ' Add to root_dv (Power) as a child 
        '-----------------------------------------
        root_dv.AssociatedDevice_Add(hs, dv.Ref(hs))
        dv.Relationship(hs) = Enums.eRelationship.Child
        dv.AssociatedDevice_Add(hs, BaseRef)

    Else
        hs.WriteLog(Type, "Error in creating new dv for Heos " & HeosName & ": dv.NewDeviceEx failed for Device " & HSName)
    End If

    ' Create volume device, and add as a child to Root (Power)
    '-----------------------------------------
    HSName = "Volume"
    dvRef = hs.DeviceExistsAddress(HeosAddressVolume, True)

    If dvRef <> -1 Then
        ' Zone device is already created.  Instead of deleting it, we are going to delete and regenerate the dv pairs within the zone.  
        ' This will allow us to save the reference of the Zone, and simply re-define it's presentation...in theory
        '---------------------------------------------------------------------------------------------------------
        hs.WriteLog(Type, "Device " & HSName & " exists for " & HeosType & " with Ref of " & dvRef & " Reconfiguring . . .")
        hs.DeviceScriptButton_DeleteAll(dvRef)
        hs.DeviceVSP_ClearAll(dvRef, True)
        hs.DeviceVGP_ClearAll(dvRef, True)

        'Retrieve the actual pointer to the existing device, via it's reference number
        '-------------------------------------------------------------------------------
        dv = hs.getdevicebyref(dvRef)
    Else
        ' Starting from Scratch. Create a new device, and assign Floor and Room parameters
        hs.WriteLog(Type, "Creating Device - " & HSName & " for " & HeosType)
        dv = hs.NewDeviceEx(HSName)

        ' Assign Floor and Room values here
        '--------------------------------
        dv.location2(hs) = HeosFloor
        dv.location(hs) = HeosType
    End If

    If dv IsNot Nothing Then
        ' Create all the control and Status values for the Device
        '-----------------------------------------------------------
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.RangeStart = 0
        Pair.RangeEnd = 100
        Pair.RangeStatusSuffix = "%"
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Pair.RangeStatusDecimals = 0
        'Pair.ControlUse = ePairControlUse._Dim 'HSTouch?
        hs.DeviceVSP_AddPair(dv.Ref(hs), Pair)

        ' Create all the Graphics for the device
        '------------------------------------------
        GPair = New VGPair
        GPair.PairType = VSVGPairType.SingleValue
        GPair.set_value = 0
        GPair.Graphic = VolZeroImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.Range
        GPair.RangeStart = 1
        GPair.RangeEnd = 33
        GPair.Graphic = VolLowImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.Range
        GPair.RangeStart = 34
        GPair.RangeEnd = 66
        GPair.Graphic = VolMedImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        GPair = New VGPair
        GPair.PairType = VSVGPairType.Range
        GPair.RangeStart = 67
        GPair.RangeEnd = 100
        GPair.Graphic = VolHighImage
        hs.DeviceVGP_AddPair(dv.Ref(hs), GPair)

        ' Create all the Button Code for the device
        '-----------------------------------------
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Vol Dn", 101, HeosControlScript, "ControlsStatusCmd", "VolDn", 1, 1, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Vol Up", 102, HeosControlScript, "ControlsStatusCmd", "VolUp", 1, 2, 1)
        hs.DeviceScriptButton_AddButton(dv.ref(hs), "Mute", 103, HeosControlScript, "ControlsStatusCmd", "Mute", 1, 3, 1)

        ' Final device housekeeping
        '--------------------------------------------
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
        dv.Device_Type_String(hs) = HSType
        dv.Address(hs) = HeosAddressVolume
        hs.setdevicevaluebyref(dv.Ref(hs), 0, True)

        ' Add to root_dv (Power) as a child 
        '-----------------------------------------
        root_dv.AssociatedDevice_Add(hs, dv.Ref(hs))
        dv.Relationship(hs) = Enums.eRelationship.Child
        dv.AssociatedDevice_Add(hs, BaseRef)
    Else
        hs.WriteLog(Type, "Error in creating new dv for Heos " & HeosName & ": dv.NewDeviceEx failed for Device " & HSName)
    End If

    hs.WriteLog(Type, "Finished Creating/Updating new Heos Devices for: " & HeosName & ". Adjust the display filters on the Device management page to view")
End Sub