<%@ Language=VBScript %>
<!--METADATA TYPE="TypeLib" UUID="94A0E92D-43C0-494E-AC29-FD45948A5221"-->
<% 
	Dim oImage
	Dim oDeviceManager
	Dim oDeviceInfo
	Dim oDevice
	Dim oItem
	Dim oVector
	
    Set oDeviceManager = Server.CreateObject("WIA.DeviceManager")
    
    For Each oDeviceInfo In oDeviceManager.DeviceInfos
        If oDeviceInfo.Type = VideoDeviceType Then
            Set oDevice = oDeviceInfo.Connect
            Exit For
        End If
    Next
    
    If oDevice Is Nothing Then 
		Response.Write "There is no Video Device"
		Response.End
	End if
    
    Set oItem = oDevice.ExecuteCommand(wiaCommandTakePicture)
    
    Set oImage = oItem.Transfer
    Set oVector = oImage.FileData
    Response.BinaryWrite oVector.BinaryData
%>