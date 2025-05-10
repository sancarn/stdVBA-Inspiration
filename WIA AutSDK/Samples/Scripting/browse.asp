<%@ Language=VBScript %>
<% 
	Dim oFSO
	Dim oFolder
	Dim oFolders
	Dim oFiles
	Dim oFile
	Dim oImage
	Dim oProcess
	Dim oThumb
	Dim oProperty
	Dim oRational
	Dim oVector
	
	Dim bGotPics
	Dim i
	Dim s
	
	Dim sCmd
	Dim sDir
	Dim sFile
	
	sCmd = Request.QueryString("C")
	
	if sCmd = "Image" then
		sFile = Request.QueryString("F")
		set oImage = Server.CreateObject("WIA.ImageFile")
		oImage.LoadFile sFile

		If ((oImage.FormatID <> "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}") And _
			(oImage.FormatID <> "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}")) Then
			set oProcess = Server.CreateObject("WIA.ImageProcess")

			oProcess.Filters.Add "{42A6E907-1D2F-4b38-AC50-31ADBE2AB3C2}" 'Convert
			oProcess.Filters(1).Properties(1).Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" 'JPEG
			oProcess.Filters(1).Properties(2).Value = 100 'JPEG Compression
			
			set oImage = oProcess.Apply(oImage)
		End if
		if Not oImage is Nothing then
			set oVector = oImage.FileData
			Response.BinaryWrite oVector.BinaryData
		End if
	elseif sCmd = "BigThumb" then
		sFile = Request.QueryString("F")
		set oImage = Server.CreateObject("WIA.ImageFile")
		oImage.LoadFile sFile

		set oProcess = Server.CreateObject("WIA.ImageProcess")

		oProcess.Filters.Add "{4EBB0166-C18B-4065-9332-109015741711}" 'Scale
		oProcess.Filters(1).Properties(1).Value = 500
		oProcess.Filters(1).Properties(2).Value = 500

		If ((oImage.FormatID <> "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}") And _
			(oImage.FormatID <> "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}")) Then
			oProcess.Filters.Add "{42A6E907-1D2F-4b38-AC50-31ADBE2AB3C2}" 'Convert
			oProcess.Filters(2).Properties(1).Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" 'JPEG
			oProcess.Filters(2).Properties(2).Value = 100 'JPEG Compression
		End If			
			
		set oImage = oProcess.Apply(oImage)
		if Not oImage is Nothing then
			set oVector = oImage.FileData
			Response.BinaryWrite oVector.BinaryData
		End if
	elseif sCmd = "Thumb" then
		sFile = Request.QueryString("F")
		set oImage = Server.CreateObject("WIA.ImageFile")
		oImage.LoadFile sFile

		set oProcess = Server.CreateObject("WIA.ImageProcess")

		oProcess.Filters.Add "{4EBB0166-C18B-4065-9332-109015741711}" 'Scale
		oProcess.Filters(1).Properties(1).Value = 100
		oProcess.Filters(1).Properties(2).Value = 100

		If ((oImage.FormatID <> "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}") And _
			(oImage.FormatID <> "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}")) Then
			oProcess.Filters.Add "{42A6E907-1D2F-4b38-AC50-31ADBE2AB3C2}" 'Convert
			oProcess.Filters(2).Properties(1).Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" 'JPEG
			oProcess.Filters(2).Properties(2).Value = 100 'JPEG Compression
		End If			
			
		set oImage = oProcess.Apply(oImage)

		if Not oImage is Nothing then
			set oVector = oImage.FileData
			Response.BinaryWrite oVector.BinaryData
		End if
	elseif sCmd = "Details" then
%>
<HTML>
<HEAD>
<TITLE>ImageFile Demo</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<table width="100%"><tr><td width=40% valign=top>
<%	
		sFile = Request.QueryString("F")
		set oImage = Server.CreateObject("WIA.ImageFile")
		oImage.LoadFile sFile
		
		Response.Write "<img src=""" & Request.ServerVariables("SCRIPT_NAME") & "?C=BigThumb&F=" & Server.URLEncode(sFile) & """>"
		Response.Write "</td><td width=""60%"" valign=top><table width=""100%"">"

		for each oProperty in oImage.Properties
			if oProperty.IsVector then
				s = ""
				set oVector = oProperty.Value
				for i = 1 to oVector.Count
					if IsObject(oVector(i)) Then
						set oRational = oVector(i)
						s = s & oRational.Numerator & "/" & oRational.Denominator
					else
						s = s & oVector(i)
					end if
					if i <> oVector.Count then
						s = s & ", "
					end if
				next
			else
				if IsObject(oProperty.Value) Then
					set oRational = oProperty.Value
					s = oRational.Numerator & "/" & oRational.Denominator
				else
					s = CStr(oProperty.Value)
				end if
			end if
			Response.Write "<tr><td width=""20%"">" & oProperty.Name & "</td><td width=""80%"">" & s & "</td></tr>"
		next
%>		
</td></tr></table>
</BODY>
</HTML>
<%		
	elseif sCmd = "Dir" then
%>
<HTML>
<HEAD>
<TITLE>ImageFile Demo</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%	
		set oImage = Server.CreateObject("WIA.ImageFile")

		sDir = Request.QueryString("D")
		Response.Write "<H1>Directory Listing for " & sDir & "</H1>"

		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set fldr = fso.GetFolder("c:\inetpub\wwwroot\pictures" & sDir)
		set fldrs = fldr.Subfolders
		for each f in fldrs
			if left(f.name, 1) <> "_" then
				Response.Write "<A HREF=""" & Request.ServerVariables("SCRIPT_NAME") & "?C=Dir&D=" & Server.URLEncode(sDir & f.Name & "\") & """>" & f.Name & "</A><br>"
			end if
		next
		
		set files = fldr.Files

		bGotPics = false
		for each f in files
			on error resume next
			oImage.LoadFile f.Path
			if err.number = 0 then
				bGotPics = true
				exit for
			else
				err.clear
			end if
		next
				
		if bGotPics then
%>
<table width="100%"><tr>
<%		
			i = 0
			for each f in files
				on error resume next
				oImage.LoadFile f.Path
				if err.number = 0 then
					if i mod 5 = 0 then
						Response.Write "</tr><tr>"
					end if
					Response.Write "<td>"
					Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?C=Image&F=" & Server.URLEncode(f.path) & """>"
					Response.Write "<img src=""" & Request.ServerVariables("SCRIPT_NAME") & "?C=Thumb&F=" & Server.URLEncode(f.path) & """>"
					Response.Write "</a>"
					Response.Write "<br>"
					Response.Write "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?C=Details&F=" & Server.URLEncode(f.path) & """>"
					Response.Write f.Name & " - Details"
					Response.Write "</a>"
					Response.Write "</td>"
					i = i + 1
				else
					err.clear
				end if
			next
%>		
</tr></table>
<%
		end if
%>
</BODY>
</HTML>
<%		
	else
		Response.Redirect Request.ServerVariables("SCRIPT_NAME") & "?C=Dir&D=\"
	end if
%>
