
Class MpfLight
    Public Name
    Public Number
    Public SubType
End Class


Sub ExtractLightsSection(filePath)
    Dim fso, file, line, inLightsSection, currentIndent
    Dim lights(), lightCount
    Dim currentLight : Set currentLight = Nothing
    lightCount = 0
    inLightsSection = False
    currentIndent = 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading

        Do Until file.AtEndOfStream
            line = file.ReadLine

            ' Check for entering the lights section
            If Trim(line) = "lights:" Then
                inLightsSection = True
                ReDim lights(0)
            ElseIf inLightsSection And Trim(line) = "" Then
                ' Exiting the lights section when encountering an empty line
                inLightsSection = False
            ElseIf inLightsSection Then
                Dim currentLineIndent
                currentLineIndent = Len(line) - Len(LTrim(line))

                ' Start of a new light entry
                If currentLineIndent = 2 Then
                    If Not currentLight Is Nothing Then
                        ' Save the previous light if exists
                        If lightCount > 0 Then ReDim Preserve lights(lightCount)
                        Set lights(lightCount) = currentLight
                        lightCount = lightCount + 1
                    End If
                    Set currentLight = New MpfLight
                    currentLight.Name = Trim(Split(line, ":")(0))
                ElseIf Not currentLight Is Nothing And currentLineIndent = 4 Then
                    ' Properties of the light
                    Dim propertyName, propertyValue
                    propertyName = LTrim(Split(line, ":")(0))
                    propertyValue = Trim(Split(line, ":")(1))
                    
                    Select Case LCase(propertyName)
                        Case "number"
                            currentLight.Number = propertyValue
                        Case "subtype"
                            currentLight.SubType = propertyValue
                    End Select
                End If
            End If
        Loop

        ' Add the last light if not added
        If Not currentLight Is Nothing Then
            If lightCount > 0 Then ReDim Preserve lights(lightCount)
            Set lights(lightCount) = currentLight
            lightCount = lightCount + 1
        End If

        file.Close
    Else
        msgbox "File not found: " & filePath
        Exit Sub
    End If

    ' Output the results
    Dim i
    Dim mpfUpdateLamps
    mpfUpdateLamps = "Sub MPFUpdateLamps(changedLamp, brightness)" & vbCrLf
    mpfUpdateLamps = mpfUpdateLamps & "  Select Case changedLamp" & vbCrLf
    For i = 0 To UBound(lights)
        If lights(i).SubType = "led" Then
            mpfUpdateLamps = mpfUpdateLamps & "    Case """&lights(i).Number&"-r""" & vbCrLf
            mpfUpdateLamps = mpfUpdateLamps & "      lightCtrl.LightColor "&lights(i).Name&", ChangeColorChannel(lightCtrl.GetLightColor("&lights(i).Name&"), ""red"", (brightness*255))" & vbCrLf
            mpfUpdateLamps = mpfUpdateLamps & "    Case """&lights(i).Number&"-g""" & vbCrLf
            mpfUpdateLamps = mpfUpdateLamps & "      lightCtrl.LightColor "&lights(i).Name&", ChangeColorChannel(lightCtrl.GetLightColor("&lights(i).Name&"), ""green"", (brightness*255))" & vbCrLf
            mpfUpdateLamps = mpfUpdateLamps & "    Case """&lights(i).Number&"-b""" & vbCrLf
            mpfUpdateLamps = mpfUpdateLamps & "      lightCtrl.LightColor "&lights(i).Name&", ChangeColorChannel(lightCtrl.GetLightColor("&lights(i).Name&"), ""blue"", (brightness*255))" & vbCrLf
        End If
        'mpfUpdateLamps = mpfUpdateLamps & "    Case """&lights(i).Number&"""" & vbCrLf
        'mpfUpdateLamps = mpfUpdateLamps & "      "&lights(i).Name&".State=brightness" & vbCrLf
        'msgbox "Name: " & lights(i).Name & ", Number: " & lights(i).Number & ", SubType: " & lights(i).SubType
    Next
    mpfUpdateLamps = mpfUpdateLamps & "  End Select" & vbCrLf
    mpfUpdateLamps = mpfUpdateLamps & "End Sub"
    'MsgBox mpfUpdateLamps
    ExecuteGlobal mpfUpdateLamps
    'MsgBox mpfUpdateLamps
End Sub


' Call the subroutine
ExtractLightsSection("mpf/config/config.yaml")



Function ChangeColorChannel(currentColor, channel, newValue)
    Dim red, green, blue

    ' Extract RGB components
    'debug.print(currentColor)
    currentColor = clng(currentColor)
    red = currentColor Mod 256
    green = (currentColor \ 256) Mod 256
    blue = (currentColor \ (256 * 256)) Mod 256

    ' Update the specified channel
    Select Case LCase(channel)
        Case "red"
            red = newValue
        Case "green"
            green = newValue
        Case "blue"
            blue = newValue
    End Select

    ' Reconstruct the RGB value
    'debug.print("red:" &red & ", green: "& green & ", blue:" & blue)
    ChangeColorChannel = RGB(red,green,blue)
End Function
