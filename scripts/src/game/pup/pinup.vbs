'***********************************************************************************************************************
'*****  PIN UP                                                 	                                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************

' COPY EVERYTHING BELOW TO THE TOP OF YOUR TABLE SCRIPT UNDER OPTION EXPLICIT

'****** PuP Variables ******

Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!
Dim pBackglass:pBackglass=2

'*************************** PuP Settings for this table ********************************

cPuPPack = "tmntpro"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    
    Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
    If PuPlayer is Nothing Then
        PUPStatus=false
        MsgBox("Could not start Pup")
    Else
        PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
        PUPStatus=true
        InitPupLabels
    End If

End Sub

Sub pupevent(EventNum)
    if PUPStatus=false then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack

'***********************************************************************************************************************
