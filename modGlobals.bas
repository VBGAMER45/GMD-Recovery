Attribute VB_Name = "modGlobals"
'################################################
' modGlobals.bas
' vbgamer45
'################################################
Public Const Funny = "You hacker what the hell are you doing looking in here?"
Public Const GameVer60 = "TRunnerForm"
Public Const GameVer54 = "Game_Maker 5.4"
Public Const GameVer53Beta = "Game_Maker 5.3"
Public Const GameVer53 = "Game_Maker 5.3"
Public Const GameVer52 = "Game_Maker 5.2"
Public Const GameVer51 = "Game_Maker 5.1"
Public Const GameVer50 = "Game_Maker 5.0"
Public Const GameVer43 = "Game_Maker 4.3"

Global VersionDetected As Boolean
Global GameVersion As String
Global ShowBmpOffsets As Boolean
Global ShowWavoffsets As Boolean
Global OverRideEncryptKey As Boolean
Global OverRideNormalKey As Boolean
Global KeyFileName As String
Global NormalFileName As String
Global Version As String
Global KeyFound As Boolean


'Main Data offsets for each game begin position
Public Const OffsetVer50 = 1229825
Public Const OffsetVer51 = 1402369
Public Const OffsetVer52 = 1440029

'Offsets of Encryption key in the exe
Public Const eOffset50 = 1250000 'Offset of reg key too
Public Const eOffset51 = 1450000
Public Const eOffset52 = 1450000
Public Const eOffset53beta = 1500000
Public Const eOffset53 = 1500000
Public Const eOffset53a = 1500000

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
(ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Function GetShortPathName(ByVal strLongPath As String) As String
    Const cchBuffer = 256
    Dim strShortPath As String
    Dim lResult As Long, nPos As Long

    strShortPath = String$(cchBuffer, 0)
    lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
    If lResult = 0 Then
        'Just use the long name as this is usually good enough
        GetShortPathName = strLongPath
    Else

      nPos = InStr(strShortPath, vbNullChar)
      If nPos > 0 Then
          GetShortPathName = Left$(strShortPath, nPos - 1)
      Else
          GetShortPathName = strShortPath
      End If

    End If
End Function
Sub PrintReadMe()
'Just in case people try to pass my program without readme
On Error Resume Next
    Open App.Path & "\Read Me.txt" For Output As #1
        Print #1, "-------------------------------------"
        Print #1, "Game Maker GMD Recovery"
        Print #1, "VisualBasicZone.com"
        Print #1, "Version: " & Version
        Print #1, "For GM Versions 4.3 5.0 5.1 5.2 5.3beta 5.3 and 5.3a"
        Print #1, "Tested on WinXp Home & Professional Edition "
        Print #1, "--------------------------------------"
        Print #1, ""
        Print #1, "Table of Contents"
        Print #1, " 1. Features"
        Print #1, " 2. Using Attach Process"
        Print #1, " 3. Using Open EXE"
        Print #1, " 4. Using Force Decrypt"
        Print #1, " 5. Using Brute Force"
        Print #1, " 6. Using Bmp Extractor"
        Print #1, " 7. Using Wav Extractor"
        Print #1, " 8. Making your own Plugin"
        Print #1, " 9. Help"
        Print #1, " 10. Contact"
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 1. Features: - Is a very powerful program that allows you to peek"
        Print #1, "    Inside a GameMaker exe or any other exe for that matter.  Allows you"
        Print #1, "    you to dump the memory contents of a running process and attempt to "
        Print #1, "    decrypt a gamemaker exe program.  Even comes with a guide to do it step by step"
        Print #1, "    I highly suggest using the guide!. Thats how I decrypt each exe!"
        Print #1, "    Even if you can't get it to decrypt you can still dump the bmps and"
        Print #1, "    the wavs, and scripts. Just make sure you set your memory limit high enough."
        Print #1, "    One step beta decompile has been in since ver .19.  I hope in the future that"
        Print #1, "    I can decode the exe without going in to memory that will make it work for so many more people then."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 2. Using Attach Process"
        Print #1, "    Attach process allows you dump the memory of a certain process."
        Print #1, "    Allows you to set the high and low memory limits of an exe to dump the memory"
        Print #1, "    When you dump the memory it allows my program to find the encryption key. "
        Print #1, "    New and improved! Now contains Application Icon pictures!"
        Print #1, "    The lowerlimit is the base address of the running process."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 3. Using Open Exe"
        Print #1, "    Main function of this is to to get the game.enc. Or the gmd that is encrypted"
        Print #1, "    And decrypt it using the key either found from memory or overrided"
        Print #1, "    from the options panel.  Then once you have the game.enc you just extract the gmd."
        Print #1, "    You should use attach to process before using this function."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 4. Using Force Decrypt"
        Print #1, "    Force Decrypt is a function that allows you decrypt a file according to the gamemaker"
        Print #1, "    encryption version so you can use 5.3a encryption or 5.3 encryption."
        Print #1, "    It is of not much use to most people. I only use it to decode 5.4 exe's."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 5. Using Brute Force"
        Print #1, "    To use the brute force function first you must run your game and"
        Print #1, "    use attach to process or the guide to dump the memory of it."
        Print #1, "    Then the brute force will attempt to find keys in dump file."
        Print #1, "    It may find more than one try them each until you find one that works."
        Print #1, "    Use the key by double clicking the offset in the listbox this loads the encryption key."
        Print #1, "    Then do an open exe and select the same exe to decrypt it."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 6. Using Bmp Extractor"
        Print #1, "    Is a tool that you can use to extract bmps out of memory"
        Print #1, "    And even other exes and files."
        Print #1, "    Just make sure you have the limit high enough."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 7. Using Wav Extractor"
        Print #1, "    Is a tool that you can use to extract wavs out of memory"
        Print #1, "    And even other exes and files."
        Print #1, "    Just make sure you have the limit high enough."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 8. Making your own Plugin"
        Print #1, "    I have included a sample plugin coded in Visual Basic."
        Print #1, "    I will soon try to have a C++ example as well."
        Print #1, "    With plugins you can write your own extractors for other file formats"
        Print #1, "    Some Ideas are PNG GIF JPEG files etc"
        Print #1, "    In order to run your plugin you need to place .dll file in the plugins folder."
        Print #1, ""
        Print #1, "----------------------------------------"
        Print #1, " 9. Help"
        Print #1, "    In this version I have tired to make it easier. "
        Print #1, "    I have included a step by step guide to decrypting"
        Print #1, "    a gamemaker exe into a gamemaker gmd file."
        Print #1, ""
        Print #1, "    Questions:"
        Print #1, ""
        Print #1, "    Why does it say keyfound=false?"
        Print #1, "    Because it cannot find the encryption key"
        Print #1, "    Different versions of windows have different offsets for the key."
        Print #1, "    I suggest you install all the vb runtime files"
        Print #1, ""
        Print #1, "    Why does it make game.enc but not game.gmd?"
        Print #1, "    That means it probably keyfound=false and that the encryption could not be found"
        Print #1, "    Game.enc is the attempted decrypted file. I then try to to extract the gmd from game.enc"
        Print #1, ""
        Print #1, "    It won't extract the gmd!"
        Print #1, "    It should say Finding GMD Start Offset Location"
        Print #1, "    Then check the second line there should be a number if not"
        Print #1, "    then it did not find the encryption key."
        Print #1, ""
        Print #1, "    The program complains about psapi"
        Print #1, "    Means you are running win98 or lower because thoose O/S lack that dll file."
        Print #1, "    I use psapi for two important api functionis for attach process"
        Print #1, ""
        Print #1, "    The program complains about a missing file"
        Print #1, "    First search on google.com for that file"
        Print #1, "    Make sure you have the vb runtime files installed."
        Print #1, "----------------------------------------"
        Print #1, " 10. Contact"
        Print #1, "    You can contact me at gmdecompiler@yahoo.com"
        Print #1, "    I do not decompile apps for you."
        Print #1, "    Only contact for bugs, comments and or suggestions."
        Print #1, "    Source code can be bought please contact me."
        Print #1, "    I have vb6, vb.net, c++, and Asm versions of the code in development."
        Print #1, "    I am currently working on a way to decrypt the .exe's without going into memory."
        Print #1, "    Keep checking back for more things to come."
    Close #1

End Sub
Sub AddLog(sText As String)
    On Error Resume Next
    Dim F As Integer
    F = FreeFile
    Dim tempUser As String
    tempUser = Environ("USERNAME")
    Open App.Path & "\data.log" For Append As F
     Print #F, sText & " by " & tempUser
    Close F
End Sub
