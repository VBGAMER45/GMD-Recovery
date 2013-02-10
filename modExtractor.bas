Attribute VB_Name = "modExtractor"
'###########################################
'modExtractor.bas
'vbgamer45
'###########################################
Dim Filecount As Integer
Dim WavCount As Integer
Global Finalkey43 As String * 256
Global Finalkey50 As String * 256
Global Finalkey51 As String * 256
Global Finalkey52 As String * 256
Global Finalkey53Beta As String * 256
Global Finalkey53A As String * 256
Global Finalkey54 As String * 256
Global Finalkey60 As String * 256
Sub GetBitmap(Filename As String, Offset As Long)
 Dim Height As Long
 Dim Width As Long
 Dim EndofBmp As Long
 Dim bmpHeader As String * 55
 'BM6 VERSION
 Dim f As Long
 f = 19

Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset
   Get f, Offset, bmpHeader
   Seek f, Offset + 18
   'get height
   Get f, Offset + 18, Height
   
   'Get width
   Get f, Offset + 22, Width
  If Height > 2000 Then Exit Sub
  If Width > 2000 Then Exit Sub
  If Height < 0 Then Exit Sub
  If Width < 0 Then Exit Sub
   'Calcute end of bmp
   EndofBmp = 55 + (Height * Width * 3)

  
   

   
Dim k As String * 1
Dim i As Long
g = FreeFile
    Open App.Path & "\dump\images\img" & Filecount & ".bmp" For Binary Access Write Lock Write As g
         Put g, , bmpHeader
         Seek f, Offset + 55
         For i = 55 To EndofBmp
            Get f, Offset + i, k
            Put g, , k
         Next i
    Close g
  Close f

'Keep track bmp nubmer
Filecount = Filecount + 1

End Sub
Sub GetWav(Filename As String, Offset As Long)

 Dim Size As Long

 'RIFF VERSION
 Dim f As Long
 f = 20

Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset
   'get size
   Get f, Offset + 4, Size

Size = Size + 7

  If Size > 1000000 Then Exit Sub
  If Size < 0 Then Exit Sub
   'Calcute end of bmp

Dim k As String * 1
Dim i As Long
g = FreeFile
    KillOldWavs (App.Path & "\dump\sounds\wav" & WavCount & ".wav")
    Open App.Path & "\dump\sounds\wav" & WavCount & ".wav" For Binary Access Write Lock Write As g

         For i = 0 To Size
            Get f, Offset + i, k
            Put g, , k
         Next i
    Close g
  Close f

'Keep track wav nubmer
WavCount = WavCount + 1

End Sub
Sub Get52Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   'Get f, Offset + 40, modExtractor.Finalkey52
   Get f, Offset + 56, modExtractor.Finalkey52
 Close f
 
  Open App.Path & "\final52.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey52
   
 Close f
End Sub

Sub Get51Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

  ' Get f, Offset + 40, modExtractor.Finalkey51
   Get f, Offset + 56, modExtractor.Finalkey51
 Close f
 
  Open App.Path & "\final51.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey51
   
 Close f
End Sub
Sub Get43Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   'Get f, Offset + 124, modExtractor.Finalkey50
   Get f, Offset + 120, modExtractor.Finalkey43
   
 Close f
 
  Open App.Path & "\final43.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey43
   
 Close f
End Sub
Sub Get50Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   'Get f, Offset + 124, modExtractor.Finalkey50
   Get f, Offset + 204, modExtractor.Finalkey50
   
 Close f
 
  Open App.Path & "\final50.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey50
   
 Close f
End Sub
Sub Get53BetaKey(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   'Get f, Offset + 40, modExtractor.Finalkey53Beta
   Get f, Offset + 56, modExtractor.Finalkey53Beta
 Close f
 
  Open App.Path & "\final53Beta.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey53Beta
   
 Close f
End Sub
Sub Get53AKey(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   'Get f, Offset + 40, modExtractor.Finalkey53Beta
   Get f, Offset + 56, modExtractor.Finalkey53A
 Close f
 
  Open App.Path & "\final53A.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey53A
   
 Close f
End Sub
Sub Get54Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset

   Get f, Offset + 56, modExtractor.Finalkey54
 Close f
 
  Open App.Path & "\final54.txt" For Binary Access Write Lock Write As f

   

   Put f, , modExtractor.Finalkey54
   
 Close f
End Sub
Sub Get60Key(Filename As String, Offset As Long)
 Dim f As Long
 f = 20


Close
 Open Filename For Binary Access Read Lock Read As f

   
   Seek f, Offset
    'Get F, Offset + 41, modExtractor.Finalkey60
   Get f, Offset + 41, modExtractor.Finalkey60
 Close f
 
  Open App.Path & "\final60.txt" For Binary Access Write Lock Write As f

   Put f, , modExtractor.Finalkey60
   
 Close f
End Sub
Sub KillOldWavs(Filename As String)
On Error Resume Next
Kill (App.Path & "\dump\sounds\wav" & WavCount & ".wav")
End Sub
Sub ExtractGmd(Offset As Long)
'On Error GoTo nofile:
Dim Temp As Byte
Open App.Path & "\temp\game.enc" For Binary Access Read Lock Read As #2
            Filesize = LOF(2)
            'Now Decrypt the file
            Open App.Path & "\temp\game.gmd" For Binary Access Write Lock Write As #3
                Seek #2, Offset
                Do Until EOF(2)
                 
                    Get #2, , Temp

                 
                 
                 If LOF(2) <> Loc(2) - 1 Then
                    Put #3, , Temp
                 End If
             Loop
             Dim d As String
             d = vbCrLf & "This gmd was made from the GameMaker 5.x Decompiler" & vbCrLf
             Put #3, , d
             d = "It was decompiled on : " & Date & vbCrLf
             Put #3, , d1
             d = "By: " & Environ("USERNAME") & vbCrLf
             Put #3, , d
             d = "Version of GmDecompiler=" & Version & vbCrLf
             Put #3, , d
             d = "To stop decompilers get the antidecompiler at gmupload"
             Put #3, , d
             
            Close #3
        Close #2
'Exit Sub
'nofile:
  '  MsgBox err.Description
'Exit Sub
End Sub
