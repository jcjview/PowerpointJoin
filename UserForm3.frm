VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "PowerpointJoin"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim fd(10000) As Object
 Dim count As Integer
 Dim filenum&
 Dim myfile(10000) As Object
 Dim opencount&
 Dim thisone As Presentation

 
 Sub SlideCopy(ByVal SourcePPT As Presentation)

      ' Variable declarations.
      Dim SourceView, answer As Integer
      Dim SourceSlides, NumPres, x As Long

      ' Count the open presentations.
      NumPres = Presentations.count

     

     If SourcePPT.Name = thisone.Name Then
        MsgBox "error"
     End If
        
      ' Stores the current view of the source presentation.
      SourceView = SourcePPT.Windows(1).ViewType

      ' Count the number of slides in source presentation.
      SourceSlides = SourcePPT.Slides.count
'      thisone.Windows(1).ViewType = ppViewSlide
'      SourcePPT.Windows(1).ViewType = ppViewSlide
      ' Loop through all the slides and copy them to destination one by
      ' one.
      For x = 1 To SourceSlides
         ' Select the first slide in the presentation and copy it.
'         SourcePPT.Windows(1).Activate
         SourcePPT.Slides(x).Copy
         ' Switch to destination presentation.
         
         thisone.Slides.Paste Index:=thisone.Slides.count + 1
'         thisone.Windows(1).Activate
         
'         thisone.Windows(1).View.GotoSlide Index:=ActivePresentation.Slides.count
'         thisone.Windows(1).View.Paste
         
      Next x

      ' Restore the current view to source.
      ActiveWindow.ViewType = SourceView
      thisone.Windows(1).ViewType = SourceView
      thisone.Windows(1).Activate

   End Sub
 
 
Private Sub CommandButton1_Click()
  Dim fdo As Object
  Dim ft1, ft2, fs
  Dim f1, f2
  Dim count&
  Dim ppt As Presentation
  
  filenum = 0
  Set thisone = Presentations(1)
  For Each ppt In Presentations
    If InStr(ppt.Name, "PowerPointJoin") Or InStr(ppt.Name, "RunAllInOne_plus") Then
        Set thisone = ppt
    End If
  Next ppt
'  MsgBox thisone.Name
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(TextBox1.Text) = False Then
      MsgBox "文件夹不存在"
     Exit Sub
  End If
  Set fdo = fso.GetFolder(TextBox1.Text)
   
  Set fd(1) = fdo
  count = 1
  Do While count > 0
        Set ft1 = fd(count).Files  '文件
        Set fs = fd(count).SubFolders '文件夹
        count = count - 1
        If ft1.count <> 0 Then
            For Each f1 In ft1
                If InStr(f1.Name, ".ppt") Or InStr(f1.Name, ".pptx") Then
                    filenum = filenum + 1
                    Set myfile(filenum) = f1
                End If
            Next
        End If
    '******************这段是递归到子文件夹**********
    '        If fs.count <> 0 Then
    '            For Each f2 In fs
    '                  count = count + 1
    '                  Set fd(count) = f2
    '            Next
    '              End If
    '******************这段是递归到子文件夹**********
  Loop
  
  
  Label1.Caption = filenum
  Label3.Caption = thisone.Name
  opencount = 0
Set dstSlides = thisone.Slides
Dim x As Long
For x = dstSlides.count To 1 Step -1
    dstSlides(x).Delete
Next x

Set Slide = dstSlides.Add(1, ppLayoutBlank)

End Sub

Private Sub CommandButton2_Click()
    Dim tp As Object
    Dim ppt As Object
'  Application.Visible = False
  opencount = opencount + 1
  
  If opencount > filenum Or opencount = 0 Then
    MsgBox "错误"
    Exit Sub
  End If
  
'  For Each tp In Workbooks
'            If opencount > 1 Then
'                If tp.Name = myfile(opencount - 1).Name Then tp.Close savechanges:=False
'
'            ElseIf tp.Name = myfile(opencount).Name Then
'                MsgBox "已经有名为" & myfile(opencount).Name & "相同文件打开，是否打开？"
'            End If
'  Next tp
Set dstSlides = thisone.Slides

    
For opencount = 1 To filenum
  On Error GoTo openerror
  If (myfile(opencount).Name = thisone.Name) Then GoTo nextfor
  
  Set ppt = Presentations.Open(FileName:=myfile(opencount).Path)
  mySlidesCount = ppt.Slides.count
  ppt.Close
  dstSlides.InsertFromFile myfile(opencount).Path, dstSlides.count, 1, mySlidesCount
nextfor:
Next opencount


  On Error GoTo 0
'  Application.Visible = True
  Exit Sub
  
openerror:
  MsgBox "打开错误 " & myfile(opencount).Name
  Resume Next
  
End Sub

Private Sub CommandButton3_Click()

     Set fdtemp = Application.FileDialog(msoFileDialogFolderPicker)
    If fdtemp.Show Then
        TextBox1.Text = fdtemp.SelectedItems(1)
        Call CommandButton1_Click
    End If
End Sub



Private Sub UserForm_Initialize()
Set thisone = Presentations(1)
  For Each ppt In Presentations
    If InStr(ppt.Name, "PowerPointJoin") Or InStr(ppt.Name, "RunAllInOne_plus") Then
        Set thisone = ppt
    End If
  Next ppt
TextBox1.Text = thisone.Path
End Sub
