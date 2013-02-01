VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "ADB_core.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2010
      Hidden          =   -1  'True
      Left            =   720
      Pattern         =   "*.jpg;*.bmp;*.gif;*.png"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBdir$, mins, autoclose, timerId, a, file$

    
Private Sub Form_Load()
Randomize
    If (App.PrevInstance = True) Then End
    
    If (PathFileExists(App.Path + "/config.ini")) Then
         Open App.Path + "/config.ini" For Input As #1
         Input #1, DBdir$
         Input #1, mins
         Input #1, autoclose
         Close #1
    End If
    
    File1.Path = DBdir$
    mins = Int(mins)
    Call sc
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   timerId = timerId + 1
   
    If (PathFileExists(App.Path + "/config.ini")) Then
         Open App.Path + "/config.ini" For Input As #1
         Input #1, DBdir$
         Input #1, mins
         Input #1, autoclose
         Close #1
    End If
    File1.Path = DBdir$
    mins = Int(mins)
 '如果有全屏3d程序运行..不切换壁纸
   
 '切换壁纸
  If (timerId >= mins * 60) Then
     
    Call sc
    
    timerId = 0
    
  End If
End Sub

Public Function sc()
    
    file$ = DBdir + "\" + File1.List(Int(File1.ListCount * Rnd()))

    Dim Path     As String, strSave       As String
    strSave = String(50, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave)))
    If PathFileExists(Path & "\IloveA3.bmp") = 0 Then
    '设定背景
        Open App.Path & "/a.bat" For Output As #2
        Print #2, "set regadd=reg add " & Chr(34) & "HKEY_CURRENT_USER\Control Panel\Desktop" & Chr(34)
        Print #2, "%regadd% /v TileWallpaper /d 0 /f"
        Print #2, "%regadd% /v Wallpaper /d " & Chr(34) & Path & "\IloveA3.bmp" & Chr(34) & " /f "
        Print #2, "%regadd% /v WallpaperStyle /d 1 /f"
        Print #2, "RunDll32.exe USER32.DLL,UpdatePerUserSystemParameters"
        Close #2
        a = Shell(App.Path & "/a.bat", vbMinimizedNoFocus)
    End If
    '转换图片并保存到Windows目录下面
    SavePicture LoadPicture(file$), Path & "\IloveA3.bmp"
    
    a = SystemParametersInfo(20, 0, Path & "\IloveA3.bmp", 1)
End Function
