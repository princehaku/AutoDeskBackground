VERSION 5.00
Begin VB.Form config_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoDB"
   ClientHeight    =   1425
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动停用"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "运行全屏程序时自动停用"
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Text            =   ".\Screen"
      Top             =   180
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Text            =   "15"
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "壁纸文件夹"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "更换频率"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "分钟"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "config_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Open App.Path + "/config.ini" For Output As #1
        Write #1, Text2.Text
        Write #1, Text1.Text
        Write #1, Check1.Value
     Close #1
     a = MsgBox("设置保存成功", vbOKOnly, "success")
      Shell ("core.exe")
     End
End Sub

Private Sub Form_Load()
    If (PathFileExists(App.Path + "/config.ini") <> 1) Then
      MsgBox ("此程序献给我爱的A3~~...by princehaku")
    End If
    Text2.Text = App.Path + "\Screen"
End Sub

Private Sub Text2_Click()
    Dim Shell, myPath
     Set Shell = CreateObject("Shell.Application")
     Set myPath = Shell.BrowseForFolder(&O0, "请选择壁纸所在文件夹", &H1 + &H10, "")
     If Not myPath Is Nothing Then Text2.Text = myPath.Items.Item.Path
     Set Shell = Nothing
     Set myPath = Nothing
End Sub
