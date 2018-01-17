VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VIEW"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "ITEM DETAILS"
      Height          =   4935
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   5175
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "LOCATION :"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "UNITS :"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "RATE:"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "ITEM NAME :"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "COMPANY NAME :"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "ITEM NO. :"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Caption         =   "INCOMPLETE FIELDS !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "THIS ITEM NO. IS NOT UNIQUE !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1920
      TabIndex        =   19
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "INVENTORY MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Combo2.Clear
If Combo1.Text = "Apsara" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "eraser"
Combo2.AddItem "sharpner"
Combo2.AddItem "scale"
ElseIf Combo1.Text = "Natraj" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "eraser"
Combo2.AddItem "sharpner"
Combo2.AddItem "scale"
Combo2.AddItem "marker"
ElseIf Combo1.Text = "Claro" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "scale"
ElseIf Combo1.Text = "Win" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "marker"
ElseIf Combo1.Text = "Classmate" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "geometry box"
Combo2.AddItem "copy"
Combo2.AddItem "scale"
ElseIf Combo1.Text = "FLI" Then
Combo2.AddItem "pen"
Combo2.AddItem "scale"
ElseIf Combo1.Text = "Local" Then
Combo2.AddItem "pencil"
Combo2.AddItem "pen"
Combo2.AddItem "eraser"
Combo2.AddItem "sharpner"
Combo2.AddItem "scale"
Combo2.AddItem "marker"
Combo2.AddItem "copy"
Combo2.AddItem "chalk"
End If


End Sub

Private Sub Command1_Click()
Label9.Visible = False

If flag = 0 Then
 Label8.Visible = False
 
 End If

Text1.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""


End Sub

Private Sub Command2_Click()
Dim flag As Integer
flag = 0
Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
 If Text1.Text = Data1.Recordset.Fields(0) Then
 flag = 1
 End If
 Data1.Recordset.MoveNext
 Loop
 If flag = 1 Then
 Label8.Visible = True
 
 

 
 
 

 End If
 
 
 If Text1.Text = "" Then
  Label9.Visible = True
 ElseIf Text2.Text = "" Then
 Label9.Visible = True
 ElseIf Text2.Text = "" Then
 Label9.Visible = True
 ElseIf Text3.Text = "" Then
 Label9.Visible = True
 ElseIf Text4.Text = "" Then
 Label9.Visible = True
 ElseIf Combo1.Text = "" Then
 Label9.Visible = True
 ElseIf Combo2.Text = "" Then
 Label9.Visible = True
 
 End If
 
 
 
 
 
 

Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Val(Text1.Text)
Data1.Recordset.Fields(1) = Combo1.Text
Data1.Recordset.Fields(2) = Combo2.Text
Data1.Recordset.Fields(3) = Val(Text2.Text)
Data1.Recordset.Fields(4) = Val(Text3.Text)
Data1.Recordset.Fields(5) = Text4.Text
Data1.Recordset.Update
Text1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False

Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
 If Data1.Recordset.Fields(0) = "0" Then
 Data1.Recordset.Delete
 End If
 Data1.Recordset.MoveNext
 Loop
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me

End Sub

Private Sub Command4_Click()
Form3.Show
Unload Me

End Sub

Private Sub Command5_Click()
Form4.Show
Unload Me

End Sub

Private Sub Form_Load()
Data1.Visible = False






If flag = 0 Then
 Label8.Visible = False
 Label9.Visible = False
 
 
 End If


Data1.DatabaseName = "C:\Users\user\Desktop\project new\ABCD.mdb"
Data1.RecordSource = "INVENTORYSTOCK"
Data1.Refresh
Text1.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.AddItem "Apsara"
Combo1.AddItem "Natraj"
Combo1.AddItem "Claro"
Combo1.AddItem "Win"
Combo1.AddItem "Classmate"
Combo1.AddItem "FLI"
Combo1.AddItem "Local"
End Sub

