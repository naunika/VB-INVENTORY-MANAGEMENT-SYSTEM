VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00400000&
   Caption         =   "Form3"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840
   LinkTopic       =   "Form3"
   ScaleHeight     =   4710
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "VIEW"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SEARCH"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HOME"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   3960
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      Caption         =   "UPDATED SUCCESSFULLY !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   6480
      TabIndex        =   12
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "LOCATION:"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "UNITS:"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "RATE:"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Enter the item no. you want to update:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
 If Text1.Text = Data1.Recordset.Fields(0) Then
 Data1.Recordset.Edit
 Data1.Recordset.Fields(3) = Text2.Text
 Data1.Recordset.Fields(4) = Text3.Text
 Data1.Recordset.Fields(5) = Text4.Text
 Data1.Recordset.Update
 
 End If
 Data1.Recordset.MoveNext
 Loop
 Label5.Visible = True
 Text1.Enabled = False
 Text2.Enabled = False
 Text3.Enabled = False
 Text4.Enabled = False
 
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
 If Text1.Text = Data1.Recordset.Fields(0) Then
 Text2.Text = Data1.Recordset.Fields(3)
 Text3.Text = Data1.Recordset.Fields(4)
 Text4.Text = Data1.Recordset.Fields(5)
  
 End If
 Data1.Recordset.MoveNext
 
 Loop

 
 
 
End Sub

Private Sub Command4_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Data1.Visible = False
Label5.Visible = False


Data1.DatabaseName = "C:\Users\user\Desktop\project new\ABCD.mdb"
Data1.RecordSource = "INVENTORYSTOCK"
Data1.Refresh

End Sub

