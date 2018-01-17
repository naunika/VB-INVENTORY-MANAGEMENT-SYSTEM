VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00400000&
   Caption         =   "Form4"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form4"
   ScaleHeight     =   4650
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "VIEW"
      Height          =   615
      Left            =   6600
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
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
   Begin VB.CommandButton Command2 
      Caption         =   "home"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "delete"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "DELETED SUCCESSFULLY !"
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
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "enter the item no. you want to delete :"
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
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
 If Data1.Recordset.Fields(0) = Text1.Text Then
 Data1.Recordset.Delete
 End If
 Data1.Recordset.MoveNext
 Loop
 Text1.Enabled = False
 
Label2.Visible = True

 

End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me

End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Data1.Visible = False
Label2.Visible = False



Data1.DatabaseName = "C:\Users\user\Desktop\project new\ABCD.mdb"
Data1.RecordSource = "INVENTORYSTOCK"
Data1.Refresh
End Sub

