VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00400000&
   Caption         =   "Form2"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List6 
      Height          =   3570
      Left            =   8400
      TabIndex        =   6
      Top             =   1560
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
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "home"
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Height          =   3570
      Left            =   6720
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   3570
      Left            =   5040
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   3570
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "INVENTORY STOCK DATABASE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label6 
      Caption         =   "LOCATION:"
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "UNITS:"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "RATE:"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ITEM NAME:"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "COMPANY NAME:"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "ITEM NO.:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form1.Show
Unload Me

End Sub

Private Sub Form_Load()
Data1.Visible = False
Data1.DatabaseName = "C:\Users\user\Desktop\project new\ABCD.mdb"
Data1.RecordSource = "INVENTORYSTOCK"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While (Not Data1.Recordset.EOF)
List1.AddItem (Data1.Recordset.Fields(0))
List2.AddItem (Data1.Recordset.Fields(1))
List3.AddItem (Data1.Recordset.Fields(2))
List4.AddItem (Data1.Recordset.Fields(3))
List5.AddItem (Data1.Recordset.Fields(4))
List6.AddItem (Data1.Recordset.Fields(5))
Data1.Recordset.MoveNext

Loop


End Sub
