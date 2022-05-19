VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagination"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbNums 
      Height          =   315
      ItemData        =   "Form7.frx":0000
      Left            =   240
      List            =   "Form7.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resigstros: 0 - 0 / 0"
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   2850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   536
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Página"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' - Random data
Private Const FIRST_NAMES = "Alan,Alison,Alfie,Andrew,Bill,Dawn,Hannah,Brian,David,Jane,Jennifer,Karen,Gavin,Kerry,Kim,James,Laura,Lucy,Michael,Mellisa,Paula,Patrick,Sarah,Emily,Susan"
Private Const LAST_NAMES = "Anderson,Allen,Black,Evans,Bloggs,Brown,Clarke,Cole,Davis,Dawson,Brown,Gate,Johnson,Jones,Lawson,Lee,Richards,Ryan,Smith,Stephens,Temple,Turner,Wallace,White,Williams"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Private dbc  As SQLiteConnection


' PAGINATE
'----------------------------
Private lCount      As Long
Private lPage       As Long
Private lPageCount  As Long
Private lPerPage    As Long


Private Sub Form_Load()
Dim Sql As String
Dim i As Long

    Randomize Timer '[Random data]
    
    Set dbc = SQLite.Connection(":memory:")
    dbc.Errmode = ErrShow
   
    Sql = "CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, first_name TEXT, last_name TEXT, age INTEGER);"
    dbc.Execute Sql
    
    dbc.Transaction = BEGIN
    For i = 1 To 80
        Sql = "INSERT INTO users (first_name,last_name,age) VALUES ('" & RandomName & "','" & RandomName2 & "'," & RandomNum(60) & ");"
        If dbc.Execute(Sql) <> SQLITE_OK Then Exit For
    Next
    If dbc.errCode = 0 Then dbc.Transaction = COMMIT Else dbc.Transaction = ROLLBACK
    
    
    'OBTENER CANTIDAD REGISTROS
    '----------------------------
    lCount = Val(dbc.ResultString("SELECT COUNT(*) FROM users;"))
    cbNums.ListIndex = 1
    
End Sub
Private Sub cbNums_Click()
Dim i As Long

    List1.Clear
    List2.Clear

    lPerPage = Val(cbNums)
    lPageCount = (lCount \ lPerPage)
    If lPerPage * lPageCount < lCount Then lPageCount = lPageCount + 1
    
    For i = 1 To lPageCount
        List1.AddItem "Página " & i
    Next
    
    If List1.ListCount Then List1.ListIndex = 0
    
End Sub


Private Sub List1_Click()
Dim lOfset  As Long

    ' ESTE METODO DE PAGINACION ES LENTO SEGUN LA WIKI DE SQLITE
    '-----------------------------------------------------------------------------------------
    ' https://www2.sqlite.org/cvstrac/wiki?p=ScrollingCursor

    lOfset = List1.ListIndex * lPerPage
    List2.Clear
    
    With dbc.Query("SELECT * FROM users LIMIT " & lPerPage & " OFFSET " & lOfset & ";")
        Do While .Step = SQLITE_ROW
            List2.AddItem !last_name & ", " & !first_Name
        Loop
    End With
    
    Label3 = "" & (lOfset + 1) & " - " & (lOfset) + List2.ListCount & " / " & lCount
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set dbc = Nothing
End Sub







' Get Random Data
'---------------------------------------------------------------------
Private Function RandomName() As String
    RandomName = Split(FIRST_NAMES, ",")(RandomNum(25))
End Function
Private Function RandomName2() As String
    RandomName2 = Split(LAST_NAMES, ",")(RandomNum(25))
End Function
Private Function RandomNum(Max As Long)
    RandomNum = Int(Max * Rnd)
End Function



