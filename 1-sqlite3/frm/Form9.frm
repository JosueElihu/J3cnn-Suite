VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extensiones"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text5 
      Height          =   345
      Left            =   4320
      TabIndex        =   10
      Top             =   3840
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   4320
      TabIndex        =   8
      Top             =   3120
      Width           =   4575
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   4320
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   264
      X2              =   264
      Y1              =   24
      Y2              =   280
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SHA1"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   3600
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MD5"
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   2880
      Width           =   315
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   288
      X2              =   592
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   288
      X2              =   592
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CSV Import"
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
      TabIndex        =   7
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unroot13"
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Root13"
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Funcion personalizada:"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbc As SQLiteConnection
Dim sFile As String
    
    Set dbc = SQLite.Connection(":memory:")
    dbc.Errmode = ErrShow

    'HABILITAR LA CARGA DE EXTENSIONES
    '----------------------------------------------------------------------------------
    dbc.Extensions = True
    
    'CARGA DE EXTENSIONES: https://www.sqlite.org/loadext.html
    '----------------------------------------------------------------------------------
    ' sPath             : Ruta de la libreria o dll
    ' InitFunctionName  : Punto de inicio de la libreria(Opcional)
    '----------------------------------------------------------------------------------
    
 
    dbc.LoadExt App.Path & "\ext\example.dll", "sqlite3_example_init"   ' - EXTENSION Ejemplo
    dbc.LoadExt App.Path & "\ext\csv.dll", "sqlite3_csv_init"           ' - EXTENSION Csv
    dbc.LoadExt App.Path & "\ext\root13.dll", "sqlite3_rot_init"        ' - EXTENSION Root13
    
    dbc.LoadExt App.Path & "\ext\md5.dll"           ' - EXTENSION Md5
    dbc.LoadExt App.Path & "\ext\sha1.dll"          ' - EXTENSION Sha1
    

    'Ejemplo funcion personalizada: (example.dll)
    '----------------------------------------------------------------------------------
    Text1 = dbc.ResultString("SELECT myCustomFunc('Hello world');")
    
    'Root13
    '----------------------------------------------------------------------------------
    Text2 = dbc.ResultString("SELECT rot13('Hello world');")
    Text3 = dbc.ResultString("SELECT rot13(rot13('Hello world'));")
    
    'MD5 - SHA1
    '----------------------------------------------------------------------------------
    Text4 = dbc.ResultString("SELECT hex(md5('Hello world'));")
    Text5 = dbc.ResultString("SELECT sha1('Hello world');")
    
    
    'CSV: (Importar datos)
    '----------------------------------------------------------------------------------
    sFile = App.Path & "\db\personal.csv"
    dbc.Execute "CREATE VIRTUAL TABLE temp.csv USING csv(filename='" & sFile & "');"
    
    'CSV: (Leer datos)
    With dbc.Query("SELECT * FROM csv;")
        Do While .Step = SQLITE_ROW
            List1.AddItem .Value(0) & " - " & .Value(1)
        Loop
    End With
    
    Set dbc = Nothing
    
    '----------------------------------------------------------------------------------
    ' Las extensiones deben compilarse en una dll: https://www.sqlite.org/loadext.html
    '----------------------------------------------------------------------------------
    
End Sub

