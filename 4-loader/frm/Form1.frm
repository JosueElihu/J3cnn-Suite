VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loader"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   814
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   814
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   12210
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   814
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   12210
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Portable  [No regsvr32, No SxS manifest]"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J3cnnLoader"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1170
      End
   End
   Begin Proyecto1.JGridLite lGrid 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
      BackColor       =   16777215
      HeaderH         =   26
      ForeColor       =   0
      HeaderColor     =   0
      HeaderBack      =   15790320
      BorderColor     =   9471874
      SelColor        =   16737792
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000C3C83&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    '-------------------------------------------------------------------------------------------------------
    ' J3cnnLoader.bas
    '-------------------------------------------------------------------------------------------------------
    ' Para usar J3cnn.dll || J3cnn_c.dll || J3cnn_mc.dll sin registrar la dll:
    '
    ' - Añada el modulo J3cnnLoader.bas
    ' - Copie la dll y sus dependencias a la carpeta o subcarpeta de su proyecto.
    ' - J3cnnLoader.bas buscará y cargará la dll desde la carpeta o subcarpeta de su proyecto, o
    '   puede cargar la dll de forma rapida:# J3cnnLoader.LoadLib App.Path & "\plugins\J3cnn.dll"
    ' - Si utiliza el Loader, utilize el constructor para crear las instancias de conexion:
    '
    '
    '  EJEMPLO 1 (Recomendado):
    '
    '       Dim dbc As SQLiteConnection
    '       Set dbc = SQLite.Connection(":memory:")
    '
    '   Este metodo si aun no se ha cargado la dll hará una busqueda y cargará la dll y creará el
    '   constructor.
    '
    '  EJEMPLO 2:
    '
    '       Dim dbc2 As SQLiteConnection
    '       Call LoadLib("J3cnn.dll")
    '       Set dbc2 = NewObj("SQLiteConnection")
    '
    '   Antes de usar este metodo previamente debe cargar la dll, LoadLib y NewObj son funciones
    '   contenidas en el modulo J3cnnLoader.bas
    '
    '  EJEMPLO 3:
    '
    '       Dim dbc3 As SQLiteConnection
    '       Set dbc3 = NewObj("J3cnn.SQLiteConnection")
    '
    '   En este ejemplo si aun no se ha cargado alguna de las librerias de 'J3cnn', el metodo hara una
    '   busqueda de la dll en la carpeta y subcarpeta del proyecto y lo cargara para crear el objeto.
    '
    '-------------------------------------------------------------------------------------------------------
    ' J3cnnLoader.bas se puede usar para cargar otras DLL Activex, pero para crear una nueva instancia
    ' de los objetos con J3cnnLoader.bas debe de especificar la libreria y el objeto:
    '
    '               Dim Obj As Sound
    '               Set Obj = NewObj("BasFx.Sound")
    '
    '
    ' El siguiente ejemmplo solo es válido para J3cnn.dll, J3cnn_c.dll o J3cnn_mc.dll (Previamente cargados)
    '           Set Obj = NewObj("SQLiteConnection")
    '-------------------------------------------------------------------------------------------------------
    ' Ante cualquier error el constructor de SQLite o el metodo NewObj retorna siempre Nothing
   
    With lGrid
        .AddColumn "id", 80
        .AddColumn "txt", 150
        .AddColumn "Double"
        .AddColumn "Integer"
        .AddColumn "Date", 200
        .AddColumn "Boolean", , vbCenter
    End With
    
    Dim dbc As SQLiteConnection
    Dim T1  As Single
    
    Set dbc = SQLite.Connection(App.Path & "\db\data.db")
    lGrid.NoDraw = True
    T1 = Timer
    With dbc.Query("SELECT * FROM test;")
        Do While .Step = SQLITE_ROW
            lGrid.AddRow .Value(0), .Value(1), .Value(2), .Value(3), .Value(4), .Value(5)
        Loop
    End With
    Label2 = lGrid.ItemCount & " items read in " & Format$(Timer - T1, "0.00 Seconds")
    lGrid.NoDraw = False
    Set dbc = Nothing
    
End Sub


