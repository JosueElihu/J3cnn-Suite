VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "J3cnn_c"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8160
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
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   8160
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   8160
      Begin VB.Label lblc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblc"
         ForeColor       =   &H00B5752F&
         Height          =   195
         Left            =   5880
         TabIndex        =   9
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQLCipher"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cifrado AES transparente de 256 bits"
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   2685
      End
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejemplo2"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employees"
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
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    '  J3cnn_c.dll
    '---------------------------------------------------------------------------------
    ' Es una dll Activex que conecta SQLCipher y VB6, permite el uso de bases de datos
    ' cifradas de sqlite. Requiere de sqlcipher.dll vinculado con libcrypto-1_1.dll
    '
    '                                      https://www.zetetic.net/sqlcipher/
    '                                      https://www.zetetic.net/sqlcipher/license/
    '---------------------------------------------------------------------------------
    
    ' Abrir base de datos (Ej-1)
    '---------------------------------------------------------------------------------
    ' Set dbc = New SQLiteConnection
    ' dbc.SQliteOpen App.Path & "\db\data.enc"
    ' dbc.Execute "PRAGMA key = 'my_db_key';"  ||  dbc.Pragma("key") = "my_db_key"
    
    
    ' Abrir base de datos (Ej-2)
    '---------------------------------------------------------------------------------
    ' Set dbc = New SQLiteConnection
    ' dbc.SQliteOpen App.Path & "\db\data.enc", "my_db_key"
    
    
    ' Abrir base de datos (Ej-3) *Recomendado*
    '---------------------------------------------------------------------------------
    ' Set dbc = SQlite.Connection(App.Path & "\db\data.enc", "my_db_key")
    
    
    ' Cambiar clave
    '----------------------------------------------------------------------------------
    ' dbc.Execute "PRAGMA rekey = '123456';"
    ' dbc.Pragma("rekey") = "123456"
    
    
    ' Adjuntar base de datos cifrada
    '----------------------------------------------------------------------------------
    ' dbc.Attach App.Path & "\db2.enc", "mdb2", "my_db2_key"
    
    ' Cambiar clave base de datos adjuntada
    '----------------------------------------------------------------------------------
    ' dbc.Execute "PRAGMA mdb2.rekey = '123456';"
    ' dbc.Pragma("mdb2.rekey") = "123456"
    
    
    ' Adjuntar base de datos sin formato
    '----------------------------------------------------------------------------------
    ' dbc.Attach App.Path & "\data.db", "mdb3", vbNullstring
    
    
    ' Cifrar base de datos sin formato
    '---------------------------------------------------------------------------------
    ' Set dbc = SQLite.Connection(App.Path & "\db\data.db")
    ' dbc.Attach App.Path & "\db\data.enc", "encrypted", "my_db_key"
    ' dbc.Execute "SELECT sqlcipher_export('encrypted');"
    ' dbc.Detach "encrypted"
    
    
    ' Descifrar base de datos SQLCipher a una base de datos sin formato
    '---------------------------------------------------------------------------------
    ' Set dbc = SQLite.Connection(App.Path & "\db\data.enc", "my_db_key")
    ' dbc.Attach App.Path & "\db\data.db", "mdb2", vbNullString
    ' dbc.Execute "SELECT sqlcipher_export('mdb2');"
    ' dbc.Detach "mdb2"
    
    
    ' Copiar base de datos sin formato a nueva base de datos cifrada vacía
    '---------------------------------------------------------------------------------
    ' Set dbc = SQLite.Connection(App.Path & "\db\data.enc", "my_db_key")
    ' dbc.Attach App.Path & "\db\data.db", "mdb2", vbNullString
    ' dbc.Execute "SELECT sqlcipher_export('main', 'mdb2');"
    ' dbc.Detach "mdb2"
    
    
    '-------------------------------------------------------------------------------------
    ' USAR ESTE OBJETO ACTIVEX SIN REGISTRO:
    '-------------------------------------------------------------------------------------
    ' Incluir J3cnnLoader.bas al proyecto, e incluir J3cnn_c.dll y Sqlite3.dll en la
    ' carpeta o una sub carpeta del proyecto (bin, plugis, lib), o cargue el objeto activex
    ' desde una ruta personalizada de esta manera:
    '
    '    Call J3cnnLoader.LoadLib(App.Path & "\bin\jcnn_c\J3cnnc_c.dll")
    '
    '-------------------------------------------------------------------------------------
    Call J3cnnLoader.LoadLib(App.Path & "\..\0-activex\J3cnn_c\J3cnn_c.dll")
    

    Dim dbc As SQLiteConnection
    Set dbc = SQLite.Connection(App.Path & "\db\data.enc", "my_db_key")
    
    With dbc.Query("SELECT * FROM customers")
        Do While .Step = SQLITE_ROW
            List2.AddItem .Value(2) & "," & .Value(1)
        Loop
    End With
    
    Dim i   As Long
    With dbc.ResultSet("SELECT * FROM employees")
        For i = 0 To .Count - 1
            List1.AddItem .Matrix(i, 2) & ", " & .Matrix(i, 1)
        Next
    End With
    
    lblc = "SQLite: " & dbc.Version & vbNewLine & "SQLCipher: " & dbc.Pragma("cipher_version")
    Set dbc = Nothing
End Sub
Private Sub Command1_Click()
    Form2.Show 1, Me
End Sub

