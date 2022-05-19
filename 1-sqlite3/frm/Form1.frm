VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQLite3  Connector"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
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
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton BtnAdvanced 
         Caption         =   "Extensiones"
         Height          =   510
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   6360
         Width           =   2775
      End
      Begin VB.CommandButton BtnAdvanced 
         Caption         =   "  ResultSet  -  Collection   ResultString"
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   5520
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Paginación"
         Height          =   510
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Ejemplo Login"
         Height          =   510
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Grafico de datos"
         Height          =   510
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Adjuntar Base de datos"
         Height          =   510
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Insertar BLOB's - Binding"
         Height          =   510
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Campos BLOB"
         Height          =   510
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton BtnTest 
         Caption         =   "Medir tiempo de insercion"
         Height          =   510
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avanzado"
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
         TabIndex        =   16
         Top             =   5160
         Width           =   1290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   3000
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   240
         X2              =   3000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otros ejemplos"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.CommandButton BtnMain 
         Caption         =   "        Set Info       (db pragma)"
         Height          =   615
         Index           =   3
         Left            =   3720
         TabIndex        =   6
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton BtnMain 
         Caption         =   "        Get Info       (db pragma)"
         Height          =   615
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton BtnMain 
         Caption         =   "          Backup            (Guardar en el disco)"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   5715
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton BtnMain 
         Caption         =   "Actualizar"
         Height          =   495
         Index           =   0
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base de datos en memoria"
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
         TabIndex        =   5
         Top             =   360
         Width           =   2250
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'/* Random data */
Private Const FIRST_NAMES = "Alan,Alison,Alfie,Andrew,Bill,Dawn,Hannah,Brian,David,Jane,Jennifer,Karen,Gavin,Kerry,Kim,James,Laura,Lucy,Michael,Mellisa,Paula,Patrick,Sarah,Emily,Susan"
Private Const LAST_NAMES = "Anderson,Allen,Black,Evans,Bloggs,Brown,Clarke,Cole,Davis,Dawson,Brown,Gate,Johnson,Jones,Lawson,Lee,Richards,Ryan,Smith,Stephens,Temple,Turner,Wallace,White,Williams"
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Private dbc   As SQLiteConnection
Private Query As SQLiteQuery

Private Sub Form_Load()
Dim Sql     As String
Dim i       As Long
    

     Randomize Timer '[Random data]
    
    ' J3cnn.dll
    '-------------------------------------------------------------------------------------
    ' Es una dll Activex que permite conectar el motor de base de datos de SQLite3 y VB6
    ' Requiere de Sqlite3.dll
    '                                              - https://www.sqlite.org/index.html
    
    ' Para crear una conexión J3cnn proporciona un objeto constructor SQLite, esto
    ' permite crear una conexión en un solo paso:
   
    ' VB6 Method                               J3cnn Constructor (Recomendado)
    '-------------------------------------------------------------------------------------
    '    Set dbc = New SQLiteConnection        Set dbc = SQLite.Connection(":memory:")
    '    dbc.SQliteOpen ":memory:"
    '
    '-------------------------------------------------------------------------------------
    ' En este ejemplo se usara el objeto constructor para crear las conexiones de bases
    ' de datos.
    
    
    '-------------------------------------------------------------------------------------
    ' USAR ESTE OBJETO ACTIVEX SIN REGISTRO:
    '-------------------------------------------------------------------------------------
    ' Incluir J3cnnLoader.bas al proyecto, e incluir J3cnn.dll y Sqlite3.dll en la carpeta
    ' o una sub carpeta del proyecto (bin, plugis, lib, etc), o cargue el objeto activex
    ' desde una ruta personalizada de esta manera:
    '
    '    Call J3cnnLoader.LoadLib(App.Path & "\bin\jcnn\J3cnn.dll")
    '
    '-------------------------------------------------------------------------------------
    Call J3cnnLoader.LoadLib(App.Path & "\..\0-activex\J3cnn\J3cnn.dll")
    
    ' OBJETOS
    '-------------------------------------------------------------------------------------
    ' SQLiteConnection  -   Conexion a una base de datos
    ' SQLiteQuery       -   Objeto de consulta Sincronica
    ' SQLiteResultSet   -   Objeto de consulta Asincronica
    ' SQLiteCommand     -   Objeto de comando (insertar, actualizar)
    '-------------------------------------------------------------------------------------
    
    Set dbc = SQLite.Connection(":memory:") '- db en memoria
    
    ' MODO DE ERROR
    '-------------------------------------------------------------------------------------
    'dbc.Errmode = ErrShow  '- Muestra mensaje ante errores secundarios
    'dbc.Errmode = ErrRaise '- Genera error ante errores secundarios

    Sql = "CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, first_name TEXT, last_name TEXT, age INTEGER);"
    If dbc.Execute(Sql) <> SQLITE_OK Then dbc.ShowError
    
    ' TRANSACTIONS:
    '-------------------------------------------------------------------------------------
    '' - dbc.Transaction = BEGIN     -> Inicia transaccion
    '' - dbc.Transaction = COMMIT    -> Confirma transaccion
    '' - dbc.Transaction = ROLLBACK  -> Cancela transaccion [Puede usar ante un error]
    
    dbc.Transaction = BEGIN
    
    'For i = 0 To 30
    '    Sql = "INSERT INTO users (first_name,last_name,age) VALUES ('" & RandomName & "','" & RandomName2 & "'," & RandomNum(60) & ");"
    '    Call dbc.Execute(Sql)
    'Next
    
    Dim Cmd  As SQLiteCommand
    
    'Set cmd = dbc.Command("INSERT INTO users (first_name,last_name,age) VALUES (?,?,?);")
    'Set cmd = dbc.Command("INSERT INTO users VALUES (?,?,?,?);")
    Set Cmd = dbc.Command("INSERT INTO users VALUES(@id,@first_name,@last_name,@age);")
    
    For i = 0 To 30
        With Cmd
        
            '/* Por nombre de argumento */
            '------------------------------
            '.Bind !first_name, RandomName
            '.Bind !last_name, RandomName2
            '.Bind !age, RandomNum(60)
            '.Step
            
            '/* Por indice de argumento */
            '------------------------------
            '.Bind 1, Null
            '.Bind 2, RandomName
            '.Bind 3, RandomName2
            '.Bind 4, RandomNum(60)
            '.Step
            
            '/* Por propiedad Bindings */
            '------------------------------
            .Bindings = Array(Null, RandomName, RandomName2, RandomNum(60))
            .Step
        End With
        
    Next
    dbc.Transaction = COMMIT
    
    
    Set Query = dbc.Query("SELECT * FROM users;")
    Do While Query.Step = SQLITE_ROW
        '/*  Query.Value(2) ==  Query.Value("last_name") == Query!last_name    */
        List1.AddItem Query!last_name & ", " & Query!first_Name
    Loop
    
    
End Sub

Private Sub btnMain_Click(Index As Integer)
Dim tmp As String

     Select Case Index
    
        Case 0: 'RELOAD
            
            List1.Clear
            Query.Reset
            Do While Query.Step = SQLITE_ROW
                List1.AddItem Query!last_name & ", " & Query!first_Name
            Loop
            
        Case 1: 'Backup - volcar base de datos en memoria al disco
            
            ' BACKUP (Params)
            '-----------------------------------------------------------------------------
            ' BackupFile    : Archivo de destino
            ' dbSourceName  : Nombre de base de datos a copiar
            ' dbDestName    : Nombre de base de datos de destino
            '-----------------------------------------------------------------------------
    
            tmp = Format(Now, "ddmmyy-hhmmss") & ".bkup"
            If dbc.Backup(App.Path & "\" & tmp) = SQLITE_OK Then
                MsgBox "Backup creado: " & tmp, vbInformation
            Else
                MsgBox "Ocurrio un error al crear el backup", vbCritical
            End If
            
        
            ' dbSourceName, dbDestName:
            '--------------------------
            ' El nombre de la base de datos es "main" para la base de datos principal,
            ' "temp" para la base de datos temporal o el nombre especificado en una
            ' instrucción ATTACH para una base de datos adjunta.
        
        
        Case 2 'Get db pragma
        
            ' IMPORTANTE
            '-----------------------------------------------------------------------------
            ' dbc.Pragma("....") Devuelve o asigna valores pragma simples o de un unico
            ' valor, para pragmas que retornan una tabla de valores use SQLiteQuery Ejemplo:
            ' Set Query = dbc.Query("PRAGMA database_list;")
            
            tmp = "Encoding: " & dbc.Pragma("encoding") & vbNewLine
            tmp = tmp & "Journal mode: " & dbc.Pragma("journal_mode") & vbNewLine
            tmp = tmp & "Synchronous: " & dbc.Pragma("synchronous") & vbNewLine
            tmp = tmp & "Foreign keys: " & dbc.Pragma("foreign_keys") & vbNewLine
            tmp = tmp & "User version: " & dbc.Pragma("user_version") & vbNewLine
            tmp = tmp & "Application ID: " & dbc.Pragma("main.application_id") & vbNewLine
            
            MsgBox tmp, vbInformation
            
        Case 3 'Set db pragma
        
            dbc.Pragma("foreign_keys") = 1
            dbc.Pragma("user_version") = RandomNum(20)
            dbc.Pragma("main.application_id") = RandomNum(100)
            
            '- Setup wal mode
            '-------------------------------------
            'dbc.Pragma("journal_mode") = "WAL"
            'dbc.Pragma("synchronous") = "1"
            
            '- Setup defult mode
            '-------------------------------------
            'dbc.Pragma("journal_mode") = "DELETE"
            'dbc.Pragma("synchronous") = "2"
        
    End Select
End Sub
Private Sub BtnTest_Click(Index As Integer)
Dim tmp   As String
Dim i     As Long

    Select Case Index
        Case 0 ' MEDIR TIEMPO DE INSERCION
            Form4.Show 1, Me
        Case 1 ' BLOB'S
            Form2.Show 1, Me
        Case 2 ' INSERT BLOB'S
            Form3.Show 1, Me
        Case 3
        
            If dbc.Attach(App.Path & "\db\data.db", "mdb2") <> SQLITE_OK Then
                dbc.ShowError
                Exit Sub
            End If
            
            Dim Query2 As SQLiteQuery
 
            'CONSULTAR TABLAS DE LA BASE DE DATOS ADJUNTADA
            '--------------------------------------------------------------
            Set Query2 = dbc.Query("SELECT * FROM mdb2.sqlite_master WHERE name NOT LIKE 'sqlite_%';")
            Do While Query2.Step = SQLITE_ROW
                tmp = tmp & Query2!Type & ": " & Query2!tbl_name & vbNewLine
            Loop
            If tmp = vbNullString Then tmp = "Sin información"
            MsgBox tmp, vbInformation, "TABLES"
            
            
            'CONSULTAR BASES DE DATOS
            '--------------------------------------------------------------
            Set Query2 = dbc.Query("PRAGMA database_list;")
            tmp = "SEQ" & vbTab & "NAME" & vbTab & "FILE" & vbNewLine
            tmp = tmp & String$(75, "-") & vbNewLine
            Do While Query2.Step = SQLITE_ROW
                tmp = tmp & Query2!seq & vbTab & Query2!Name & vbTab & GetMPath(Query2!file) & vbNewLine
            Loop
            MsgBox tmp, vbInformation, "DATA BASES"
            
            
            'CONSULTAR DATOS DE LA BASE DE DATOS ADJUNTADA
            '--------------------------------------------------------------
            'Set Query2 = dbc.Query("SELECT * FROM mdb2.employees;")
            'Do While Query2.Step = SQLITE_ROW
            '    Debug.Print Query2!first_name & ", " & Query2!last_name
            'Loop
            
            Set Query2 = Nothing
            dbc.Detach "mdb2"
            
        Case 4 'GRAFICO DE DATOS
        
            Form5.Show 1, Me
            
        Case 5 'LOGIN
        
            Form6.Show 1, Me
            
        Case 6 'PAGINATION
            
            Form7.Show 1, Me
        
    End Select
    
End Sub

Private Sub BtnAdvanced_Click(Index As Integer)
    Select Case Index
        Case 0: 'Objects
            Form8.Show 1, Me
        Case 1 ' Extensions
            Form9.Show 1, Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Query = Nothing
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

Private Function GetMPath(sPath As String) As String
On Error GoTo e
    GetMPath = Right$(sPath, Len(sPath) - Len(App.Path))
e:
End Function
