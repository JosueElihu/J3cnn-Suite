VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert data"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Tabla Temporal"
      Height          =   615
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sqlite3 Exec"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Sqlite3 Bind"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insertar 50000 Resgistros"
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label lblTiming 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click en el botón para inicar la prueba"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   5460
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private dbc         As SQLiteConnection
Private db_File    As String


Private Sub Form_Load()
    db_File = App.Path & "\db\inserts.db"
End Sub
Private Sub Command1_Click()
    Call TestInsert1
End Sub
Private Sub Command2_Click()
    Call TestInsert2
End Sub
Private Sub Command3_Click()
    Call TestInsert3
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '
End Sub



Private Sub TestInsert1()
On Error Resume Next
Dim Cmd  As SQLiteCommand
Dim T1   As Single
Dim i    As Long


    EnableControls False
    Kill db_File
    
    Set dbc = SQLite.Connection(db_File)
    dbc.Errmode = ErrShow '- Mostrar mensaje de error
    
    dbc.Execute "CREATE TABLE test (ID INTEGER PRIMARY KEY, txt TEXT, dbl DOUBLE, int INTEGER, mdate SHORTDATE, mbln BOOLEAN);"
    Set Cmd = dbc.Command("INSERT INTO test Values(?,?,?,?,?,?);")
    
    T1 = Timer
    dbc.Pragma("synchronous") = "OFF"
    dbc.Transaction = BEGIN
    For i = 1 To 50000
    
        Cmd.Bind 1, Null
        Cmd.Bind 2, "SomeText_" & i
        Cmd.Bind 3, i / 2
        Cmd.Bind 4, i * 100
        Cmd.Bind 5, Now
        Cmd.Bind 6, IIf(i Mod 2, True, False)
        
        
        ' Bindings: Se puede enlazar datos con la propiedad 'Cmd.Bindings', mire la nota al final.
        '-----------------------------------------------------------------------------------------------------------
        ' Cmd.Bindings = Array(Null, "SomeText_" & i, i / 2, i * 100, Now, IIf(i Mod 2, True, False))
        '-----------------------------------------------------------------------------------------------------------
        
        If Cmd.Step <> SQLITE_DONE Then Exit For
        
        If i Mod 1000 = 0 Then
            lblTiming.Caption = Format$(i / 50000, "Percent")
            DoEvents
        End If
        
    Next
    dbc.Transaction = COMMIT
    dbc.Pragma("synchronous") = 2
    lblTiming.Caption = Format$(Timer - T1, "0.00 Seconds")
    
    Set Cmd = Nothing
    Set dbc = Nothing
    
    EnableControls True
End Sub

Private Sub TestInsert2()
On Error Resume Next
Dim T1  As Currency
Dim Sql As String
Dim i   As Long

    EnableControls False
    Kill db_File
    
    Set dbc = SQLite.Connection(db_File)
    dbc.Errmode = ErrShow                   '- Mostrar mensaje de error
    
    dbc.Execute "CREATE TABLE test (ID INTEGER PRIMARY KEY, txt TEXT, dbl DOUBLE, int INTEGER, mdate SHORTDATE, mbln BOOLEAN);"

    T1 = Timer
    dbc.Pragma("synchronous") = "OFF"
    dbc.Transaction = BEGIN
    For i = 1 To 50000
        Sql = "INSERT INTO test (txt,dbl,int,mdate,mbln) VALUES ('" & "SomeText_" & i & "','" & i / 2 & "','" & i * 100 & "','" & Now & "','" & True & "');"
        If dbc.Execute(Sql) <> SQLITE_OK Then Exit For
        
        If i Mod 1000 = 0 Then
            lblTiming.Caption = Format$(i / 50000, "Percent")
            DoEvents
        End If
    Next
    dbc.Transaction = COMMIT
    dbc.Pragma("synchronous") = 2
    lblTiming.Caption = Format$(Timer - T1, "0.00 Seconds")
    
    Set dbc = Nothing
    
    EnableControls True
End Sub

Private Sub TestInsert3()
On Error Resume Next
Dim Cmd  As SQLiteCommand
Dim T1   As Currency
Dim i    As Long

    EnableControls False
    Kill db_File
    
    Set dbc = SQLite.Connection(db_File)
    dbc.Errmode = ErrShow                   '- Mostrar mensaje de error
    
    dbc.Execute "CREATE TABLE temp.test (id INTEGER PRIMARY KEY, txt TEXT, dbl DOUBLE, int INTEGER, mdate SHORTDATE, mbln BOOLEAN);"
    Set Cmd = dbc.Command("INSERT INTO temp.test Values(:arg1,:arg2,:arg3,:arg4,:arg5,:arg6);")
    
    T1 = Timer
    dbc.Pragma("synchronous") = 0
    dbc.Transaction = BEGIN
    For i = 1 To 50000
    
        Cmd.Bind 1, Null
        Cmd.Bind 2, "SomeText_" & i
        Cmd.Bind 3, i / 2
        Cmd.Bind 4, i * 100
        Cmd.Bind 5, Now
        Cmd.Bind 6, True
        
        If Cmd.Step <> SQLITE_DONE Then Exit For
        
        If i Mod 1000 = 0 Then
            lblTiming.Caption = Format$(i / 50000, "Percent")
            DoEvents
        End If
        
    Next
    dbc.Transaction = COMMIT
    dbc.Execute "CREATE TABLE test AS SELECT * FROM temp.test;"
    dbc.Pragma("synchronous") = 2

    lblTiming.Caption = Format$(Timer - T1, "0.00 Seconds")
    
    Set Cmd = Nothing
    Set dbc = Nothing
    
    EnableControls True
End Sub

Private Sub EnableControls(Value As Boolean)
    Command1.Enabled = Value
    Command2.Enabled = Value
End Sub

' Bindings Property:
'-----------------------------------------------------------------------------------------------------------
' Se puede enlazar datos al cursor con la propiedad 'Cmd.Bindings', ejemplos:
'
'   Ejemplo 1:
'       Cmd.Bindings = Array("Some value", "Some data", ByteArray)
'
'   Ejemplo 2:
'       Dim sData(2) As Variant
'       sData(0) = "Some value"
'       sData(1) = "Some data"
'       sData(1) = ByteArray
'       Cmd.Bindings = sData
'
'   Ejemplo 3:
'       Dim cColl As New Collection
'       cColl.Add "Some value"
'       cColl.Add "Some data"
'       cColl.Add ByteArray
'       Cmd.Bindings = cColl
'
'   Ejemplo 4:
'       Cmd.Bindings = "Some value"
'
'   Ejemplo 5
'       Cmd.Bindings = Nothing   - Limpia los datos enlazados igual que 'Cmd.Clear'
'
' NOTA:
'-----------------------------------------------------------------------------------------------------------
' El usuario es responsable de asignar el array con la cantidad de valores segun la cantidad de parametros
' declarados en la instruccion SQL, 'Cmd.Bindings' enlaza los datos desde el PRIMER(1) parametro hasta el
' ultimo parametro de la instruccion SQL,o hasta el ultimo elemento del array o coleccion asignado.
'
' No asignar en el array los datos correctos o la cantidad de datos declarados en la instruccion SQL puede
' resultar en SQLITE_MISMATCH, especialmente con los campos 'PrimaryKey', Ejemplo:
'
'       Set Cmd = dbc.Command("INSERT INTO my_table VALUES(?,?,?)")  '  3 argumentos
'       Cmd.Bindings = Array("Some value", "Some data")              '  2 elementos asignados
'
' Si el primer campo de 'my_table' es PRIMARY KEY AUTO INCREMENT se obtendrá SQLITE_MISMATCH, porque a
' 'Cmd.Bindings' se le asigno un array con solo dos elementos, y 'Cmd.Bindings' enlazará los datos solo
' a dos de los argumentos de la instruccion SQL, quedando el tercer argumento como NULL, el siguiente
' ejemplo es correcto:
'
'       Cmd.Bindings = Array(Null, "Some value", "Some data")      '  3 elementos asignados
'
' Luedo de llamar a Cmd.Bindings  se debe llamar a Cmd.Step para ejecutar y guardar los datos
'-----------------------------------------------------------------------------------------------------------
