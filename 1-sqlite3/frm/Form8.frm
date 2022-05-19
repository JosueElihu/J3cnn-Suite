VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "ResultSet - Collection - GetString"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   7560
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   4095
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ResultString"
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Collection"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ResultSet"
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
      TabIndex        =   1
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbc As SQLiteConnection
Dim i   As Long

    
    Set dbc = SQLite.Connection(App.Path & "\db\data.db")
    
    
    ' SQLiteResultSet
    '-----------------------------------------------------------------------------------------------------------
    ' SQLiteResultSet guarda una copia completa de los datos en un array bidimencional con Index de FILAS y
    ' COLUMNAS basadas en 0 y la propiedad 'Matrix(0,0)'.
    '
    ' Propiedades:
    '
    '   Count       : Cantidad total de Registros (FILAS)
    '   FieldCount  : Cantidad total de campos   (COLUMNAS)
    '   FieldName   : Nombre del campo o columna, el primer INDEX es 0
    '   Matrix(n,n) : Retorna el valor almacenado en el array, el primer INDEX  de la FILA y COLUMNA es 0
    '   Position    : Muestra la posicion actual del registro al recorrer con 'NextRow'
    '   Value(col)  : Retorna el valor almacenado según la column y la posición actual del registro ('Position')
    '
    ' Rutinas:
    '
    '   Reset           : Reestablece la posicion de registros
    '   NextRow         : Move la posicion actual a la siguiente
    '-----------------------------------------------------------------------------------------------------------
    
    
    Dim Rs As SQLiteResultSet
    Set Rs = dbc.ResultSet("SELECT * FROM customers;")
  
    Do While Rs.NextRow
        
        'Obtener valores [Value]
        '------------------------------------------------------
        ' - Rs.Value("first_Name")
        ' - Rs.Value(1)
        ' - Rs!first_Name
        '------------------------------------------------------
        'List1.AddItem Rs!first_Name & ", " & Rs!last_name
    Loop
    

    For i = 0 To Rs.Count - 1
    
        'Obtener valores [Matrix]
        '------------------------------------------------------
        ' - Rs.Matrix(i, "first_Name")
        ' - Rs.Matrix(i, 1)
        '----------------------------------------------------------------
        List1.AddItem Rs.Matrix(i, "first_Name") & ", " & Rs.Matrix(i, 2)
        
    Next
    

    
    ' COLLECTION
    '-----------------------------------------------------------------------------------------------------------
    ' Crea una coleccion da datos a partir de una consulta, si la consulta retorna un solo campo, crea una
    ' coleccion de datos simple, si la consulta retorna mas de un campo crea una coleccion de ARRAYS.
    
    Dim CollData As Collection
    
    'Set CollData = dbc.Collection("SELECT * FROM products;")             'Array [0,1,2...n]
    'Set CollData = dbc.Collection("SELECT id,unit_price FROM products;") 'Array [0,1,2...n]
    'Set CollData = dbc.Collection("SELECT product_name FROM products;")  'Strings
    'Set CollData = dbc.Collection("SELECT id FROM products;")            'Strings
    
    Set CollData = dbc.Collection("SELECT * FROM products;")
    For i = 1 To CollData.Count
    
        ' READ ARRAY
        '-----------------------------------------------------------------------------
        ' Si la consulta retorna más de un campo, en cada ITEM de la coleccion
        ' se guarda un array de datos, el primer INDEX de este aray es 0
        '
        '
        '   - CollData(1)(0), CollData(1)(1), CollData(1)(2),  ... CollData(1)(n)
        '   - CollData(2)(0), CollData(2)(1), CollData(2)(2),  ... CollData(2)(n)
        '
        ' Dim vElmnt As Variant
        '    vElmnt = CollData(1)
        '    vElmnt(0), vElmnt(1), vElmnt(2),   ... vElmnt(n)
        '    vElmnt = CollData(2)
        '    vElmnt(0), vElmnt(1), vElmnt(2),   ... vElmnt(n)
        '-----------------------------------------------------------------------------
        
        Dim vElmnt As Variant
        vElmnt = CollData(i)

        'Debug.Print IsArray(vElmnt)
        'Call mvPrintArray(vElmnt)
        List2.AddItem vElmnt(0) & " - " & vElmnt(1)
        
    Next
    
    Set CollData = dbc.Collection("SELECT id FROM products;")
    For i = 1 To CollData.Count
    
        ' READ STRING
        '-----------------------------------------------------------------------------
        ' Si la consulta retorna un solo campo, en cada ITEM de la coleccion se guarda
        ' el valor de cada FILA.
        '
        ' - CollData(1), CollData(2), CollData(3), ... CollData(n)
        '-----------------------------------------------------------------------------
        
        List3.AddItem CollData(i)
        
    Next
    
    
    ' ResultString
    '-----------------------------------------------------------------------------------------------------------
    ' Retorna un String como resultado de una consulta, si hay varias columnas y filas, retorna la ultima fila
    ' y la primera columna, esta funcion envuelve la api Sqlite_exec a traves de un callback. debido a eso no
    ' no retorna datos de tipo BLOB.

    Text1 = dbc.ResultString("SELECT COUNT(*) FROM sales;")
    Text2 = dbc.ResultString("SELECT SUM(total_price) FROM view_sales;")
    
    Text1 = dbc.ResultString("SELECT COUNT(*) FROM sales;")
    Set dbc = Nothing
    
End Sub




Private Sub mvPrintArray(vData As Variant)
On Error GoTo e
Dim vElmnt As Variant
    For Each vElmnt In vData
        Debug.Print vElmnt,
    Next
    Debug.Print ""
e:
End Sub
