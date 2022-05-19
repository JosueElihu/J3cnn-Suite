VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Charts demo"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   StartUpPosition =   1  'CenterOwner
   Begin JcnnTest.ucChartArea ucChartArea1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      Title           =   "Gráfico de datos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   0
      Border          =   -1  'True
      LinesCurve      =   -1  'True
      LinesWidth      =   2
      FillGradient    =   -1  'True
      LegendAlign     =   3
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsFormats   =   "S/. {V}"
      BorderRound     =   5
      BorderColor     =   12566463
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
Dim dbc           As SQLiteConnection
Dim CollData      As Collection
Dim CollProducts  As Collection
Dim sDate1      As String
Dim sDate2      As String
Dim Sql         As String
Dim tmp         As String
Dim i           As Long
Dim j           As Long

    
    'AÑADIR MESES AL GRÁFICO
    '-----------------------------------------------------------
    Set CollData = New Collection
    For i = 1 To 6
        CollData.Add Choose(i, "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    Next
    ucChartArea1.AddAxisItems CollData
    
    
    Set dbc = SQLite.Connection(App.Path & "\db\data.db")
    
    ' OBTENER LOS PRODUCTOS A UNA COLECCION
    '------------------------------------------------------------
    ' - CollProducts(1)(0)= ID Producto
    ' - CollProducts(1)(1)= Nombre Producto
    Set CollProducts = dbc.Collection("SELECT id, product_name FROM products;")

    
    'AÑADIR VALORES
    '------------------------------------------------------------
    
    For j = 1 To CollProducts.Count
    
        Set CollData = New Collection
        
        For i = 1 To 6  ' ENERO - FEBRERO
        
            sDate1 = "2021-" & Format(i, "00") & "-01"
            sDate2 = LastMonthDay(i, 2021)
            
            Sql = "SELECT SUM(total_price) FROM view_sales WHERE sale_date >= '" & sDate1 & "' AND sale_date <= '" & sDate2 & "' AND product_id='" & CollProducts(j)(0) & "';"
            tmp = dbc.ResultString(Sql)
            CollData.Add Format(Val(tmp), "0.00")
            
        Next
       
        ucChartArea1.AddLineSeries CollProducts(j)(1), CollData, RandomColor
    Next
        
    Set CollData = Nothing
    Set CollProducts = Nothing
    Set dbc = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With ucChartArea1
        .Move .Left, .Top, Me.ScaleWidth - (.Left * 2), Me.ScaleHeight - (.Top * 2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub






' PRIVATE
'---------------------------------------------------------------------------------------------------------------------
Private Function LastMonthDay(ByVal sMonth As String, ByVal sYear As String)
    LastMonthDay = sYear & "-" & Format(sMonth, "00") & "-" & DaysInMonth(sMonth & "-" & sYear) '-> RETURN: 2021-02-28 [yyyy/mm/dd]
End Function
Private Function DaysInMonth(dteInput As Date) As Integer
    DaysInMonth = DateAdd("m", 1, dteInput) - dteInput
End Function

Private Function RandomColor() As Long
    RandomColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
End Function

