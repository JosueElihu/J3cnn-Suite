VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Blob's"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   2175
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1980
         Left            =   120
         ScaleHeight     =   132
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   134
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   2010
         Begin VB.Image Image1 
            Height          =   1920
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1920
         End
      End
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Guardar"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton btnMain 
      Caption         =   "Selecionar perfil"
      Height          =   495
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   241
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   3615
         Begin VB.TextBox txt1 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txt0 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' StdPicture - With GDI+ [bmp|jpg|png|ico|gif]
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private gdip_token As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Private dbc    As SQLiteConnection
Private cmndex As CmnDialogEx
Private Strm() As Byte



Private Sub Form_Load()
    
    ' GDI+
    '-----------------------------------------------------------------------
    ' Inicializamos GDI+ para convertir array de bytes en Objetos StdPicture
    '-----------------------------------------------------------------------
    Call ManageGDIP(True)
    
    
    'SQLite db
    '------------------------------------------------------------------------
    Set dbc = SQLite.Connection(App.Path & "\db\data.db")
    
    
    'OPEN FILE DIALOG
    '------------------------------------------------------------------------
    Set cmndex = New CmnDialogEx
    cmndex.Filter = "Imagenes|*.bmp;*.ico;*.png;*.gif;*.jpg;*.jpeg"
    
End Sub

Private Sub btnMain_Click(Index As Integer)
Dim Cmd As SQLiteCommand

    Select Case Index
        Case 0:
        
            If cmndex.ShowOpen(Me.hWnd) Then
                Erase Strm
                Strm = GetFileArray(cmndex.FileName)
                Set Image1.Picture = ArrayToPicture(Strm)
            End If
            
        Case 1:
            
            If Len(Trim$(Txt0)) = 0 Then MsgBox "¡Complete los datos!", vbCritical: Txt0.SetFocus: Exit Sub
            If Len(Trim$(Txt1)) = 0 Then MsgBox "¡Complete los datos!", vbCritical: Txt1.SetFocus: Exit Sub
            If IsArray(Strm) = 0 Then MsgBox "¡Seleccione perfil!", vbCritical: BtnMain(0).SetFocus: Exit Sub
            
            
            'CREAR EL OBJETO DE COMMANDO
            '-----------------------------------------------------------------------------------------------------------
            'Set Cmd = dbc.Command("INSERT INTO employees (first_name,last_name,profile) VALUES (?,?,?);")
            Set Cmd = dbc.Command("INSERT INTO employees (first_name,last_name,profile) VALUES (:arg1,:arg2,:profile);")
            
            
            ' Bind Function:
            '-----------------------------------------------------------------------------------------------------------
            ' Enlaza datos al objeto de commando por INDICE del argumento o por el NOMBRE del argumento especificado en
            ' la instruccion SQL, el INDICE de los argumentos en el objeto de comando siempre inicia en 1.
            '
            '   Enlazar datos por INDICE del argumento : [1, 2, 3, 4 ]
            '       -   Cmd.Bind 1, "some value"
            '       -   Cmd.Bind 2, "some value"
            '       -   Cmd.Bind 3, "some value"
            '
            '   Enlazar datos por NOMBRE del argumento : [:arg1, :arg2, :arg3, :arg4 ]
            '       -   Cmd.Bind ":arg1", "some value"
            '       -   Cmd.Bind ":arg2", "some value"
            '       -   Cmd.Bind ":arg3", "some value"
            '
            ' Se debe de llamar a Cmd.Step luego de enlazar los datos para ejecutar y guardar los datos
            '-----------------------------------------------------------------------------------------------------------
            
            Cmd.Bind "arg1", Txt0       ' Bind text
            Cmd.Bind "arg2", Txt1       ' Bind text
            Cmd.Bind "profile", Strm    ' Bind Image Bytes
            
            
            If Cmd.Step = SQLITE_DONE Then
            
                MsgBox "¡Guardado! " & vbNewLine & "ID: " & dbc.LastInsertID, vbInformation
                
                Txt0 = vbNullString
                Txt1 = vbNullString
                Set Image1.Picture = Nothing
                Erase Strm
                
            Else
                dbc.ShowError
            End If
            Set Cmd = Nothing
            
            
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
            
            
            
            'INSERT DATA - SQLITE MULTIPLE BINDING
            '-----------------------------------------------------------------------------------------------------------
            'Dim lResult As ssSQliteResult
            'Dim i       As Integer
            '
            'Set Cmd = dbc.Command("INSERT INTO employees (first_name,last_name,profile) VALUES (:arg1,:arg2,:profile);")
            
            'dbc.Transaction = BEGIN
            'For i = 0 To 100
            '    Cmd.Bind "arg1", txt0
            '    Cmd.Bind "arg2", txt1
            '    Cmd.Bind "profile", Strm
            '
            '    lResult = Cmd.Step
            'Next
        
            'If lResult = SQLITE_DONE Then
            '    dbc.Transaction = COMMIT
            'Else
            '    dbc.Transaction = ROLLBACK
            'End If

            '' - dbc.Transaction = BEGIN     -> Inicia transaccion
            '' - dbc.Transaction = COMMIT    -> Confirma transaccion
            '' - dbc.Transaction = ROLLBACK  -> Cancela transaccion
            

    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set dbc = Nothing
    Set cmndex = Nothing
    
    ' - Finalizar GDI+
    ManageGDIP False
End Sub





' Get Byte Array From  Image File
'-------------------------------------------------------------------------
Private Function GetFileArray(sFileName As String) As Byte()
On Error GoTo e
Dim FF      As Integer

    FF = FreeFile
    Open sFileName For Binary As #FF
    ReDim GetFileArray(LOF(FF) - 1)
    Get #FF, , GetFileArray
    Close #FF
e:
End Function

Private Function IsArray(ByRef data() As Byte) As Boolean
On Error GoTo e
    IsArray = UBound(data) >= 0
e:
End Function


' GDIP  -   Convert to StdPicture
'-------------------------------------------------------------------------
Public Function ArrayToPicture(ByRef bvData() As Byte) As StdPicture
On Error GoTo e
Dim IStream   As IUnknown
Dim hBmp      As Long
Dim hBmp2     As Long


    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If IStream Is Nothing Then Exit Function
    If GdipLoadImageFromStream(IStream, hBmp) = 0 Then
        Dim GUID(3) As Long
        Dim lPic(4) As Long
        
        GdipCreateHBITMAPFromBitmap hBmp, hBmp2, 0
        GdipDisposeImage hBmp
        
        lPic(0) = 20
        lPic(1) = vbPicTypeBitmap
        lPic(2) = hBmp2
        
        'IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        GUID(0) = &H7BF80980: GUID(1) = &H101ABF32: GUID(2) = &HAA00BB8B: GUID(3) = &HAB0C3000
        Call OleCreatePictureIndirect(lPic(0), GUID(0), True, ArrayToPicture)
    
    End If
    Set IStream = Nothing
e:
End Function
Private Sub ManageGDIP(ByVal Startup As Boolean)
    If Startup Then
        If gdip_token = 0& Then
            Dim gdipSI(3) As Long
            gdipSI(0) = 1&
            Call GdiplusStartup(gdip_token, gdipSI(0), ByVal 0)
        End If
    Else
        If gdip_token <> 0 Then Call GdiplusShutdown(gdip_token): gdip_token = 0
    End If
End Sub

