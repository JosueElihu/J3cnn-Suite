VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BLOB'S"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6015
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
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Frame Frame1 
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
      Height          =   3375
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1980
         Left            =   480
         ScaleHeight     =   132
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   134
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "Form2"
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



Private Pics   As Collection

Private Sub Form_Load()

    ' GDI+
    '-----------------------------------------------------------------------
    ' Inicializamos GDI+ para convertir array de bytes en Objetos StdPicture
    '-----------------------------------------------------------------------
    Call ManageGDIP(True)

    Dim dbc   As SQLiteConnection
    Set dbc = SQLite.Connection(App.Path & "\db\data.db3mc", "my_db_key")
    Set Pics = New Collection
    
    With dbc.Query("SELECT * FROM employees;")
        Do While .Step = SQLITE_ROW
            List1.AddItem !last_name & ", " & !first_name
            Pics.Add ArrayToPicture(.Blob(4))
        Loop
    End With
    
    Set dbc = Nothing
    
    ' - Finalizar GDI+
    Call ManageGDIP(False)
    
    If List1.ListCount Then List1.ListIndex = 0
End Sub
Private Sub List1_Click()
    Set Image1.Picture = Pics(List1.ListIndex + 1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Pics = Nothing
End Sub




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
        Dim gdipSI(3) As Long
        gdipSI(0) = 1&
        Call GdiplusStartup(gdip_token, gdipSI(0), ByVal 0)
    Else
        Call GdiplusShutdown(gdip_token): gdip_token = 0
    End If
End Sub





