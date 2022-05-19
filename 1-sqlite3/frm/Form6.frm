VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnMain 
      Caption         =   "Ingresar"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton BtnMain 
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Txt0 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   272
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private dbc     As SQLiteConnection


Private Sub Form_Load()
    Set dbc = SQLite.Connection(App.Path & "\db\data.db")
End Sub

Private Sub btnMain_Click(Index As Integer)
    Select Case Index
        Case 0: Unload Me
        Case 1:
        
            If Len(Trim$(Txt0)) = 0 Then MsgBox "¡Complete los datos!", vbCritical: Txt0.SetFocus: Exit Sub
            If Len(Trim$(Txt1)) = 0 Then MsgBox "¡Complete los datos!", vbCritical: Txt1.SetFocus: Exit Sub
            
            Dim ssRet As ssSQliteResult
            ssRet = dbc.Query("SELECT id FROM employees WHERE first_name='" & Txt0 & "' AND pwd='" & Txt1 & "';").Step
            If ssRet = SQLITE_ROW Then
                MsgBox "¡Bienvenido!", vbInformation
                Unload Me
            Else
                MsgBox "¡Usuario o contraseña incorrecta!", vbCritical: Txt0.SetFocus
            End If
            
            
            ' ResultString :
            '------------------------------------------------------------------------------------------------------------------------
            'If dbc.ResultString("SELECT id FROM employees WHERE first_name='" & txt0 & "' AND pwd='" & txt1 & "';") <> vbNullString Then
            '    MsgBox "¡Bienvenido!", vbInformation
            '    Unload Me
            'Else
            '    MsgBox "¡Usuario o contraseña incorrecta!", vbCritical: txt0.SetFocus
            'End If
            
            
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set dbc = Nothing
End Sub
