VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Predeterminar"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtEliminar 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Text            =   "0"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtColumnas 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Text            =   "50"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmInfo.frx":058A
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Eliminar archivo luego de imprimir (0:no-1:si)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuracion predeterminada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim mColumnas As Integer
    Dim mEliminar As Integer
    
    mColumnas = Int(txtColumnas.Text)
    mEliminar = Int(txtEliminar.Text)
    
    Open "config.ini" For Output As #1
    Write #1, mColumnas
    Write #1, mEliminar
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim mColumnas As Integer
    Dim mEliminar As Integer
    Dim strValue As String
    Dim mFile As String
    Dim mLongitud As Integer
    Dim mIni  As Long
    
    Open "config.ini" For Input As #1
    'cargo columnas
    Input #1, strValue
    If Err = 0 Then
        mColumnas = Int(strValue)
    Else
        mColumnas = 50
    End If
    
    'cargo auto eliminar
    Input #1, strValue
    If Err = 0 Then
        mEliminar = Int(strValue)
    Else
        mEliminar = 0
    End If
    Close #1
    txtColumnas = mColumnas
    txtEliminar = mEliminar
    
    mFile = ""
    If Command <> "" Then
        mLongitud = Len(Command)
        mFile = Mid(Command, 2, mLongitud - 2)
        Open mFile For Input As #2
        If Err = 0 Then
            Printer.Font = "Courier"
            Printer.FontSize = 8
            
            Do While Not EOF(2)
                Input #2, strValue
                If Len(strValue) > mColumnas Then
                    For mIni = 1 To Len(strValue) Step mColumnas
                        Printer.Print Mid(strValue, mIni, mColumnas)
                    Next
                Else
                    Printer.Print strValue
                End If
            Loop
            Close #2
            If mEliminar = 1 Then Kill (mFile)
            Printer.EndDoc
        End If
        Unload Me
    End If
    
End Sub
