VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrar o registrarse"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton signUpBtn 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton singInBtn 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox passwordTxt 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox usernameTxt 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label passwordLbl 
      Alignment       =   2  'Center
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label usernameLbl 
      Alignment       =   2  'Center
      Caption         =   "Nombre de usuario"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Cuando un usuario quiere registrarse.
Private Sub signUpBtn_Click()
    Dim username As String, password As String

    username = Replace(Trim(usernameTxt.Text), "'", "")
    password = Replace(Trim(passwordTxt.Text), "'", "")

    If username = "" Or username = " " Then
        MsgBox "El nombre de usuario es inválido", vbCritical
    ElseIf password = "" Or password = " " Then
        MsgBox "La contraseña es inválida", vbCritical
    End If

    On Error GoTo ErrCatch
        conn.Execute "EXEC sign_up '" & username & "', '" & password & "';"
        currentSession = username
        MsgBox "Fuiste registrado y autenticado exitosamente", vbInformation
        Unload Me
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "No se pudo registrar tu cuenta: " & Err.Description, vbCritical
    End If
End Sub

' Cuando un usuario quiere iniciar sesión
Private Sub singInBtn_Click()
    Dim username As String, password As String

    username = Replace(Trim(usernameTxt.Text), "'", "")
    password = Replace(Trim(passwordTxt.Text), "'", "")

    If username = "" Or username = " " Then
        MsgBox "El nombre de usuario es inválido", vbCritical
    ElseIf password = "" Or password = " " Then
        MsgBox "La contraseña es inválida", vbCritical
    End If

    On Error GoTo ErrCatch
        ' adExecuteNoRecords ayuda a controlar mejor los errores cuando no se esperan registros.
        conn.Execute "EXEC sign_in '" & username & "', '" & password & "';", , adExecuteNoRecords
        currentSession = username
        MsgBox "Bienvenido, " & username, vbInformation
        Unload Me
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "No se pudo iniciar sesión: " & Err.Description, vbCritical
    End If
End Sub
