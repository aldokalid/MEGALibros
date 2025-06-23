VERSION 5.00
Begin VB.Form GenreEditorForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Géneros favoritos"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton saveBtn 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CheckBox scfiChk 
      Caption         =   "Ciencia ficción"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CheckBox suspenseChk 
      Caption         =   "Suspenso"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CheckBox dramaChk 
      Caption         =   "Drama"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CheckBox romanceChk 
      Caption         =   "Romance"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CheckBox horrorChk 
      Caption         =   "Terror"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "GenreEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' Obtiene los ajustes del usuario.
    Dim rs As ADODB.Recordset
    Set rs = New Recordset
    
    Dim query As String
    query = "SELECT * FROM get_liked_genres('" & currentSession & "')"
    
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    Do Until rs.EOF
        Select Case (rs!genre)
            Case "Terror"
                horrorChk.Value = 1
            Case "Romance"
                romanceChk.Value = 1
            Case "Drama"
                dramaChk.Value = 1
            Case "Suspenso"
                suspenseChk.Value = 1
            Case "Ciencia ficción"
                scfiChk.Value = 1
            Case Else
                Debug.Print ">ERR:INVALID; got invalid genre: " & rs!genre
        End Select
        
        rs.MoveNext
    Loop
    
    rs.Close: Set rs = Nothing
End Sub

' Guardar los ajustes.
Private Sub saveBtn_Click()
    Dim query As String
    query = "EXEC update_liked_genres '" & _
        currentSession & "', " & _
        CStr(horrorChk.Value) & ", " & _
        CStr(romanceChk.Value) & ", " & _
        CStr(dramaChk.Value) & ", " & _
        CStr(suspenseChk.Value) & ", " & _
        CStr(scfiChk.Value) & ";"
    
    On Error GoTo ErrCatch
        conn.Execute query, , adExecuteNoRecords
        Unload Me
    
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "Los ajustes no se guardaron: " & Err.Description, vbCritical, "Error"
    End If
End Sub
