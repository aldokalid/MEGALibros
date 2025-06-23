VERSION 5.00
Begin VB.Form AddOrUpdateBookForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton saveBtn 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CheckBox recommendChk 
      Caption         =   "Recomendar"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
   End
   Begin VB.ComboBox rateCbb 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3240
      Width           =   4455
   End
   Begin VB.ComboBox genreCbb 
      Height          =   315
      ItemData        =   "AddOrUpdateBookForm.frx":0000
      Left            =   120
      List            =   "AddOrUpdateBookForm.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox authorTxt 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox titleTxt 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label rateLbl 
      Caption         =   "Calificación"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label genreLbl 
      Caption         =   "Género"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label authorLbl 
      Caption         =   "Autor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label titleLbl 
      Caption         =   "Título"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "AddOrUpdateBookForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bookId As String

Dim genres() As String
Dim currentGenre As String

' Carga los géneros desde la base de datos.
Private Sub loadGenres()
    Dim dimension As Integer ' Dimensión del arreglo.
    Dim capacity As Integer ' Capacidad del arrreglo por bloques.
    dimension = 0
    capacity = 10
    
    Erase genres ' Limpia el arreglo de géneros.
    ReDim genres(0 To 9) ' Inicializa el arreglo de géneros.

    ' ** Conexión a la base de datos.
    Dim rs As ADODB.Recordset
    Set rs = New Recordset
    
    rs.Open "SELECT * FROM Genre", conn, adOpenStatic, adLockReadOnly
    
    ' Asignación de datos.
    genreCbb.Clear
    genreCbb.AddItem "Selecciona una opción..."
    genreCbb.ListIndex = 0
    
    Do Until rs.EOF
        genreCbb.AddItem rs!Name
        
        ' Redimensiona el arreglo de géneros si se requiere.
        If dimension > capacity Then
            capacity = capacity + 10
            ReDim Preserve genres(0 To capacity)
        End If
        
        ' Agrega el género y su id al arreglo de géneros.
        genres(dimension) = rs!Name & "-" & rs!Idgenre
        dimension = dimension + 1
        
        rs.MoveNext
    Loop
End Sub

' Obtiene la llave del género seleccionado.
Private Function getSelectedGenreId() As String
    Dim genre As String
    Dim i As Integer
    
    For i = LBound(genres) To UBound(genres)
        Dim auxGenre() As String
        auxGenre = Split(genres(i), "-")

        If auxGenre(0) = genreCbb.Text Then
            genre = auxGenre(1)
            Exit For
        End If
    Next i
    
    getSelectedGenreId = genre
End Function

' Obtiene el índice de la lista de géneros según la coincidencia
Private Function getGenreIndex(idOrName As String) As Integer
    Dim i As Integer
    
    For i = LBound(genres) To UBound(genres)
        Dim auxGenre() As String
        auxGenre = Split(genres(i), "-")

        If auxGenre(0) = idOrName Or auxGenre(1) = idOrName Then
            getGenreIndex = i
            Exit For
        End If
    Next i
End Function


' Carga un libro de la base de datos si se le ha proporcionado un id para modificarlo.
Private Sub loadBook()
    Dim found As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New Recordset
    found = False
    
    rs.Open "SELECT * FROM Book WHERE Idbook = " & bookId & ";", conn, adOpenStatic, adLockReadOnly
        
    If rs.EOF Then
        MsgBox "No se encontró el libro", vbCritical, "Libro no encontrado"
        Unload Me
    Else
        titleTxt.Text = rs!title ' Coloca el título.
        authorTxt.Text = rs!author ' Coloca el autor.
        
        ' Asigna el género.
        genreCbb.ListIndex = getGenreIndex(rs!Idgenre) + 1
        ' Asigna la calificación.
        rateCbb.ListIndex = rs!rate
        ' Asigna la recomendación.
        recommendChk.Value = IIf(rs!Recommended, 1, 0)
    End If
End Sub

Private Sub Form_Load()
    ' Carga el título de la ventana.
    AddOrUpdateBookForm.Caption = IIf(bookId <> "", "Editar libro", "Agregar libro")
    
    ' Limpia el título.
    titleTxt.Text = ""
    ' Limpia el autor.
    authorTxt.Text = ""
    
    ' Carga elementos al combobox de géneros.
    loadGenres

    ' Carga elementos al combobox de calificación.
    rateCbb.Clear
    rateCbb.AddItem "Selecciona una opción..."
    rateCbb.ListIndex = 0
    rateCbb.AddItem "1"
    rateCbb.AddItem "2"
    rateCbb.AddItem "3"
    rateCbb.AddItem "4"
    rateCbb.AddItem "5"
    
    ' Limpia la recomendación
    recommendChk.Value = 0
    
    ' Carga el libro por modificar si fue especificado.
    If bookId <> "" Then
        loadBook
    End If
End Sub

' Guardar libro.
Private Sub saveBtn_Click()
    Dim title As String: Dim author As String
    title = Replace(Replace(Trim(titleTxt.Text), "'", ""), "|", "")
    author = Replace(Replace(Trim(authorTxt.Text), "'", ""), "|", "")
    
    If title = "" Or title = " " Then
        MsgBox "Ingresa un título para el libro", vbExclamation
    ElseIf author = "" Or author = " " Then
        MsgBox "Ingresa el autor del libro", vbExclamation
    ElseIf genreCbb.ListIndex <= 0 Then
        MsgBox "Selecciona un género", vbExclamation, "Género inválido"
    ElseIf rateCbb.ListIndex <= 0 Then
        MsgBox "Selecciona una calificación", vbExclamation, "Calificación inválida"
    Else ' Preara los datos para ser registrados.
        Dim genre As String
        Dim rate As String
        Dim recommend As String
        
        rate = rateCbb.ListIndex
        recommend = IIf(recommendChk.Value = 0, 0, 1)
        genre = getSelectedGenreId
        
        ' Realiza la consulta.
        On Error GoTo ErrCatch
            Dim query As String
            query = "EXEC create_or_update_book " & _
                IIf(bookId <> "", "'" & bookId & "'", "NULL") & _
                ", '" & title & "', '" & _
                author & "', " & _
                genre & ", " & _
                rate & ", " & _
                recommend & ";"
            
            Debug.Print query
            
            conn.Execute query, , adExecuteNoRecords
            
            MsgBox "Libro " & IIf(bookId <> "", "actualizado.", "guardado."), vbInformation
            Unload Me
    End If
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "No se pudo realizar la consulta: " & Err.Description, vbCritical
    End If
End Sub
