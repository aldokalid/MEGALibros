VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MEGALibros"
   ClientHeight    =   5655
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView booksListView 
      Height          =   4695
      Left            =   2040
      TabIndex        =   9
      Top             =   840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "genre"
         Text            =   "Género"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "author"
         Text            =   "Autor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "title"
         Text            =   "Título"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "rate"
         Text            =   "Calificación"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.CommandButton searchBtn 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   9600
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox searchBarTxt 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton readBtn 
         Caption         =   "Leídos"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton genreListBtn 
         Caption         =   "Mis géneros favoritos"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton nonLikedBtn 
         Caption         =   "No gustados"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton toReadBtn 
         Caption         =   "Por leer"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label listsBtn 
         Caption         =   "Listas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label searchByTitleLbl 
      Caption         =   "Buscar por título"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu megaToolbarItem 
      Caption         =   "MEGA"
      Begin VB.Menu fullCatalogBtn 
         Caption         =   "Todo el catálogo"
      End
      Begin VB.Menu recommendedBtn 
         Caption         =   "Recomendaciones"
      End
      Begin VB.Menu addBookBtn 
         Caption         =   "Agregar libro"
      End
   End
   Begin VB.Menu userToolbarItem 
      Caption         =   "Usuario"
      Begin VB.Menu signInOutBtn 
         Caption         =   "Iniciar sesión"
      End
      Begin VB.Menu favGenresBtn 
         Caption         =   "Géneros favoritos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu bookToolbarItem 
      Caption         =   "Libro"
      Begin VB.Menu markBookAsReadBtn 
         Caption         =   "Marcar como leído"
      End
      Begin VB.Menu addBookToForReadBtn 
         Caption         =   "Agregar a Por leer"
      End
      Begin VB.Menu markBookAsNonLikedBtn 
         Caption         =   "Marcar como No gustado"
      End
      Begin VB.Menu updateBookBtn 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu deleteBookBtn 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables globales
Dim currentBookKey As String
' Arreglo de ajustes de libros (id|'leído'|'quiere leer'|'no le gustó')
Dim booksSettingsArr() As String

Private Function getCurrentBookSettingsIndex() As Integer
    If currentBookKey = "" Then
        getCurrentBookSettingsIndex = -1
        Exit Function
    End If
    
    Dim i As Integer
    
    For i = LBound(booksSettingsArr) To UBound(booksSettingsArr)
        Dim auxSs() As String
        
        auxSs = Split(booksSettingsArr(i), "|")
        
        If auxSs(0) = currentBookKey Then
            getCurrentBookSettingsIndex = i
            Exit Function
        End If
    Next i
    
    getCurrentBookSettingsIndex = -1
End Function

' Cambia el estado de la ventana según si la sesión está abierta o cerrada.
Private Sub toggleWindowState()
    If Not booksListView.SelectedItem Is Nothing Then
        booksListView.SelectedItem.Selected = False
        booksListView.SelectedItem = Nothing
        currentBookKey = ""
    End If

    signInOutBtn.Caption = IIf(currentSession = "", "Iniciar sesión", "Salir (" & currentSession & ")")
    readBtn.Enabled = currentSession <> ""
    toReadBtn.Enabled = currentSession <> ""
    nonLikedBtn.Enabled = currentSession <> ""
    genreListBtn.Enabled = currentSession <> ""
    favGenresBtn.Enabled = currentSession <> ""
    enableBookToolbarOption
    
    loadMEGACatalogue ""
End Sub

' Actualiza la ventana según los ajustes del libro seleccionado.
Private Sub refreshWindowSettings()
    If currentBookKey = "" Then
        Exit Sub
    End If
    
    Dim bSIdx As Integer, ss() As String
    bSIdx = getCurrentBookSettingsIndex
    
    ss = Split(booksSettingsArr(bSIdx), "|")
    
    ' Leído.
    markBookAsReadBtn.Caption = IIf(ss(1) = "1", "Quitar de leídos", "Marcar como leído")
    addBookToForReadBtn.Caption = IIf(ss(2) = "1", "Quitar de Por leer", "Marcar a Por leer")
    markBookAsNonLikedBtn.Caption = IIf(ss(3) = "1", "Quitar de No gustado", "Marcar como No gustado")
End Sub

' Habilita, si se cumple las condiciones, la sección de libros en la barra de herramientas.
Private Sub enableBookToolbarOption()
    bookToolbarItem.Enabled = currentSession <> "" And currentBookKey <> ""
End Sub

' Deshabilita la sección de libros en la barra de herramientas.
Private Sub disableBookToolbarOption()
    currentBookKey = ""
    bookToolbarItem.Enabled = False
End Sub

' Limpia la barra de búsqueda.
Private Sub clearSearchBar()
    searchBarTxt.Text = ""
End Sub

' Carga todo el catálogo de MEGA.
Private Sub loadMEGACatalogue(queryFilter As String)
    ' Inhabilita la sección de libro en la barra de herramientas.
    disableBookToolbarOption
    ' ** Carga el catálogo de MEGA.
    ' Lector de tuplas. Se usa para leer los registros.
    Dim rs As ADODB.Recordset
    Dim query As String ' Consulta SQL.
    
    ' Consulta SQL.
    query = "SELECT * FROM get_books(" & _
        IIf(currentSession = "", "NULL", "'" & currentSession & "'") & _
        ")"
    
    If queryFilter <> "" Then
        query = query & " WHERE " & queryFilter
    End If
    
    query = query & ";"
        
    ' Se usa Set para asignar objetos. Solo se usa = para tipos primitivos.
    Set rs = New ADODB.Recordset
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    ' Limpia la lista de libros.
    booksListView.ListItems.Clear
    
    If Not rs.EOF Then
        Dim item As ListItem
        
        ' Inicializa el arreglo de ajustes.
        ReDim booksSettingsArr(0 To 9)
        Dim count As Integer
        count = -1
        
        Do Until rs.EOF
            Set item = booksListView.ListItems.Add(, "b-" & rs!Idbook, rs!genre)
            item.SubItems(1) = rs!author
            item.SubItems(2) = rs!title
            item.SubItems(3) = rs!rate
            
            ' Agrega el libro al arreglo de ajustes.
            count = count + 1
            ' Incrementa la capacidad si es necesario.
            If count > UBound(booksSettingsArr) Then
                ReDim Preserve booksSettingsArr(0 To count + 9)
            End If
            
            booksSettingsArr(count) = CStr(rs!Idbook) & _
                "|" & CStr(rs!Has_read) & _
                "|" & CStr(rs!Wants_to_read) & _
                "|" & CStr(rs!Non_liked)
            
            ' Va a la siguiente tupla.
            rs.MoveNext
        Loop
    End If
    
    For i = LBound(booksSettingsArr) To UBound(booksSettingsArr)
        Debug.Print booksSettingsArr(i)
    Next i
    
    ' Cierra la sesión de la base de datos y libera recursos.
    rs.Close
    Set rs = Nothing
End Sub

Private Sub addBookBtn_Click()
    If currentSession = "" Then
        MsgBox "Debes iniciar sesión para usar esta opción", vbInformation, "Identifícate"
    ElseIf currentSession <> "admin" Then
        MsgBox "Solo el administrador puede usar esta opción", vbCritical, "Acceso denegado"
    Else
        AddOrUpdateBookForm.Show vbModal
    End If
End Sub

' Marca un libro como pendiente por leer.
Private Sub addBookToForReadBtn_Click()
    ' Obtiene los ajustes del libro.
    Dim bSIdx As Integer
    bSIdx = getCurrentBookSettings
    ' Obtiene el ajuste de leído.
    Dim ss() As String
    ss = Split(booksSettingsArr(bSIdx), "|")
    
    On Error GoTo ErrCatch
        ' Ejecuta la instrucción.
        Dim query As String
        query = "EXEC add_or_remove_from_for_read '" & _
            currentSession & "', " & _
            currentBookKey & ";"
        
        conn.Execute query, , adExecuteNoRecords
        ' Mensaje de confirmación.
        MsgBox "El libro fue " & IIf(ss(2) = "0", "agregado a", "quitado de") & " la lista para leer"
        ' Modifica el arreglo de ajustes del libro.
        ss(2) = IIf(ss(2) = "0", "1", "0") ' Por leer.
        ss(1) = IIf(ss(2) = "1", "0", ss(1)) ' Ya leído.
        ' Modifica el arreglo de los ajustes.
        booksSettingsArr(bSIdx) = Join(ss, "|")
        refreshWindowSettings
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "Ocurrió un problema: " & Err.Description, vbCritical, "Error"
    End If
End Sub

' Cuando la tabla recibe un clic.
Private Sub booksListView_Click()
    ' Extrae la llave del libro.
    currentBookKey = Split(booksListView.SelectedItem.Key, "-")(1)
    ' Actualiza la interfaz.
    refreshWindowSettings
    ' Habilita la sección libro de la barra de herramientas.
    enableBookToolbarOption
End Sub

' Elimina un libro.
Private Sub deleteBookBtn_Click()
    If currentSession <> "admin" Then
        MsgBox "Solo el administrador puede realizar esta acción", vbCritical, "Acceso denegado"
        Exit Sub
    End If

    Dim diagRes As Integer
    
    diagRes = MsgBox("Estás seguro que quieres borrar el libro seleccionado?", vbYesNo + vbExclamation)
    
    If diagRes = vbYes Then
        conn.Execute "DELETE FROM Book WHERE Idbook = " & currentBookKey
        MsgBox "El libro fue eliminado", vbInformation
        loadMEGACatalogue ""
    End If
End Sub

' Gestión de géneros favoritos.
Private Sub favGenresBtn_Click()
    GenreEditorForm.Show vbModal
End Sub

' Desde la carga del programa.
Private Sub Form_Load()
    ' Set crea una instancia de un objeto o para crear constantes.
    Set conn = New ADODB.Connection
    ' Cursores del lado del cliente (consulta de la BD desde el cliente.
    conn.CursorLocation = adUseClient
    
    ' Dim declara variables.
    Dim connString As String
    
    ' Cadena de conexión.
    connString = "Provider=SQLOLEDB.1;Data Source=.;" & _
        "Initial Catalog=megalibros;Integrated Security=SSPI;"

    ' Conecta la base de datos.
    conn.Open connString
    
    ' Carga el catálogo de MEGA.
    loadMEGACatalogue ""
End Sub

' Cuando se quiere obtener todo el catálogo.
Private Sub fullCatalogBtn_Click()
    clearSearchBar
    loadMEGACatalogue ""
End Sub

' Libros por géneros favoritos.
Private Sub genreListBtn_Click()
    loadMEGACatalogue "Idgenre IN (SELECT Idgenre FROM dbo.LIKED_GENRES WHERE Iduser = (SELECT u.Iduser FROM dbo.""User"" u WHERE u.username = '" & currentSession & "'))"
End Sub

' Marca un libro como no gustado.
Private Sub markBookAsNonLikedBtn_Click()
    ' Obtiene los ajustes del libro.
    Dim bSIdx As Integer
    bSIdx = getCurrentBookSettings
    ' Obtiene el ajuste de leído.
    Dim ss() As String
    ss = Split(booksSettingsArr(bSIdx), "|")
    
    On Error GoTo ErrCatch
        ' Ejecuta la instrucción.
        Dim query As String
        query = "EXEC add_or_remove_from_non_liked '" & _
            currentSession & "', " & _
            currentBookKey & ";"
        
        conn.Execute query, , adExecuteNoRecords
        ' Mensaje de confirmación.
        MsgBox "El libro fue " & IIf(ss(3) = "0", "agregado a", "quitado de") & " de no gustados"
        ' Modifica el arreglo de ajustes del libro.
        ss(3) = IIf(ss(3) = "0", "1", "0")
        ' Modifica el arreglo de los ajustes.
        booksSettingsArr(bSIdx) = Join(ss, "|")
        refreshWindowSettings
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "Ocurrió un problema: " & Err.Description, vbCritical, "Error"
    End If
End Sub

' Marca un libro como leído.
Private Sub markBookAsReadBtn_Click()
    ' Obtiene los ajustes del libro.
    Dim bSIdx As Integer
    bSIdx = getCurrentBookSettings
    ' Obtiene el ajuste de leído.
    Dim ss() As String
    ss = Split(booksSettingsArr(bSIdx), "|")
    
    On Error GoTo ErrCatch
        ' Ejecuta la instrucción.
        Dim query As String
        query = "EXEC add_or_remove_from_as_read '" & _
            currentSession & "', " & _
            currentBookKey & ";"
        
        conn.Execute query, , adExecuteNoRecords
        ' Mensaje de confirmación.
        MsgBox "El libro fue " & IIf(ss(1) = "0", "agregado a", "quitado de") & " la lista de leídos"
        ' Modifica el arreglo de ajustes del libro.
        ss(1) = IIf(ss(1) = "0", "1", "0") ' Por leer.
        ss(2) = IIf(ss(1) = "1", "0", ss(2)) ' Ya leído.
        ' Modifica el arreglo de los ajustes.
        booksSettingsArr(bSIdx) = Join(ss, "|")
        refreshWindowSettings
ErrCatch:
    If Err.Description <> "" Then
        MsgBox "Ocurrió un problema: " & Err.Description, vbCritical, "Error"
    End If
End Sub

' Cuando se quieren ver los libros no gustados.
Private Sub nonLikedBtn_Click()
    loadMEGACatalogue "Non_liked = 1"
End Sub

' Cuando se quieren ver los libros leídos
Private Sub readBtn_Click()
    loadMEGACatalogue "Has_read = 1"
End Sub

' Cuando se quiere obtener lo recomendado
Private Sub recommendedBtn_Click()
    clearSearchBar
    loadMEGACatalogue "Recommended = 1"
End Sub

' Cuando se hace clic en buscar.
Private Sub searchBtn_Click()
    Dim txtValue As String
    
    txtValue = searchBarTxt.Text

    If txtValue = "" Then
        loadMEGACatalogue ""
    Else
        loadMEGACatalogue "Title LIKE '%" & txtValue & "%' OR Author LIKE '%" & txtValue & "%'"
    End If
End Sub

' Cuando el usuario quiere iniciar sesión.
Private Sub signInOutBtn_Click()
    If currentSession = "" Then
       LoginForm.Show vbModal
    
        If currentSession <> "" Then
            ' La sesión se inició correctamente.
            toggleWindowState
        End If
    Else
        Dim result As Integer
        result = MsgBox("¿Estás seguro que deseas salir?", vbYesNo + vbQuestion, "Cerrar sesión")
        
        If result = vbYes Then
            currentSession = ""
            toggleWindowState
            MsgBox "Cerraste tu sesión", vbInformation, "Sesión cerrada"
        End If
    End If
End Sub

' Cuando se quieren obtener los libros por leer.
Private Sub toReadBtn_Click()
    loadMEGACatalogue "Wants_to_read = 1"
End Sub

' Actualiza el contenido de un libro.
Private Sub updateBookBtn_Click()
    If currentSession <> "admin" Then
        MsgBox "Solo el administrador puede usar esta opción.", vbCritical, "Acceso denegado"
        Exit Sub
    End If

    AddOrUpdateBookForm.bookId = currentBookKey
    AddOrUpdateBookForm.Show vbModal
End Sub
