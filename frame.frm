VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frame 
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_favoritos 
      Caption         =   "Libros Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton btn_generos_favoritos 
      Caption         =   "Generos Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton btn_quiero 
      Caption         =   "Quiero Leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton btn_eliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton btn_modificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton btn_agregar 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton btn_no_gustar 
         Caption         =   "No te Gustaron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton btn_leistes 
         Caption         =   "Ya Leistes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton btn_catalogo 
         Caption         =   "Catalogo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView list_libros 
      Height          =   5535
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btn_agregar_Click()
   frmLibro.EditandoID = 0
   frmLibro.Show vbModal
End Sub

Private Sub btn_catalogo_Click()
     CargarLibros ""

End Sub

Private Sub CargarLibros(filtroSQL As String)
    On Error GoTo ErrorHandler

    Dim rs As ADODB.Recordset
    Dim sql As String

    ' Incluir LibroID en la consulta
    sql = "SELECT L.LibroID, L.Titulo, L.Autor, G.Nombre AS Genero, " & _
          "L.Calificacion, L.Prestado, L.PrestadoA, L.FechaPrestamo " & _
          "FROM Libros L INNER JOIN Generos G ON L.GeneroID = G.GeneroID"

    If filtroSQL <> "" Then
        sql = sql & " " & filtroSQL
    End If

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    list_libros.ListItems.Clear

    Do While Not rs.EOF
        Dim prestadoTexto As String
        prestadoTexto = "No"
        If Not IsNull(rs("Prestado")) And CBool(rs("Prestado")) Then
            prestadoTexto = "Sí"
        End If

        Dim prestadoA As String
        prestadoA = IIf(IsNull(rs("PrestadoA")), "", rs("PrestadoA"))

        Dim fechaPrestamo As String
        fechaPrestamo = ""
        If Not IsNull(rs("FechaPrestamo")) Then
            fechaPrestamo = Format(rs("FechaPrestamo"), "yyyy-mm-dd")
        End If

        Dim item As ListItem
        Set item = list_libros.ListItems.Add(, , rs("Titulo"))
        item.Tag = rs!libroID ' Guardar el ID del libro
        item.SubItems(1) = rs("Autor")
        item.SubItems(2) = rs("Genero")
        item.SubItems(3) = rs("Calificacion")
        item.SubItems(4) = prestadoTexto
        item.SubItems(5) = prestadoA
        item.SubItems(6) = fechaPrestamo

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar libros: " & Err.Description, vbCritical
End Sub



Private Sub btn_eliminar_Click()
    On Error GoTo ErrorHandler

    ' Verifica si hay un libro seleccionado
    If list_libros.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro para eliminar.", vbExclamation
        Exit Sub
    End If

    ' Confirmar eliminación
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Estás seguro de que deseas eliminar este libro?", vbYesNo + vbQuestion, "Confirmar eliminación")
    If respuesta = vbNo Then Exit Sub

    ' Obtener el ID del libro desde el Tag
    Dim libroID As Integer
    libroID = list_libros.SelectedItem.Tag

    ' Ejecutar consulta de eliminación
    Dim sql As String
    sql = "DELETE FROM Libros WHERE LibroID = " & libroID

    conn.Execute sql

    ' Recargar lista de libros
    CargarLibros ""

    MsgBox "Libro eliminado correctamente.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error al eliminar el libro: " & Err.Description, vbCritical
End Sub


Private Sub btn_favoritos_Click()
  CargarLibros "WHERE L.Calificacion >= 8"
End Sub

Private Sub btn_generos_favoritos_Click()
  CargarLibros "WHERE G.EsFavorito = 1"

End Sub

Private Sub btn_leistes_Click()
 CargarLibros "WHERE L.Leido = 1"
End Sub

Private Sub btn_modificar_Click()
    ' Verifica si hay un libro seleccionado
    If list_libros.SelectedItem Is Nothing Then
        MsgBox "Selecciona un libro para modificar.", vbExclamation
        Exit Sub
    End If

    ' Obtener el ID del libro desde el Tag (debes haberlo guardado ahí al cargar)
    Dim libroID As Integer
    libroID = list_libros.SelectedItem.Tag

    ' Pasar el ID al formulario frmLibro
    frmLibro.EditandoID = libroID
    frmLibro.Show vbModal

    ' Recargar libros después de cerrar el formulario
    CargarLibros ""
End Sub


Private Sub btn_no_gustar_Click()
  CargarLibros "WHERE L.Calificacion <= 2"

End Sub

Private Sub btn_quiero_Click()
 CargarLibros "WHERE L.PorLeer=1"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Set conn = New ADODB.Connection
    conn.CursorLocation = adUseClient

    Dim connString As String
    connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-9OMEEHK\BDD23A;" & _
                 "Initial Catalog=BibliotecaDB;Integrated Security=SSPI;"

    conn.Open connString

    ' Configurar columnas del ListView
    With list_libros
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ListItems.Clear

        .ColumnHeaders.Add , , "Titulo", 2000
        .ColumnHeaders.Add , , "Autor", 2000
        .ColumnHeaders.Add , , "Genero", 1500
        .ColumnHeaders.Add , , "Calificación", 1000
        .ColumnHeaders.Add , , "Prestado", 1000
        .ColumnHeaders.Add , , "Prestado a", 2000
        .ColumnHeaders.Add , , "Fecha préstamo", 2000
    End With   ' ? ¡Aquí estaba el error!

    Exit Sub

ErrorHandler:
    MsgBox "Error en la conexión: " & Err.Description, vbCritical
End Sub


