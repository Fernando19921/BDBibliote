VERSION 5.00
Begin VB.Form frmLibro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   15
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   14
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox txtPrestadoA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   7560
      Width           =   4095
   End
   Begin VB.CheckBox chkPrestado 
      Caption         =   "Prestado Actualmente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CheckBox chkRecomendado 
      Caption         =   "Recomendado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CheckBox chkPorLeer 
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
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CheckBox chkLeido 
      Caption         =   "Ya Leido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtCalificacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
      Width           =   615
   End
   Begin VB.ComboBox cboGenero 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtAutor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Prestado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Calificacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Genero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Autor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Titulo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   870
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditandoID As Integer

Private Sub chkLeido_Click()
    If chkLeido.Value = 1 Then
        chkPorLeer.Value = 0
        txtCalificacion.Enabled = True
    Else
        txtCalificacion.Enabled = False
    End If
End Sub

Private Sub chkPorLeer_Click()
    If chkPorLeer.Value = 1 Then
        chkLeido.Value = 0
    End If
End Sub

Private Sub chkPrestado_Click()
    If chkPrestado.Value = 1 Then
        txtPrestadoA.Enabled = True
    Else
        txtPrestadoA.Enabled = False
        txtPrestadoA.Text = ""
    End If
End Sub

Private Sub cmdAceptar_Click()
    ' Validaciones básicas
    If Trim(txtTitulo.Text) = "" Or Trim(txtAutor.Text) = "" Then
        MsgBox "El título y el autor son obligatorios.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    If cboGenero.ListIndex = -1 Then
        MsgBox "Por favor selecciona un género.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    ' Validar calificación si está marcado como leído
    If chkLeido.Value = 1 Then
        If Trim(txtCalificacion.Text) = "" Then
            MsgBox "Debes ingresar una calificación si el libro está marcado como leído.", vbExclamation, "Datos incompletos"
            Exit Sub
        End If
    End If

    ' Validar rango de calificación
    Dim calif As Variant
    If Trim(txtCalificacion.Text) <> "" Then
        calif = Val(txtCalificacion.Text)
        If calif < 1 Or calif > 10 Then
            MsgBox "La calificación debe ser del 1 al 10.", vbExclamation, "Datos incorrectos"
            Exit Sub
        End If
    Else
        calif = 0
    End If

    ' Validar campo PrestadoA si está marcado como prestado
    If chkPrestado.Value = 1 Then
        If Trim(txtPrestadoA.Text) = "" Then
            MsgBox "Debes indicar a quién se prestó el libro.", vbExclamation, "Datos incompletos"
            Exit Sub
        End If
    End If

    ' Obtener ID del género
    Dim generoID As Integer
    generoID = cboGenero.ItemData(cboGenero.ListIndex)

    Dim prestadoA As String
    prestadoA = Replace(txtPrestadoA.Text, "'", "''")
    If prestadoA = "" Then prestadoA = "Desconocido"

    Dim fechaPrestamo As String
    If chkPrestado.Value = 1 Then
        fechaPrestamo = "'" & Format(Date, "yyyy-mm-dd") & "'"
    Else
        fechaPrestamo = "'1900-01-01'"  ' Fecha por defecto aunque no la uses visualmente
    End If

    Dim sql As String

    ' ?? Aquí está la diferencia: INSERT o UPDATE
    If EditandoID = 0 Then
        ' INSERT
        sql = "INSERT INTO Libros (Titulo, Autor, GeneroID, Calificacion, Leido, PorLeer, Recomendado, Prestado, PrestadoA, FechaPrestamo) VALUES (" & _
              "'" & Replace(txtTitulo.Text, "'", "''") & "', " & _
              "'" & Replace(txtAutor.Text, "'", "''") & "', " & _
              generoID & ", " & _
              calif & ", " & _
              chkLeido.Value & ", " & chkPorLeer.Value & ", " & chkRecomendado.Value & ", " & _
              chkPrestado.Value & ", '" & prestadoA & "', " & fechaPrestamo & ")"
    Else
        ' UPDATE
        sql = "UPDATE Libros SET " & _
              "Titulo = '" & Replace(txtTitulo.Text, "'", "''") & "', " & _
              "Autor = '" & Replace(txtAutor.Text, "'", "''") & "', " & _
              "GeneroID = " & generoID & ", " & _
              "Calificacion = " & calif & ", " & _
              "Leido = " & chkLeido.Value & ", " & _
              "PorLeer = " & chkPorLeer.Value & ", " & _
              "Recomendado = " & chkRecomendado.Value & ", " & _
              "Prestado = " & chkPrestado.Value & ", " & _
              "PrestadoA = '" & prestadoA & "', " & _
              "FechaPrestamo = " & fechaPrestamo & " " & _
              "WHERE LibroID = " & EditandoID
    End If

    conn.Execute sql

    MsgBox "Libro guardado correctamente.", vbInformation
    Unload Me
End Sub




Private Sub cmdCancelar_Click()
  Unload Me
End Sub
Private Sub CargarDatosLibro(ByVal libroID As Integer)
    On Error GoTo ErrorHandler

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT * FROM Libros WHERE LibroID = " & libroID

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        txtTitulo.Text = rs("Titulo")
        txtAutor.Text = rs("Autor")

        ' Seleccionar el género en el ComboBox
        Dim i As Integer
        For i = 0 To cboGenero.ListCount - 1
            If cboGenero.ItemData(i) = rs("GeneroID") Then
                cboGenero.ListIndex = i
                Exit For
            End If
        Next i

        txtCalificacion.Text = rs("Calificacion")

        chkLeido.Value = IIf(rs("Leido") = True, 1, 0)
        chkPorLeer.Value = IIf(rs("PorLeer") = True, 1, 0)
        chkRecomendado.Value = IIf(rs("Recomendado") = True, 1, 0)
        chkPrestado.Value = IIf(rs("Prestado") = True, 1, 0)

        txtPrestadoA.Text = IIf(IsNull(rs("PrestadoA")), "", rs("PrestadoA"))
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al cargar datos del libro: " & Err.Description, vbCritical
End Sub



Private Sub Form_Load()
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset

    ' Consultar géneros
    rsG.Open "SELECT GeneroID, Nombre FROM Generos ORDER BY Nombre", conn, adOpenStatic, adLockReadOnly

    ' Limpiar ComboBox
    cboGenero.Clear

    ' Llenar ComboBox con los nombres y guardar el ID en ItemData
    Do Until rsG.EOF
        cboGenero.AddItem rsG!Nombre
        cboGenero.ItemData(cboGenero.NewIndex) = rsG!generoID
        rsG.MoveNext
    Loop

    rsG.Close: Set rsG = Nothing

    If EditandoID = 0 Then
        ' Modo agregar, limpiar campos
        txtTitulo.Text = ""
        txtAutor.Text = ""
        cboGenero.ListIndex = -1          ' Sin selección
        txtCalificacion.Text = ""

        chkLeido.Value = 0
        chkPorLeer.Value = 0
        chkRecomendado.Value = 0
        chkPrestado.Value = 0

        txtCalificacion.Enabled = False
        txtPrestadoA.Text = ""
        txtPrestadoA.Enabled = False

        Me.Caption = "Agregar Libro"
    Else
        ' Modo editar, cargar datos del libro
        CargarDatosLibro EditandoID
        Me.Caption = "Modificar Libro"
    End If
End Sub


