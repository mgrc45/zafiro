VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Admin_Code 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrador de codigos"
   ClientHeight    =   8100
   ClientLeft      =   3825
   ClientTop       =   3120
   ClientWidth     =   12000
   Icon            =   "Admin_code.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Admin_code.frx":591A
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox arch_temp 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   240
      Pattern         =   "*.pf"
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Txt_mod 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Text            =   "000000001"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton CB_Mod 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton CB_des 
      Caption         =   "Desactivar"
      Height          =   495
      Left            =   6840
      TabIndex        =   12
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton CB_Act 
      Caption         =   "Activar"
      Height          =   495
      Left            =   6840
      TabIndex        =   11
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mostrar:"
      Height          =   1695
      Left            =   9480
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
      Begin VB.OptionButton Opt_Sn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sin clasificación"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Opt_Blo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bloqueos"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Opt_Des 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Desempeño"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.TextBox Txt_ruta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   10935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   600
      TabIndex        =   16
      Top             =   5040
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   250.016
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Button          =   -1  'True
            ColumnWidth     =   109.984
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   6
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DWORD"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "NotRun"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "000000001"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "DWORD"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NotRon"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
End
Attribute VB_Name = "Admin_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tabla As String
Dim campo(0 To 16) As String
Dim total As String
Dim cnt As Integer

Private Sub CB_Act_Click()
    Reg_Borrar cad(campo(1)), Cad_total, campo(9)
    Reg_Crear cad(campo(1)), Cad_total, campo(9), campo(12)
    Unload Me
End Sub

Private Sub CB_des_Click()
    Reg_Borrar cad(campo(1)), Cad_total, campo(9)
    Reg_Crear cad(campo(1)), Cad_total, campo(9), campo(13)
    Unload Me
End Sub

Private Sub CB_Mod_Click()
    Modificar.L2(0).Caption = Label4.Caption 'Nombre
    Modificar.Txt_dir.Text = Txt_ruta.Text  'Direccion
    Modificar.L2(2).Caption = Txt_mod.Text 'Valor Actual
    Modificar.L2(3).Caption = Label5.Caption ' Tipo
    Modificar.Show vbModal
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then MostarDatos
End Sub

Private Sub Form_Load()
    CargarLista
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim fso As Variant
    arch_temp.Path = "C:\WINDOWS\Prefetch\"
    arch_temp.Pattern = "*.pf"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If arch_temp.ListCount > 0 Then fso.DeleteFile ("C:\WINDOWS\Prefetch\*.pf")
    'necesito crear un boton para abrir C:\WINDOWS\SYSTEM32\schtasks.exe
End Sub
Private Sub MostarDatos()
    Conexion True, Tabla
    
        Rec.Move DataGrid1.Columns(0).Text, 0
        LoadBuffer 'Sub proceso
        If DataGrid1.Columns(2).Text = "Activado" Or DataGrid1.Columns(2).Text = "Modificado" Then
        CB_Act.Enabled = False
        Else: CB_Act.Enabled = True: End If
        
        If DataGrid1.Columns(2).Text = "Desactivado" Or DataGrid1.Columns(2).Text = "No existe" Then
        CB_des.Enabled = False
        Else: CB_des.Enabled = True: End If
        
        total = Cad_total()
    
        Txt_ruta.Text = campo(1) & "\" & total
        Label1.Caption = campo(9) 'Nombre del registro
        Label4.Caption = campo(9) 'Nombre del registro
        Label2.Caption = campo(10) 'Tipo de registro
        Label3.Caption = campo(14) 'Valor Recomendado
        Label7.Caption = campo(16) ' Informacion
        Label5.Caption = " " ' Tipo de valor
        Label5.Caption = Reg_Tipo(cad(campo(1)), total, campo(9)) ' Tipo de valor
        Txt_mod.Text = " " 'Valor del campo
        Txt_mod.Text = Reg_Valor(cad(campo(1)), total, campo(9)) ' Valor del campo
        
    Conexion False, Tabla
End Sub
Private Sub Opt_Blo_Click()
    CargarLista
End Sub
Private Sub Opt_Des_Click()
    CargarLista
End Sub
Private Sub Opt_Sn_Click()
    CargarLista
End Sub

Private Sub CargarLista()
    Dim valReg, A, D As String
    Dim subCnt As Integer
    subCnt = 0
    Dim rstemp As New ADODB.Recordset
        Set rstemp = Nothing
        rstemp.Fields.Append "Id", adInteger, 3
        rstemp.Fields.Append "Nombre", adVarWChar, 50
        rstemp.Fields.Append "Estado", adVarWChar, 50
        rstemp.Open
        
        DataGrid1.ClearFields
    Set DataGrid1.DataSource = rstemp
        DataGrid1.Columns(0).Width = 25
        DataGrid1.Columns(1).Width = 265
        DataGrid1.Columns(2).Width = 70
        DataGrid1.Columns(0).Alignment = dbgCenter
    
        If Opt_Des Then Tabla = "Visual_Desempeño"
        If Opt_Sn Then Tabla = "Visual_Sin_clasificar"
        If Opt_Blo Then Tabla = "Visual_Bloqueos"
    
    Conexion True, Tabla
    
    With Rec
    Do While Not .EOF 'Cargo la lista
        LoadBuffer 'Sub proceso
        total = Cad_total()
        rstemp.AddNew
            rstemp(0) = subCnt: subCnt = subCnt + 1
            rstemp(1) = campo(0)
                
        If Exist(cad(campo(1)), total, campo(9)) Then
            If campo(12) <> " " Then A = campo(12) Else A = " "
            If campo(13) <> " " Then D = campo(13) Else D = " "
            valReg = Reg_Valor(cad(campo(1)), total, campo(9))
            Select Case valReg
            Case A: rstemp(2) = "Activado"
            Case D: rstemp(2) = "Desactivado"
            Case Else: rstemp(2) = "Modificado"
            End Select
        Else: rstemp(2) = "No existe"
        End If
        
        rstemp.Update
        .MoveNext
        DoEvents
    Loop
    End With

    Conexion False, Tabla
End Sub


Private Sub LoadBuffer()
    For cnt = 0 To 16
    If Len(Rec.Fields(cnt)) > 0 Then
    campo(cnt) = Rec.Fields(cnt)
    Else: campo(cnt) = " ": End If
    Next cnt
End Sub

Private Function Cad_total() As String
    'Creo las variables: Temporales
    Dim subTotal, totalTmp As String
    'Inicializo las variables: En blanco
    subTotal = " ": totalTmp = " "
    If campo(2) <> " " Then
        totalTmp = campo(2)
        For cnt = 3 To 7
        If campo(cnt) <> " " Then
        subTotal = "\" & campo(cnt)
        totalTmp = totalTmp + subTotal
        End If
        Next cnt
    End If
    Cad_total = totalTmp
End Function





