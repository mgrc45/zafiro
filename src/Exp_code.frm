VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Exp_code 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Explorador de codigos"
   ClientHeight    =   8070
   ClientLeft      =   15
   ClientTop       =   1635
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Exp_code.frx":0000
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Exp_code.frx":A92F
      Left            =   480
      List            =   "Exp_code.frx":A931
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   3840
      Width           =   5415
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Exp_code.frx":A933
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Favoritos"
      Height          =   1335
      Left            =   6120
      TabIndex        =   2
      Top             =   6240
      Width           =   5415
      Begin VB.CommandButton del 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Add 
         Caption         =   "Añadir"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox Fav 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4425
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.TextBox txt_dir 
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
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   10935
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   2295
      Left            =   6120
      TabIndex        =   7
      Top             =   3840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga :           %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "Exp_code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nodX As Node
Private Root, dir As String
Private cnt, fin, filtro, numIndex As Integer
Dim papa As String
Dim nomb, tipo, valor As String
Private Sub Add_Click()
nomb = InputBox("Que nombre desea poner a este vinculo?", "Nombre")
dir = Txt_dir.Text
Reg_Crear HKEY_LOCAL_MACHINE, "Software\DesarrolloDigital\SuiteZafiro\V2.11", nomb, dir
FillFav
End Sub

Private Sub Combo1_Click()
    Root = Combo1.List(Combo1.ListIndex)
    Txt_dir.Text = " "
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
    FillNodos Root, "", 0

End Sub

Private Sub del_Click()
nomb = Fav.Text
Reg_Borrar HKEY_LOCAL_MACHINE, "Software\DesarrolloDigital\SuiteZafiro\V2.11", nomb
FillFav
End Sub

Private Sub Fav_Click()
nomb = Fav.Text
dir = Reg_Valor(HKEY_LOCAL_MACHINE, "Software\DesarrolloDigital\SuiteZafiro\V2.11", nomb)
Ir dir
End Sub


Private Sub Form_Load()
With TreeView1
        .Indentation = 10 ' Separacion entre nodo y nodo (expresada en pixeles)
        .ImageList = ImageList1 ' Para conectarlo con mi lista de imagenes
        .Refresh
        .Nodes.Clear
   End With

'Buscando los roots existentes
Combo1.Clear
Txt_dir.Text = " "
If Exist(HKEY_CLASSES_ROOT, "", "") Then Combo1.AddItem "HKEY_CLASSES_ROOT"
If Exist(HKEY_CURRENT_USER, "", "") Then Combo1.AddItem "HKEY_CURRENT_USER"
If Exist(HKEY_LOCAL_MACHINE, "", "") Then Combo1.AddItem "HKEY_LOCAL_MACHINE"
If Exist(HKEY_USERS, "", "") Then Combo1.AddItem "HKEY_USERS"
If Exist(HKEY_PERFORMANCE_DATA, "", "") Then Combo1.AddItem "HKEY_PERFORMANCE_DATA"
If Exist(HKEY_CURRENT_CONFIG, "", "") Then Combo1.AddItem "HKEY_CURRENT_CONFIG"
If Exist(HKEY_DYN_DATA, "", "") Then Combo1.AddItem "HKEY_DYN_DATA"
FillFav
'Ir "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction"
Ir "HKEY_CURRENT_USER\Control Panel\desktop"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Txt_dir.Text = Node.FullPath
TreeView2.Nodes.Clear

If Not Node.Text = Node.Root Then

    If Not Node.Parent = Node.Root Then 'Nivel 2
        'Con esto filtro la cadena de direccion
        filtro = InStr(1, Node.FullPath, "\", vbTextCompare)
        filtro = Len(Node.FullPath) - filtro
        dir = Right(Node.FullPath, filtro)
        FillNodos2 Node.Root, dir, 2
    Else 'Nivel 1
        FillNodos2 Node.Root, Node.Text, 1
    End If

Else 'Nivel 0
    FillNodos2 Node.Root, "", 0
End If


If Node.Children = 0 And Not Node.Text = Node.Root Then
    papa = Node.Key
    If Not Node.Parent = Node.Root Then 'Nivel 2
        'Con esto filtro la cadena de direccion
        filtro = InStr(1, Node.FullPath, "\", vbTextCompare)
        filtro = Len(Node.FullPath) - filtro
        dir = Right(Node.FullPath, filtro)
        FillNodos Node.Root, dir, 2
    Else 'Nivel 1
        FillNodos Node.Root, Node.Text, 1
    End If
    
End If
'MsgBox Node.Key, vbInformation, "Root"
End Sub
Public Function Ir(ByVal Direccion As String)
Direccion = Direccion & "\"
Dim Direccion2, DireccionFull, x As String
Dim inum, cnt2 As Integer
Dim termina As Boolean
termina = False

filtro = InStr(1, Direccion, "\", vbTextCompare)
Root = Left(Direccion, filtro - 1) ' obtengo el root
Direccion = Right(Direccion, Len(Direccion) - filtro) ' asi obtengo la direccion

fin = Combo1.ListCount - 1
For cnt = 0 To fin
    If Combo1.List(cnt) = Root Then
        Combo1.Text = Combo1.List(cnt)
        If Len(Direccion) > 0 Then
            cnt2 = 0
            Do While termina = False
                filtro = InStr(1, Direccion, "\", vbTextCompare)
                If Len(Direccion) > 0 Then Direccion2 = Left$(Direccion, filtro - 1) ' obtengo los nombres
                
                If Not cnt2 = 0 Then
                    FillNodos Root, DireccionFull, 1, Direccion2
                    papa = papa & "\" & numIndex
                    DireccionFull = DireccionFull & "\" & Direccion2
                    Else
                    DireccionFull = Direccion2
                    FillNodos Root, DireccionFull, 0, Direccion2 'llena la clave principal
                    papa = "\" & numIndex
                    End If
                Direccion = Right(Direccion, Len(Direccion) - filtro) ' asi obtengo la direccion
                If Len(Direccion) < 1 Then
                    termina = True
                    FillNodos Root, DireccionFull, 1
                End If
                cnt2 = cnt2 + 1
            Loop
        End If
    End If
Next cnt
TreeView1.Nodes(TreeView1.Nodes.Count).EnsureVisible


        'TreeView1.Nodes.Item(1).Expanded = True
        'TreeView1.Nodes(2).EnsureVisible
        'TreeView1.SelectedItem.EnsureVisible
        'TreeView1.Nodes.Item(2).Selected = True
        ' x = TreeView1.SelectedItem.Key



'If encontrado = False Then
'MsgBox "Es posible que la ruta que usted busca ya no exista o la haya escrito de manera incorrecta", vbExclamation, "Error"
'End If
End Function
Public Function FillNodos(ByVal Root As String, ByVal Direccion As String, ByVal Nivel As Integer, Optional ByVal comp As String)
'Funcion para cargar las carpetas del treeview1

If Nivel = 0 Then ' Carga los nodos principales
    fin = Enumera(cad(Root), "", True) - 1
    TreeView1.Nodes.Clear
    Set nodX = TreeView1.Nodes.Add(, , , Root, 1) 'nodo principal
    For cnt = 0 To fin
        nomb = Key_nombre(cad(Root), "", cnt)
        If nomb = comp Then numIndex = cnt
        Set nodX = TreeView1.Nodes.Add(1, tvwChild, "\" & cnt, nomb, 1, 1)
        Label2.Caption = Round((cnt * 100) / fin, 2)
        DoEvents 'Con esto le doy tiempo al sistema operativo de hacer sus cosas
    Next cnt
End If

If Nivel = 1 Or Nivel = 2 Then
        fin = Enumera(cad(Root), Direccion, True) - 1
        For cnt = 0 To fin
            nomb = Key_nombre(cad(Root), Direccion, cnt)
            If nomb = comp Then numIndex = cnt
            Set nodX = TreeView1.Nodes.Add(papa, tvwChild, papa & "\" & cnt, nomb, 1, 1)
            DoEvents 'Con esto le doy tiempo al sistema operativo de hacer sus cosas
        Next cnt
End If

End Function
Public Function FillNodos2(ByVal Root As String, ByVal Direccion As String, ByVal Nivel As Integer)
'Funcion creada para los registros

If Nivel = 0 Then ' Carga los nodos principales
    fin = Enumera(cad(Root), "", False) - 1
    For cnt = 0 To fin
        Reg_Zafiro cad(Root), "", cnt
        DoEvents 'Con esto le doy tiempo al sistema operativo de hacer sus cosas
    Next cnt
End If

If Nivel = 1 Or Nivel = 2 Then
    fin = Enumera(cad(Root), Direccion, False) - 1
    For cnt = 0 To fin
        Reg_Zafiro cad(Root), Direccion, cnt
        DoEvents 'Con esto le doy tiempo al sistema operativo de hacer sus cosas
    Next cnt
End If


End Function
Public Function FillFav()
Fav.Clear
fin = Enumera(HKEY_LOCAL_MACHINE, "Software\DesarrolloDigital\SuiteZafiro\V2.11", False)
For cnt = 0 To fin
    nomb = Reg_nombre(HKEY_LOCAL_MACHINE, "Software\DesarrolloDigital\SuiteZafiro\V2.11", cnt)
    Fav.AddItem nomb
Next cnt
End Function
