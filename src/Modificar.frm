VERSION 5.00
Begin VB.Form Modificar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar"
   ClientHeight    =   5550
   ClientLeft      =   1200
   ClientTop       =   2055
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_dir 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Modificar.frx":0000
      Top             =   720
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Tipo: "
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   5175
      Begin VB.OptionButton Opt3 
         Caption         =   "REG_SZ (Cadena de caracteres)"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   4215
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "REG_DWORD (Datos numericos)"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   4215
      End
   End
   Begin VB.TextBox tdata 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton C2 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton C1 
      Caption         =   "Aplicar"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   1320
      X2              =   5400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label L2 
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   10
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Actual: "
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label L2 
      Caption         =   "Aqui va el valor actual"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label L2 
      Caption         =   "Aqui va el nombre"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo Valor: "
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valor Actual: "
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Direccion: "
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Lb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nombre: "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Modificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Myroot As String
Dim Mydir As String
Dim cnt, fin As Integer
Private Sub C1_Click()
Mydir = Txt_dir.Text
cnt = InStr(1, Mydir, "\", vbTextCompare)
Myroot = Left(Mydir, cnt - 1) ' obtengo el root
cnt = Len(Mydir) - cnt
Mydir = Right(Mydir, cnt) ' asi obtengo la direccion

If Len(tdata.Text) > 0 Then
    If Opt2.Value = True Then 'Reg_dword
            If IsNumeric(tdata.Text) = True Then
                If val(tdata.Text) > 9999 Then
                MsgBox "El valor no puede superar las 4 cifras", vbExclamation, "Error"
                Else
                Reg_Borrar cad(Myroot), Mydir, L2(0).Caption
                Reg_Modificar cad(Myroot), Mydir, L2(0).Caption, tdata.Text, 4
                'corre bien
                End If
            Else
            MsgBox "El valor no corresponde al tipo, accion cancelada", vbExclamation, "Error"
            End If
    End If

    If Opt3.Value = True Then 'Reg_sz
        If Not Len(tdata.Text) > 254 Then
        Reg_Borrar cad(Myroot), Mydir, L2(0).Caption
        Reg_Modificar cad(Myroot), Mydir, L2(0).Caption, tdata.Text, 1
        'Correbien
        Else
        MsgBox "La cadena de datos es demasiado grande, accion cancelada", vbExclamation, "Error"
        End If
    End If
Me.Hide
Else
MsgBox "No se puede dejar el campo vacio", vbExclamation, "Error"
End If

End Sub
Private Sub C2_Click()
Me.Hide
End Sub


Private Sub Form_Load()
tdata.Text = ""
End Sub
