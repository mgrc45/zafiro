VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Suite Zafiro V. 2.11"
   ClientHeight    =   8085
   ClientLeft      =   990
   ClientTop       =   1995
   ClientWidth     =   11985
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":591A
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fon 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      Picture         =   "MDIForm1.frx":153A0
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   799
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      Begin VB.TextBox ruta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8280
         TabIndex        =   4
         Top             =   3960
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   3240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   4080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Lb_Exp_code 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   720
         TabIndex        =   9
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label Lb_Admin_Code 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   720
         TabIndex        =   8
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Lb_Internet 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   720
         TabIndex        =   7
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label p_tipo 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   8280
         TabIndex        =   3
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label servicep 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   8280
         TabIndex        =   2
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label sistema 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8280
         TabIndex        =   1
         Top             =   2880
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dir As String
Const ProgramDir As String = "Software\DesarrolloDigital\SuiteZafiro\V2.11"

Private Sub fon_Click()

End Sub

Private Sub Lb_Admin_Code_Click()
Admin_Code.Show vbModal
End Sub

Private Sub Lb_Exp_code_Click()
Exp_code.Show vbModal
End Sub

Private Sub Lb_Internet_Click()
Internet.Show vbModal
End Sub
Private Sub MDIForm_Load()
dir = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
sistema.Caption = Reg_Valor(HKEY_LOCAL_MACHINE, dir, "ProductName")
servicep.Caption = Reg_Valor(HKEY_LOCAL_MACHINE, dir, "CSDVersion")
p_tipo.Caption = Reg_Valor(HKEY_LOCAL_MACHINE, dir, "CurrentType")
ruta.Text = Reg_Valor(HKEY_LOCAL_MACHINE, dir, "SystemRoot")
ProgressBar1.Value = 75
ProgressBar2.Value = 0

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

