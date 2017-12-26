VERSION 5.00
Begin VB.Form Internet 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Internet"
   ClientHeight    =   8070
   ClientLeft      =   1995
   ClientTop       =   2175
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Internet.frx":0000
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Original_values 
      Caption         =   "RESTAURAR VALORES ORIGINALES"
      Height          =   735
      Left            =   9840
      TabIndex        =   3
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton C2 
      Caption         =   "..::OPTIMIZAR::.."
      Height          =   735
      Left            =   9840
      TabIndex        =   26
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton DNS1 
      Caption         =   "DNS TELMEX"
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox v_16 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   24
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Aplicar 
      Caption         =   "APLICAR CAMBIOS"
      Height          =   735
      Left            =   8040
      TabIndex        =   23
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Test3 
      Caption         =   "Prueba 3"
      Height          =   375
      Left            =   10560
      TabIndex        =   22
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Test2 
      Caption         =   "Prueba 2"
      Height          =   375
      Left            =   9120
      TabIndex        =   21
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox v_2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox v_1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   19
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox v_3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Tag             =   "Seg/hops"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Test1 
      Caption         =   "Prueba 1"
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Ima_Web 
      Caption         =   "Prueba de Imagenes en Web"
      Height          =   495
      Left            =   8280
      TabIndex        =   16
      ToolTipText     =   "Prueba descargas de Imagenes Gif y JPG"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox v_4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox v_5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CheckBox v_6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox v_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox v_8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   5160
      Width           =   255
   End
   Begin VB.CheckBox v_9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox v_10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   255
   End
   Begin VB.TextBox v_11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Se expresa en Milisegundos"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox v_12 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Se expresa en Milisegundos"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox v_13 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Se expresa en Milisegundos"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox v_14 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "Se expresa en Milisegundos"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox v_15 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Se expresa en porcentaje"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox Adaptadores 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Width           =   3465
   End
   Begin VB.TextBox Dns 
      Height          =   615
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5040
      Width           =   3615
   End
   Begin VB.ComboBox ID_adapt 
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label door 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4200
      TabIndex        =   48
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label red2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4200
      TabIndex        =   47
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label vn_16 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximas conexiones por servidor 1.0"
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label vn_1 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximas conexiones por servidor"
      Height          =   255
      Left            =   5280
      TabIndex        =   45
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label vn_2 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "M.T.U."
      Height          =   375
      Left            =   1560
      TabIndex        =   44
      ToolTipText     =   "Maxima Unidad de Transmision"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Anch_Band 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho de Banda"
      Height          =   255
      Left            =   7800
      TabIndex        =   43
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Label vn_3 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "T.T.L."
      Height          =   255
      Left            =   1560
      TabIndex        =   42
      ToolTipText     =   "Tiempo de Vida (Time To Live)"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label vn_4 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "T.C.P. Tamaño del la recepcion"
      Height          =   255
      Left            =   1560
      TabIndex        =   41
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label vn_5 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Maxima Duplicacion de ACKs"
      Height          =   255
      Left            =   1560
      TabIndex        =   40
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label vn_6 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Descubrir M.T.U."
      Height          =   255
      Left            =   720
      TabIndex        =   39
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label vn_7 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "ACKs Selectivo"
      Height          =   255
      Left            =   720
      TabIndex        =   38
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label vn_8 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Detectar Hoyo Negro"
      Height          =   255
      Left            =   720
      TabIndex        =   37
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label vn_9 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Escalonado de Ventana"
      Height          =   255
      Left            =   720
      TabIndex        =   36
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label vn_10 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Timestamps"
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label vn_11 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridad D.N.S."
      Height          =   255
      Left            =   5280
      TabIndex        =   34
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label vn_12 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridad del anfitrion"
      Height          =   255
      Left            =   5280
      TabIndex        =   33
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label vn_13 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridad local"
      Height          =   255
      Left            =   1560
      TabIndex        =   32
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label vn_14 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridad de la red en Bios"
      Height          =   255
      Left            =   1560
      TabIndex        =   31
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label vn_15 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho de banda reservado"
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label x4 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "D.N.S. :"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label x3 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Puerta de Entrada :"
      Height          =   255
      Left            =   4080
      TabIndex        =   28
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label x2 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion de red asignada:"
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   3360
      Width           =   2895
   End
End
Attribute VB_Name = "Internet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const A As String = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\"
Const B As String = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
Const C As String = "SYSTEM\CurrentControlSet\Services\Tcpip\ServiceProvider"
Const dir As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkCards\"
Const Mydir As String = "SOFTWARE\DesarrolloDigital\SuiteZafiro\2.1\Red"
Private Sub Adaptadores_Click()
    Dim R As String
    R = ID_adapt.List(Adaptadores.ListIndex)
    v_2 = Reg_Valor(HKEY_LOCAL_MACHINE, A + R, "MTU")
    red2.Caption = Reg_Valor(HKEY_LOCAL_MACHINE, A + R, "DhcpIPAddress")
    door.Caption = Reg_Valor(HKEY_LOCAL_MACHINE, A + R, "DefaultGateway")
    Dns.Text = Reg_Valor(HKEY_LOCAL_MACHINE, A + R, "NameServer")
End Sub

Private Sub Aplicar_Click()
    Reg_Crear HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Internet Settings", _
    "MaxConnectionsPerServer", v_1.Text
    Reg_Crear HKEY_LOCAL_MACHINE, A, "MTU", v_2.Text
    Reg_Crear HKEY_LOCAL_MACHINE, B, "DefaultTTL", v_3.Text
    Reg_Crear HKEY_LOCAL_MACHINE, B, "TcpWindowSize", v_4.Text
    Reg_Crear HKEY_LOCAL_MACHINE, B, "TcpMaxDupAcks", v_5.Text
    Reg_Crear HKEY_LOCAL_MACHINE, B, "EnablePMTUDiscovery", v_6.Value
    Reg_Crear HKEY_LOCAL_MACHINE, B, "SackOpts", v_7.Value
    Reg_Crear HKEY_LOCAL_MACHINE, B, "EnablePMTUBHDetect", v_8.Value
    
    If v_9.Value = 1 Then Reg_Crear HKEY_LOCAL_MACHINE, B, "Tcp1323Opts", 1
    If v_10.Value = 1 Then Reg_Crear HKEY_LOCAL_MACHINE, B, "Tcp1323Opts", 2

    If v_9.Value = 1 And v_10.Value = 1 Then Reg_Crear HKEY_LOCAL_MACHINE, B, "Tcp1323Opts", 3

    Reg_Crear HKEY_LOCAL_MACHINE, C, "DnsPriority", v_11.Text
    Reg_Crear HKEY_LOCAL_MACHINE, C, "HostsPriority", v_12.Text
    Reg_Crear HKEY_LOCAL_MACHINE, C, "LocalPriority", v_13.Text
    Reg_Crear HKEY_LOCAL_MACHINE, C, "NetbtPriority", v_14.Text
    'para modificar el siguiente valor de manera manual usa:Ejecutar->'gpedit.msc'
    'Configuración del Equipo -> Plantillas Administrativas -> Red -> Configurador de paquetes QoS
    Reg_Crear HKEY_LOCAL_MACHINE, _
    "SOFTWARE\Policies\Microsoft\Windows\Psched", "NonBestEffortLimit", v_15.Text
    Reg_Crear HKEY_CURRENT_USER, _
    "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", v_16.Text
    Reg_Crear HKEY_LOCAL_MACHINE, A, "DefaultGateway", door.Caption
    Reg_Crear HKEY_LOCAL_MACHINE, A, "NameServer", Dns.Text

End Sub
Private Sub DNS1_Click()
    Dns.Text = "200.33.146.153,200.33.146.197,200.33.146.217,200.33.146.202,200.33.146.201"
End Sub
Private Sub Original_values_Click()
    'El MTU depende de la tarjeta de red v_2
    v_1.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_1")
    'v_2.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_2")
    v_3.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_3")
    v_4.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_4")
    v_5.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_5")
    'Check Box
    v_6.Value = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_6")
    v_7.Value = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_7")
    v_8.Value = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_8")
    v_9.Value = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_9")
    v_10.Value = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_10")
    'Text box
    v_11.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_11")
    v_12.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_12")
    v_13.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_13")
    v_14.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_14")
    v_15.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_15")
    v_16.Text = Reg_Valor(HKEY_LOCAL_MACHINE, Mydir, "v_16")
    
End Sub

Private Sub C2_Click()
v_1 = 20 'Conexiones por servidor
v_2 = 1492 'MTU
v_3 = 251 'TTL
v_4 = 267912 'Bufer de recepcion 1464 x 183
v_5 = 2 'Duplicacion ACK´s
v_6 = 1 'descubir MTU
v_7 = 1 'ACK Selectivo
v_8 = 0 'Detectar hoyo negro
v_9 = 1 'Escalonado de ventana
v_10 = 0 'Time stamps
v_11 = 400 'Prioridad DNS
v_12 = 500 'Prioridad Anfitrion
v_13 = 499 'Prioridad local
v_14 = 2001 'Prioridad en red
v_15 = 1 'Ancho de banda reservado
v_16 = v_1 + 20 'Conexiones por servidor 1.0

End Sub

Private Sub Form_Load()
    Dim Adp As String
    Dim bit, Adp_A As Boolean
    Dim cnt, val As Integer
    Dim exst As Boolean
    Adaptadores.Clear
    ID_adapt.Clear
    
    bit = True
    Adp_A = False
    cnt = 0
    
    Do While bit = True
        'Primero Busco los nombres de las rutas
        Adp = Key_nombre(HKEY_LOCAL_MACHINE, dir, cnt)
        If Adp = Null Then Adp = " "
        'Compruebo que exista
        Adp_A = Exist(HKEY_LOCAL_MACHINE, dir & Adp, "Description")
        'Incluyo en la lista los adaptadores existentes
        If Adp_A = True Then
            'Busco los nombres para llenar la lista de nombres
            Adaptadores.AddItem Reg_Valor(HKEY_LOCAL_MACHINE, dir & Adp, "Description")
            'Busco la direccion para llegar al registro
            ID_adapt.AddItem Reg_Valor(HKEY_LOCAL_MACHINE, dir & Adp, "ServiceName")
        Else
            bit = False
        End If
    cnt = cnt + 1
    Loop


    v_1.Text = Reg_Valor(HKEY_CURRENT_USER, _
            "Software\Microsoft\Windows\CurrentVersion\Internet Settings", _
            "MaxConnectionsPerServer")
    v_3.Text = Reg_Valor(HKEY_LOCAL_MACHINE, B, "DefaultTTL")
    v_4.Text = Reg_Valor(HKEY_LOCAL_MACHINE, B, "TcpWindowSize")
    v_5.Text = Reg_Valor(HKEY_LOCAL_MACHINE, B, "TcpMaxDupAcks")
    v_6.Value = Reg_Valor(HKEY_LOCAL_MACHINE, B, "EnablePMTUDiscovery")
    v_7.Value = Reg_Valor(HKEY_LOCAL_MACHINE, B, "SackOpts")
    v_8.Value = Reg_Valor(HKEY_LOCAL_MACHINE, B, "EnablePMTUBHDetect")
    val = Reg_Valor(HKEY_LOCAL_MACHINE, B, "Tcp1323Opts")
    Select Case val
    Case 1
        v_9.Value = 1
    Case 2
        v_10.Value = 1
    Case 3
        v_9.Value = 1
        v_10.Value = 1
    End Select

    v_11.Text = Reg_Valor(HKEY_LOCAL_MACHINE, C, "DnsPriority")
    v_12.Text = Reg_Valor(HKEY_LOCAL_MACHINE, C, "HostsPriority")
    v_13.Text = Reg_Valor(HKEY_LOCAL_MACHINE, C, "LocalPriority")
    v_14.Text = Reg_Valor(HKEY_LOCAL_MACHINE, C, "NetbtPriority")
    v_15.Text = Reg_Valor(HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Policies\Microsoft\Windows\Psched", "NonBestEffortLimit")
    v_16.Text = Reg_Valor(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server")
    
        
    If False = Exist(HKEY_LOCAL_MACHINE, Mydir, "v_1") Then
    'Text box
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_1", v_1.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_2", v_2.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_3", v_3.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_4", v_4.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_5", v_5.Text
    'Check Box
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_6", v_6.Value
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_7", v_7.Value
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_8", v_8.Value
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_9", v_9.Value
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_10", v_10.Value
    'Text box
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_11", v_11.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_12", v_12.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_13", v_13.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_14", v_14.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_15", v_15.Text
    Reg_Crear HKEY_LOCAL_MACHINE, Mydir, "v_16", v_16.Text
    MsgBox "Se han guaradado todos los valores originales en caso que desee eliminar los cambios", _
    vbOKOnly, "Informacion"
    End If
       
    
End Sub


