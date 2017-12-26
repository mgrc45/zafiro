Attribute VB_Name = "Conexion_Access_2000"
'Esta conexion fue creada en puro codigo en vez de usar el  "Microsoft ADO Data Control" puesto que
'este me mantiene una conexion siempre activa mientras este cargado, ademas de mantener 2
'conexiones una al ADO Data Control y otra a la Data Base

'librerías hablilitadas
'Proyectos ->Referencias
'Microsoft Activex Data Objects 2.1 Library , y
'Microsoft ADO EXt. 2.1 for DDl and Security

Private Con As ADODB.Connection
Public Rec As ADODB.Recordset
Public Function Conexion(ByVal Encendido As Boolean, ByVal Nombre_Tabla As String)
'Solo permite lectura pero varios usuarios pueden conectarse a la base eso ayada pues se puede usar el marcador move
    'Establecemos
    Set Con = New ADODB.Connection
    Set Rec = New ADODB.Recordset
    'Nombre_Tabla = "Visual_Desempeño"
If Encendido Then
    On Error GoTo ErrordeConex
    'Con.Open "Provider = Microsoft.Jet.OLEDB.4.0; User ID=Miguel Angel;Password=pepe;Data Source=DataBase.zfr; Mode=Read; Persist Security Info=False"
    Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DataBase.zfr;Persist Security Info=False;JET OLEDB:DATABASE PASSWORD=z0a9u0r9o0y9a0l"
    Rec.Open "SELECT * FROM " & Nombre_Tabla, Con, adOpenStatic, adLockReadOnly

   'If Con.State = 1 And Rec.State = 1 Then MsgBox ("Conexión abierta")

Else
    'Si la conexion esta abierta cierrala
    If Rec.State = 1 Then Rec.Close
    If Rec.State = 1 Then Con.Close
End If

ErrordeConex:
    If Err.Number Then
    MsgBox "No se estableció una conexión"
    Exit Function
    End If
    
End Function

