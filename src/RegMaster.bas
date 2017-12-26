Attribute VB_Name = "Registry"
'Creado por Miguel Angel Gonzalez Alonso
Option Explicit
' Constantes requeridas para encontrar una cadena en el registry
'___________________________________________________________________________________________________
  Public Const HKEY_CLASSES_ROOT As Long = &H80000000
  Public Const HKEY_CURRENT_USER As Long = &H80000001
  Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Public Const HKEY_USERS As Long = &H80000003
  
  Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004 'Solo para NT
  Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
  Public Const HKEY_DYN_DATA As Long = &H80000006
  Public Const HKEY_FIRST = HKEY_CLASSES_ROOT
  Public Const HKEY_LAST = HKEY_DYN_DATA
  'HKEY_CLASSES_ROOT es un duplicado de HKEY_LOCAL_MACHINE\Software\Classes
  'HKEY_CURRENT_USER es un duplicado de HKEY_USERS\[Usuario]

' Constantes requeridas para especificar el tipo de valores en las llaves.
' __________________________________________________________________________________________________
  Private Const REG_NONE As Long = 0                  ' No value type
  Public Const REG_SZ As Long = 1                    ' Unicode nul terminated string
  Public Const REG_EXPAND_SZ As Long = 2             ' Unicode nul terminated string w/enviornment var
  Public Const REG_BINARY As Long = 3                ' Free form binary
  Public Const REG_DWORD As Long = 4                 ' 32-bit number
  Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4   ' 32-bit number (same as REG_DWORD)
  Private Const REG_DWORD_BIG_ENDIAN As Long = 5      ' 32-bit number
  Private Const REG_LINK As Long = 6                  ' Symbolic Link (unicode)
  Private Const REG_MULTI_SZ As Long = 7              ' Multiple Unicode strings
  Private Const REG_RESOURCE_LIST As Long = 8         ' Resource list in the resource map
  Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9 ' Resource list in the hardware description
  Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
  Private Const Nulo As Long = 0&
  
' Constantes requeridas para checar un error
' __________________________________________________________________________________________________
  Private Const ERROR_SUCCESS As Long = 0
  ' Quitar esto Variables no officiales
  Global Const ERROR_NONE = 0
  Global Const ERROR_BADDB = 1
  Global Const ERROR_BADKEY = 2
  Global Const ERROR_CANTOPEN = 3
  Global Const ERROR_CANTREAD = 4
  Global Const ERROR_CANTWRITE = 5
  Global Const ERROR_OUTOFMEMORY = 6
  Global Const ERROR_INVALID_PARAMETER = 7
  Global Const ERROR_ACCESS_DENIED = 8
  Global Const ERROR_INVALID_PARAMETERS = 87
  Global Const ERROR_NO_MORE_ITEMS = 259

' Privilegios de acceso
' __________________________________________________________________________________________________
  Private Const KEY_ALL_ACCESS As Long = &H3F
 'Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
' Derechos estandar
  Private Const SYNCHRONIZE As Long = &H100000
  Private Const READ_CONTROL As Long = &H20000
  Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
  Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
  Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
  Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
  Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
  'KEY_READ Combinacion de:
  'KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, y KEY_NOTIFY acceso.
  Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
  'KEY_WRITE Combinacion de: KEY_SET_VALUE y KEY_CREATE_SUB_KEY acceso.
  Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
  'Permiso para leer
  Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

' Abrir/Crear constantes
'___________________________________________________________________________________________________
  Private Const REG_OPTION_NON_VOLATILE As Long = 0
  Private Const REG_OPTION_VOLATILE As Long = &H1
  Private Const REG_CREATED_NEW_KEY As Long = &H1
  Private Const REG_OPENED_EXISTING_KEY As Long = &H2

' Variables del modulo
'___________________________________________________________________________________________________
  Private Registro As Long
  Private phkresult As Long

' Declaraciones requeridas para accesar al registry
'___________________________________________________________________________________________________
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Para abrir una key de sistema
    'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, ByRef phkresult As Long) As Long    ' 16 bits
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkresult As Long) As Long    ' 32 bits
    
'Para cerrar la key en 16 y 32 bits
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long

'Para enumerar y saber el nombre de las keys
    Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As FILETIME) As Long '32 Bits

'Para conocer el tipo y contenido de un registro
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByRef lpdwType As Long, ByRef lpbData As Any, ByRef cbData As Long) As Long  '32 Bits

'Para crear un Registro
    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, ByRef lpbData As Any, ByVal cbData As Long) As Long   '32 Bits

'Para eliminar un Registro
    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal szValueName As String) As Long    '16 Bits

'Para crear una key
    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOption As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkresult As Long, ByRef lpdwDisposition As Long) As Long ' 32 bits

'Para eliminar una key    16 y 32 Bits
    'Windows 95: RegDeleteKey elimina una llave y todas sus decendientes
    'Windows NT: RegDeleteKey elimina una llave especifica y no debe tener sub_llaves
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpszSubKey As String) As Long
    
'Pare enumerar y saber el nombre de los registros  16 y 32 Bits
    Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
    
'Para obtener informaciobn de las keys 32 Bits
    Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hkey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Function cad(ByVal cadena As String) As Long
' Esta funcion sirve para convertir cadenas en constantes aceptadas por el programa
On Error Resume Next
Select Case cadena
    Case "HKEY_CLASSES_ROOT"
                cad = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
                cad = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
                cad = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
                cad = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
                cad = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
                cad = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
                cad = HKEY_DYN_DATA
    Case Else
                cad = 0
End Select
End Function

Public Function tipo(ByVal tipo_num As Integer) As String
' Esta funcion sirve para convertir numeros en cadenas
On Error Resume Next
Select Case tipo_num
    Case 0 ' No value type
                tipo = "REG_NONE"
    Case 1 ' Unicode nul terminated string
                tipo = "REG_SZ"
    Case 2 'Unicode nul terminated string w/enviornment var
                 tipo = "REG_EXPAND_SZ"
    Case 3 ' Free form binary
                tipo = "REG_BINARY"
    Case 4 ' 32-bit number
                tipo = "REG_DWORD"
    Case 5 ' 32-bit number
                tipo = "REG_DWORD_BIG_ENDIAN"
    Case 6 ' Symbolic Link (unicode)
                tipo = "REG_LINK"
    Case 7 ' Multiple Unicode strings
                tipo = "REG_MULTI_SZ"
    Case 8 ' Resource list in the resource map
                tipo = "REG_RESOURCE_LIST"
    Case 9 ' Resource list in the hardware description
                tipo = "REG_FULL_RESOURCE_DESCRIPTOR"
    Case 10
                tipo = "REG_RESOURCE_REQUIREMENTS_LIST"
    Case Else
                tipo = 0
End Select
End Function

Public Function Exist(ByVal Root As Long, _
                      ByVal Key As String, _
                      Optional ByVal name As String) As Boolean
'_________________________________________________________________
' Descripcion: Funcion para consultar si existe un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
' Syntaxis:
'       boolean = Exist(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products","NoRun")
' Regresa:
'       un valor booleano
'_________________________________________________________________
    phkresult = 0
    On Error GoTo salir
    
    Registro = RegOpenKeyEx(Root, Key, Nulo, KEY_READ, phkresult) 'Abrir el key
    DoEvents ' Otorgo tiempo al sistema operativo
    If phkresult = ERROR_SUCCESS Then GoTo salir 'En caso de que no exista el key

    If Len(name) > 0 Then
        Registro = RegQueryValueEx(phkresult, name, Nulo, 0&, ByVal 0&, 0&) 'Preguntando el valor 32 Bits
        If Registro <> 0 Then GoTo salir
    End If
    

'Debemos cerrar la cadena para evitar que se corrompa
Registro = RegCloseKey(phkresult)
Exist = True 'Si llego aqui es k todo salio bien
Exit Function
salir:
'Debemos cerrar la cadena para evitar que se corrompa
Registro = RegCloseKey(phkresult)
Exist = False
End Function
Private Function Abrir(ByVal Root As Long, _
                       ByVal Key As String, _
                       Optional ByVal name As String = "", _
                       Optional ByVal AccessRights As Long = KEY_READ) As Long
'_________________________________________________________________
' Descripcion: Funcion para abrir un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
'       AccessRights - KEY_SET_VALUE, KEY_READ, etc
' Syntaxis:
'       long = Abrir(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products","NoRun",KEY_SET_VALUE)
' Regresa:
'       phkResult (Un valor que identifica al registro)
'_________________________________________________________________
    phkresult = 0
    On Error GoTo salir
    
    Registro = RegOpenKeyEx(Root, Key, Nulo, AccessRights, phkresult)
    DoEvents ' Otorgo tiempo al sistema operativo
    If phkresult = ERROR_SUCCESS Then GoTo salir 'En caso de que no exista el key
    
    If Len(name) > 0 Then
        Registro = RegQueryValueEx(phkresult, name, Nulo, 0&, ByVal 0&, 0&) 'Preguntando el valor 32 Bits
        If Registro <> 0 Then GoTo salir
    End If
    
    
'Dejo abierta la cadena para permitir modificarla
Abrir = phkresult
Exit Function
salir:
'Debemos cerrar la cadena para evitar que se corrompa
Registro = RegCloseKey(phkresult)
Abrir = ERROR_SUCCESS 'Ocurrio un error en la apertura del registro o este no existe
End Function
Public Function Reg_Valor(ByVal Root As Long, _
                              ByVal Key As String, _
                              ByVal name As String) As String
'______________________________________________________________
' Descripcion: Funcion para consultar el valor de un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
' Syntaxis:
' Cadena = Reg_Valor(HKEY_CURRENT_USER,_
'                     "Miguel\Miguel 2", "MIKE")
' Regresa:
'       Una cadena de caracteres
' _____________________________________________________________
On Error Resume Next
    Reg_Valor = ""
    
    'Abriendo el registro
    phkresult = Abrir(Root, Key, name)
    If phkresult = ERROR_SUCCESS Then Exit Function
    
    Dim tipo_dato As Long
    ' Preguntando el tipo de dato
    Registro = RegQueryValueEx(phkresult, name, Nulo, tipo_dato, ByVal 0&, 0&)
            
            If tipo_dato = REG_DWORD Then 'Dato numerico (En Decimal)
                    Dim lngBuffer As Long
                    Registro = RegQueryValueEx(phkresult, name, Nulo, tipo_dato, lngBuffer, 4&)
                    Reg_Valor = Trim(Str(lngBuffer))  'Devuelve los datos como cadena y sin espacios
                    
            Else 'REG_SZ,REG_BINARY,REG_MULTI_SZ,REG_EXPAND_SZ  (Cadena de letras)
                    Dim BufferSize As Long
                    Dim strBuffer As String
                     'Preguntando el tamaño del buffer
                    Registro = RegQueryValueEx(phkresult, name, Nulo, 0&, ByVal strBuffer, BufferSize)
                    strBuffer = Space(BufferSize) ' Carga el buffer de recepcion
                    Registro = RegQueryValueEx(phkresult, name, 0&, 0&, ByVal strBuffer, BufferSize)
                    If BufferSize > 0 Then Reg_Valor = strBuffer
              End If


Registro = RegCloseKey(phkresult)
End Function
Public Function Reg_Tipo(ByVal Root As Long, _
                         ByVal Key As String, _
                         ByVal name As String) As String
'______________________________________________________________
' Descripcion: Funcion para consultar el tipo de un registro
' Syntaxis:
' Cadena = Reg_Tipo(HKEY_CURRENT_USER,_
'                     "Miguel\Miguel 2", "MIKE")
' _____________________________________________________________
On Error Resume Next
    'Abriendo el registro
    phkresult = Abrir(Root, Key, name)
    If phkresult = ERROR_SUCCESS Then Exit Function
    
    Dim tipo_dato As Long
    ' Pregunte el tipo de dato
    Registro = RegQueryValueEx(phkresult, name, Nulo, tipo_dato, ByVal 0&, 0&)


    Select Case tipo_dato
    Case REG_NONE 'No value type
                Reg_Tipo = "REG_NONE"
    Case REG_SZ 'Unicode nul terminated string
                Reg_Tipo = "REG_SZ"
    Case REG_EXPAND_SZ 'Unicode nul terminated string w/enviornment var
                Reg_Tipo = "REG_EXPAND_SZ"
    Case REG_BINARY 'Free form binary
                Reg_Tipo = "REG_BINARY"
    Case REG_DWORD '32-bit number
                Reg_Tipo = "REG_DWORD"
    Case REG_DWORD_LITTLE_ENDIAN '32-bit number (same as REG_DWORD)
                Reg_Tipo = "REG_DWORD_LITTLE_ENDIAN"
    Case REG_DWORD_BIG_ENDIAN '32-bit number
                Reg_Tipo = "REG_DWORD_BIG_ENDIAN"
    Case REG_LINK 'Symbolic Link (unicode)
                Reg_Tipo = "REG_LINK"
    Case REG_MULTI_SZ 'Multiple Unicode strings
                Reg_Tipo = "REG_MULTI_SZ"
    Case REG_RESOURCE_LIST 'Resource list in the resource map
                Reg_Tipo = "REG_RESOURCE_LIST"
    Case REG_FULL_RESOURCE_DESCRIPTOR 'Resource list in the hardware description
                Reg_Tipo = "REG_FULL_RESOURCE_DESCRIPTOR"
    Case REG_RESOURCE_REQUIREMENTS_LIST
                Reg_Tipo = "REG_RESOURCE_REQUIREMENTS_LIST"
    Case Else 'Valor desconocido
                Reg_Tipo = ""
    End Select


Registro = RegCloseKey(phkresult)
End Function
Public Function Reg_Crear(ByVal Root As Long, ByVal Key As String, _
                               ByVal name As String, Data As Variant)
'_________________________________________________________________
' Descripcion: Sub-programa para crear un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
'       Data - Es el dato que va a contener.
' Syntaxis:
'       variant = Reg_Crear(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products","NoRun","1")
' Regresa:
'
'Nota: Si un valor previo existe lo elimina
'Valor maximo de un numero:1000000000 tal vez mas
'
'_________________________________________________________________
On Error GoTo salir
' Abriendo el registro para grabar
    phkresult = Abrir(Root, Key, , KEY_WRITE)
    If phkresult = ERROR_SUCCESS Then phkresult = Key_Crear(Root, Key) 'Si no existe el Key
    If phkresult <> REG_CREATED_NEW_KEY Then GoTo salir
    
    phkresult = Abrir(Root, Key, , KEY_WRITE)
    If phkresult = ERROR_SUCCESS Then GoTo salir
    

' Creando el dato
If IsNumeric(Data) Then 'Dato numerico
    Dim lngKeyValue As Long
    lngKeyValue = CLng(Data)
    Registro = RegSetValueEx(phkresult, name, Nulo, REG_DWORD, lngKeyValue, 4&) ' 4& = 4-byte word (long integer)
Else                    'Dato cadena
    Dim strKeyValue As String
    strKeyValue = Trim(Data) & Chr$(0)     'Caracter final nulo
    Registro = RegSetValueEx(phkresult, name, Nulo, REG_SZ, ByVal strKeyValue, Len(strKeyValue))
End If

salir:
Registro = RegCloseKey(phkresult)
End Function
Public Function Key_Crear(ByVal Root As Long, ByVal Key As String) As Long
'_________________________________________________________________
' Descripcion: Funcion para un key
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion hasta donde exista
' Syntaxis:
'       variant = Reg_Crear(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products")
' Regresa:
'
'_________________________________________________________________
On Error Resume Next
    Dim retval As Long
    'Creando la key
    Registro = RegCreateKeyEx(Root, Key, Nulo, Nulo, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0&, phkresult, retval) ' 32 bits

    Key_Crear = retval
    Registro = RegCloseKey(phkresult)
End Function
Public Function Reg_Borrar(ByVal Root As Long, _
                           ByVal Key As String, _
                           ByVal name As String) As Boolean
'_________________________________________________________________
' Descripcion: Funcion para borrar un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
' Syntaxis:
'       variant = Reg_Borrar(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products","NoRun")
' Regresa:
' Valor booleano
'Nota:Si el valor no existe no se traba
'_________________________________________________________________
On Error Resume Next
    Reg_Borrar = False
    ' Abriendo el registro para grabar
    phkresult = Abrir(Root, Key, , KEY_WRITE)
    If phkresult <> ERROR_SUCCESS Then Registro = RegDeleteValue(phkresult, name) 'Borrando el reg
    If Registro = ERROR_SUCCESS Then Reg_Borrar = True
    'Registro = RegCloseKey(phkresult) no hay nada por cerrar
End Function
Public Function Key_Borrar(ByVal Root As Long, ByVal Key As String) As Boolean
'_________________________________________________________________
' Descripcion: Funcion para borrar una key
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
' Syntaxis:
'       variant = Key_Borrar(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products")
' Regresa:
' Valor booleano
'_________________________________________________________________
On Error Resume Next
    Key_Borrar = False
    ' Abriendo el registro para grabar
    phkresult = Abrir(Root, Key, , KEY_WRITE)
    If phkresult <> ERROR_SUCCESS Then Registro = RegDeleteKey(Root, Key) 'Borrando el key
    If Registro = ERROR_SUCCESS Then Key_Borrar = True
    'Registro = RegCloseKey(phkresult) no hay nada por cerrar
End Function
Public Function Key_nombre(ByVal Root As Long, ByVal Key As String, ByVal Index As Integer) As String
'_________________________________________________________________
' Descripcion: Funcion para consultar el nombre de una key
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
' Syntaxis:
'       Cadena = Key_name(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products",1)
' Regresa:
'       Una cadena con el nombre del index usado
'_________________________________________________________________
On Error GoTo salir
    Dim lpftLastWriteTime As FILETIME
    Dim Buffer As String * 255
    Buffer = ""
    Dim filtro As String
    filtro = ""
    
    'Abriendo el registro
    phkresult = Abrir(Root, Key, "")
    If phkresult = ERROR_SUCCESS Then GoTo salir
    
    Registro = RegEnumKeyEx(phkresult, Index, Buffer, 255, Nulo, 0&, 0&, lpftLastWriteTime) '32 Bits
    filtro = Trim(Buffer) 'Retiro los espacios
    If Len(filtro) > 0 Then Key_nombre = Left(filtro, Len(filtro) - 1) 'Retiro el caracter fin de linea
    
salir:
    Registro = RegCloseKey(phkresult)
End Function

Public Function Enumera(ByVal Root As Long, ByVal Key As String, ByVal is_key As Boolean) As Integer
'_________________________________________________________________
' Descripcion: Funcion para consultar el numero claves que existen
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
' Syntaxis:
'       Cadena = EnumKeys(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products",false)
' Regresa:
'       La cantidad de sub claves (integer)
'       cadena con el nombre del index usado
'_________________________________________________________________
On Error GoTo salir
    Dim lpftLastWriteTime As FILETIME
    Dim SubKeysNum, numValues As Long
    Enumera = 0
    'Abriendo el registro
    phkresult = Abrir(Root, Key, "")
    If phkresult = ERROR_SUCCESS Then GoTo salir
    
    Registro = RegQueryInfoKey(phkresult, 0&, 0&, Nulo, SubKeysNum, 0&, 0&, numValues, 0&, 0&, 0&, lpftLastWriteTime)
    If is_key = True Then Enumera = SubKeysNum
    If is_key = False Then Enumera = numValues

salir:
    Registro = RegCloseKey(phkresult)
End Function

Public Function Reg_nombre(ByVal Root As Long, ByVal Key As String, ByVal Index As Long) As String
'_________________________________________________________________
' Descripcion: Funcion para consultar nombre de un registro
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       index           - Es el numero del registro a buscar.
' Syntaxis:
'       Cadena = Reg_nombre(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products",1)
' Regresa:
'       Una cadena con el nombre del index usado
'_________________________________________________________________
On Error GoTo salir
    Dim MaxValueNameLen As Long
    Dim lpftLastWriteTime   As FILETIME
    Dim name As String * 255
    name = ""
    Dim filtro As String
    filtro = ""
    Reg_nombre = ""
    'Abriendo el registro
    phkresult = Abrir(Root, Key, "")
    If phkresult = 0 Then GoTo salir

    Registro = RegQueryInfoKey(phkresult, 0&, 0&, Nulo, _
                                0&, 0&, 0&, 0&, MaxValueNameLen, _
                                0&, 0&, lpftLastWriteTime)
        
    Registro = RegEnumValue(phkresult, Index, name, 255, Nulo, 0&, ByVal 0&, 0&)
    filtro = Trim$(name)
    If Len(filtro) > 1 Then
        filtro = Left$(filtro, Len(filtro) - 1) ' Retiro los caracteres de fin de linea
        Reg_nombre = filtro
        End If

salir:
    Registro = RegCloseKey(phkresult)
End Function

Public Function Reg_Modificar(ByVal Root As Long, ByVal Key As String, _
                               ByVal name As String, ByVal Data As Variant, ByVal tipo_dato As Long)
'_________________________________________________________________
' Descripcion: Funcion para modificar un registro sin cambiar su tipo
' Parameteros:
'       Root - HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc
'       Key - La direccion
'       Name - El nombre del registro
'       Data - Es el dato que va a contener.
'       tipo_dato = 1,2,4 ó REG_SZ,REG_EXPAND_SZ,REG_DWORD
' Syntaxis:
'       variant = Reg_Crear(HKEY_CURRENT_USER, _
'       "Software\AAA-Registry Test\Products","NoRun","1",REG_SZ)
' Regresa:
'
'_________________________________________________________________
'Nota: No esta implementado el valor "REG_BINARY" ni "REG_MULTI_SZ"
On Error GoTo salir
    ' Abriendo el registro para grabar
    phkresult = Abrir(Root, Key, , KEY_WRITE)
    If phkresult = ERROR_SUCCESS Then GoTo salir

    Select Case tipo_dato
    Case REG_SZ, REG_EXPAND_SZ
        Dim SData As String: SData = Trim$(Data) & Chr(0) 'Caracter nulo
        Registro = RegSetValueEx(phkresult, name, Nulo, tipo_dato, ByVal SData, Len(SData))
    'Case REG_BINARY 'Forma Binaria
        'If IsNumeric(Data) Then Registro = RegSetValueEx(phkresult, Name, 0&, REG_BINARY, Data, UBound(Data))
    Case REG_DWORD 'Numero de 32bits
        If IsNumeric(Data) Then Registro = RegSetValueEx(phkresult, name, Nulo, REG_DWORD, CLng(Data), 4&)
    End Select

salir:
    Registro = RegCloseKey(phkresult)
End Function



