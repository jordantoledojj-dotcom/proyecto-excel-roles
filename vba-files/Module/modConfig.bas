Attribute VB_Name = "modConfig"
Option Explicit


' =============================================================================================================
' Constantes generales del proyecto
' =============================================================================================================
Public Const EVENTO_SIN_ACCION As Integer = 0
Public Const EVENTO_NUEVO_REGISTRO As Integer = 1
Public Const EVENTO_ACTUALIZAR_REGISTRO As Integer = 2
Public Const EVENTO_ELIMINAR_REGISTRO As Integer = 3


' =============================================================================================================
' Constantes relacionadas con la tabla de empleados
' =============================================================================================================
Public Const TABLA_EMPLEADOS As String = "tblEmpleado"

Public Const COL_EMP_RFC_CORTO As String = "RFC Corto"
Public Const COL_EMP_NOMBRE As String = "Nombre"
Public Const COL_EMP_PATERNO As String = "A. Paterno"
Public Const COL_EMP_MATERNO As String = "A. Materno"
Public Const COL_EMP_RFC As String = "RFC"
Public Const COL_EMP_F_NACIMIENTO As String = "F. Nacimiento"
Public Const COL_EMP_NOMBRAMIENTO As String = "Nombramiento actual"
Public Const COL_EMP_PUESTO As String = "Puesto funcional"
Public Const COL_EMP_SUBADMIN As String = "Subadministración"
Public Const COL_EMP_DEPARTAMENTO As String = "Departamento"
Public Const COL_EMP_JEFE_DIRECTO As String = "Jefe directo"
Public Const COL_EMP_ESTATUS As String = "Estatus"
Public Const COL_EMP_F_INGRESO As String = "F. Ingreso"
Public Const COL_EMP_F_INICIO_CARGO As String = "F. Incio del cargo"
Public Const COL_EMP_EXTENSION As String = "Extensión"

Public Const TABLA_EMPLEADOS_COLUMN_COUNT As Integer = 15


' =============================================================================================================
' Constantes relacionadas con la tabla de subadministraciones
' =============================================================================================================
Public Const TABLA_SUBS As String = "tSubs"

Public Const COL_SUBS_NOMBRE = "Nombre"
Public Const COL_SUBS_JEFE = "Subadministrador"

Public Const TABLA_SUBS_COLUMN_COUNT As Integer = 2

' =============================================================================================================
' Constantes relacionadas con la tabla de departamentos
' =============================================================================================================
Public Const TABLA_DEPTOS As String = "tDepartamentos"

Public Const COL_DEPTOS_NOMBRE = "Departamento"
Public Const COL_DEPTOS_SUBADMIN = "Subadministración"

Public Const TABLA_DEPARTAMENTEOS_COLUMN_COUNT As Integer = 2


' =============================================================================================================
' Constantes relacionadas con la tabla de plataformas
' =============================================================================================================
Public Const TABLA_PLATAFORMAS As String = "tPlataformas"

Public Const COL_PLATAFORMAS_ID_CLAVE_PLATAFORMA = "Clave de plataforma"
Public Const COL_PLATAFORMAS_NOMBRE = "Nombre plataforma"
Public Const COL_PLATAFORMAS_DESCRIPCION = "Descripción"

Public Const TABLA_PLATAFORMAS_COLUMN_COUNT As Integer = 3


' =============================================================================================================
' Constantes relacionadas con la tabla de roles
' =============================================================================================================
Public Const TABLA_ROLES As String = "tRoles"

Public Const COL_ROLES_ID_CLAVE_ROL_TECNICO = "Clave rol técnico"
Public Const COL_ROLES_ROL_FUNCIONAL = "Rol funcional"
Public Const COL_ROLES_ROL_TECNICO = "Rol técnico"
Public Const COL_ROLES_DEFINICION_ALCANCE = "Definición y alcance de rol"
Public Const COL_ROLES_FK_CLAVE_PLATAFORMA = "Clave de plataforma"

Public Const TABLA_ROLES_COLUMN_COUNT As Integer = 5


' =============================================================================================================
' Constantes relacionadas con la tabla de Empleados-Roles
' =============================================================================================================
Public Const TABLA_EMPLEADOS_ROLES As String = "tEmpleadosRoles"

Public Const COL_EMP_ROLES_PK_RFC_CORTO = "RFC Corto"
Public Const COL_EMP_ROLES_PK_CLAVE_ROL = "Clave rol técnico"
Public Const COL_EMP_ROLES_FECHA_ALTA = "F. Alta"
Public Const COL_EMP_ROLES_FECHA_BAJA = "F. Baja"
Public Const COL_EMP_ROLES_SOLICITANTE = "Solicitó"
Public Const COL_EMP_ROLES_NOTAS = "Notas"

Public Const TABLA_EMPLEADO_ROLES_COLUMN_COUNT As Integer = 6


' =============================================================================================================
' Paleta de colores del SAT
' =============================================================================================================
Public Property Get COLOR_VINO() As Long
    COLOR_VINO = RGB(98, 19, 51)
End Property
Public Property Get COLOR_BLANCO() As Long
    COLOR_BLANCO = RGB(255, 255, 255)
End Property
Public Property Get COLOR_VERDE_OSCURO() As Long
    COLOR_VERDE_OSCURO = RGB(0, 47, 42)
End Property
Public Property Get COLOR_ARENA() As Long
    COLOR_ARENA = RGB(231, 210, 149)
End Property

' =============================================================================================================
' Paleta de colores usada en la plataforma
' =============================================================================================================
Public Property Get COLOR_HABILITADO() As Long
    COLOR_HABILITADO = RGB(255, 255, 255)
End Property
Public Property Get COLOR_DESHABILITADO() As Long
    COLOR_DESHABILITADO = RGB(120, 120, 120)
End Property
Public Property Get COLOR_FUENTE_HABILITADA() As Long
    COLOR_FUENTE_HABILITADA = RGB(0, 0, 0)
End Property
Public Property Get COLOR_FUENTE_DESHABILITADA() As Long
    COLOR_FUENTE_DESHABILITADA = RGB(254, 254, 254)
End Property
Public Property Get COLOR_FONDO_FORM() As Long
    COLOR_FONDO_FORM = RGB(64, 0, 0)
End Property
Public Property Get COLOR_SAM() As Long
    COLOR_SAM = RGB(0, 47, 42)
End Property
Public Property Get COLOR_VERDE_SAM() As Long
    COLOR_VERDE_SAM = RGB(30, 91, 79)
End Property

