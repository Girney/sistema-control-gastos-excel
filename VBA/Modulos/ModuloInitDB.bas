' =========================================================================
' ModuloInitDB - Inicialización de Base de Datos en Excel
' =========================================================================
' Descripción:
'   Crea la estructura de carpetas y hojas necesarias para el sistema.
'   Inicializa los tipos de gasto por defecto.
'   Se ejecuta automáticamente al abrir el archivo.
' =========================================================================

Option Explicit

Const RUTA_BASE As String = "C:\informe\"
Const NOMBRE_ARCHIVO As String = "informe_factura.xlsx"
Const RUTA_COMPLETA As String = RUTA_BASE & NOMBRE_ARCHIVO

' =========================================================================
' Inicialización del Sistema
' =========================================================================
Sub InicializarSistema()
    On Error GoTo ErrorHandler
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta si no existe
    If Not objFSO.FolderExists(RUTA_BASE) Then
        objFSO.CreateFolder (RUTA_BASE)
    End If
    
    ' Crear subcarpeta Reportes
    If Not objFSO.FolderExists(RUTA_BASE & "Reportes") Then
        objFSO.CreateFolder (RUTA_BASE & "Reportes")
    End If
    
    ' Crear hojas si no existen
    CrearHojas
    
    ' Cargar tipos de gasto iniciales
    CargarTiposGastoIniciales
    
    ' Mostrar formulario principal
    FormularioPrincipal.Show
    
    MsgBox "Sistema inicializado correctamente", vbInformation, "Inicialización Exitosa"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error en inicialización: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' Crear Hojas de Trabajo
' =========================================================================
Private Sub CrearHojas()
    Dim hojaExistente As Worksheet
    Dim hoja As Worksheet
    
    ' Array con nombres de hojas requeridas
    Dim hojas As Variant
    hojas = Array("Registros", "TiposGasto", "Reportes", "Dashboard")
    
    ' Crear hojas que no existan
    For Each hoja In hojas
        On Error Resume Next
        Set hojaExistente = ThisWorkbook.Sheets(CStr(hoja))
        
        If hojaExistente Is Nothing Then
            ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            ActiveSheet.Name = CStr(hoja)
        End If
        
        On Error GoTo 0
    Next
    
    ' Configurar hoja Registros
    ConfigurarHojaRegistros
    
    ' Configurar hoja TiposGasto
    ConfigurarHojaTiposGasto
    
    ' Configurar hojas de reportes y dashboard
    ConfigurarHojaReportes
    ConfigurarHojaDashboard
End Sub

' =========================================================================
' Configurar Hoja Registros
' =========================================================================
Private Sub ConfigurarHojaRegistros()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Registros")
    
    ' Encabezados
    Dim encabezados As Variant
    encabezados = Array("Código", "Tipo Documento", "Fecha Emisión", _
                        "Proveedor", "Tipo Identificación", "Número Factura", _
                        "Número Nota Venta", "Nombre Proveedor", "Tipo Gasto", _
                        "Subtotal", "IVA", "Total", "Forma Pago", "Observaciones")
    
    ' Escribir encabezados si no existen
    If ws.Range("A1").Value = "" Then
        Dim i As Integer
        For i = LBound(encabezados) To UBound(encabezados)
            ws.Cells(1, i + 1).Value = encabezados(i)
        Next i
        
        ' Formatear encabezados
        With ws.Range("A1:N1")
            .Font.Bold = True
            .Interior.Color = RGB(51, 102, 153)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Ajustar ancho de columnas
        ws.Columns("A:N").AutoFit
    End If
End Sub

' =========================================================================
' Configurar Hoja Tipos Gasto
' =========================================================================
Private Sub ConfigurarHojaTiposGasto()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TiposGasto")
    
    ' Encabezados
    If ws.Range("A1").Value = "" Then
        ws.Range("A1").Value = "Tipo Gasto"
        
        With ws.Range("A1")
            .Font.Bold = True
            .Interior.Color = RGB(51, 102, 153)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
End Sub

' =========================================================================
' Configurar Hoja Reportes
' =========================================================================
Private Sub ConfigurarHojaReportes()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Reportes")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Range("A1").Value = "Reportes - Se mostrarán aquí"
    End If
End Sub

' =========================================================================
' Configurar Hoja Dashboard
' =========================================================================
Private Sub ConfigurarHojaDashboard()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Range("A1").Value = "Dashboard de Gastos"
        With ws.Range("A1")
            .Font.Bold = True
            .Font.Size = 16
        End With
    End If
End Sub

' =========================================================================
' Cargar Tipos de Gasto Iniciales
' =========================================================================
Sub CargarTiposGastoIniciales()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TiposGasto")
    
    ' Array de tipos de gasto
    Dim tiposGasto As Variant
    tiposGasto = Array( _
        "Alimentación", "Transporte", "Hospedaje", "Materiales de oficina", _
        "Servicios básicos", "Combustible", "Mantenimiento", "Equipos tecnológicos", _
        "Capacitación", "Papelería", "Internet", "Telefonía", _
        "Software", "Hardware", "Publicidad", "Impresiones", _
        "Repuestos", "Herramientas", "Seguridad", "Otros")
    
    ' Limpiar datos previos
    ws.Range("A2:A100").ClearContents
    
    ' Cargar nuevos datos
    Dim i As Integer
    For i = LBound(tiposGasto) To UBound(tiposGasto)
        ws.Cells(i + 2, 1).Value = tiposGasto(i)
    Next i
End Sub

' =========================================================================
' Obtener Próximo Código Factura
' =========================================================================
Function ObtenerProximoCodigoFactura() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Registros")
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim contador As Long
    Dim i As Long
    contador = 0
    
    ' Contar facturas existentes
    For i = 2 To ultimaFila
        If Left(ws.Cells(i, 1).Value, 3) = "Fac" Then
            contador = contador + 1
        End If
    Next i
    
    ObtenerProximoCodigoFactura = "Fac" & Format(contador + 1, "0000")
End Function

' =========================================================================
' Obtener Próximo Código Nota de Venta
' =========================================================================
Function ObtenerProximoCodigoNV() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Registros")
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim contador As Long
    Dim i As Long
    contador = 0
    
    ' Contar notas de venta existentes
    For i = 2 To ultimaFila
        If Left(ws.Cells(i, 1).Value, 2) = "NV" Then
            contador = contador + 1
        End If
    Next i
    
    ObtenerProximoCodigoNV = "NV" & Format(contador + 1, "0000")
End Function

' =========================================================================
' Punto de entrada al abrir el libro
' =========================================================================
Sub Auto_Open()
    InicializarSistema
End Sub
