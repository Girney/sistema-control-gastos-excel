# 💼 Sistema de Control de Gastos en Excel

**Aplicación profesional para administrar facturas y notas de venta mediante VBA y UserForms**

[![Excel](https://img.shields.io/badge/Excel-2016+-green.svg)](https://www.microsoft.com/excel)
[![VBA](https://img.shields.io/badge/VBA-Macro%20Enabled-blue.svg)](https://docs.microsoft.com/vba/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📋 Características

✅ **Gestión de Documentos**
- Registro de Facturas (Fac0001, Fac0002, ...)
- Registro de Notas de Venta (NV0001, NV0002, ...)
- CRUD completo (Crear, Leer, Actualizar, Eliminar)

✅ **Búsqueda Inteligente**
- Búsqueda por múltiples campos
- ComboBox con autocompletado
- Filtrado dinámico en tiempo real

✅ **Reportes Avanzados**
- Reporte de Facturas con totales
- Reporte de Notas de Venta
- Reporte General consolidado
- Agrupación por proveedor, tipo de gasto, forma de pago

✅ **Exportación**
- Exportar a Excel (.xlsx)
- Exportar a PDF con formato profesional
- Nombres automáticos con timestamp

✅ **Dashboard Dinámico**
- 12+ indicadores clave
- Gráficos actualizables automáticamente
- Análisis por proveedor, tipo de gasto y forma de pago

✅ **Validaciones Completas**
- Validación de campos obligatorios
- Prevención de duplicados
- Validación de fechas y montos
- Mensajes de error amigables

---

## 🚀 Instalación Rápida

### Paso 1: Descargar
```bash
git clone https://github.com/Girney/sistema-control-gastos-excel.git
cd sistema-control-gastos-excel
```

### Paso 2: Abrir en Excel
1. Abre `informe_factura.xlsm` en Microsoft Excel
2. Habilita las macros cuando se solicite
3. El sistema verificará y creará automáticamente:
   - Carpeta `C:\informe\`
   - Hojas de trabajo necesarias
   - Base de datos interna

### Paso 3: Comenzar
1. Presiona `Ctrl + Shift + I` o ejecuta la macro `IniciarAplicacion`
2. Se abrirá el formulario principal
3. ¡Comienza a registrar gastos!

---

## 📁 Estructura del Proyecto

```
sistema-control-gastos-excel/
├── README.md                          # Este archivo
├── GUIA_USO.md                       # Manual completo de usuario
├── informe_factura.xlsm              # Archivo Excel principal
│
├── VBA/
│   ├── ModMain.bas                   # Inicialización y menú principal
│   ├── ModDatabase.bas               # CRUD y gestión de datos
│   ├── ModValidation.bas             # Validaciones
│   ├── ModSearch.bas                 # Búsqueda y filtrado
│   ├── ModReports.bas                # Generación de reportes
│   ├── ModDashboard.bas              # Dashboard y gráficos
│   ├── ModExport.bas                 # Exportación Excel/PDF
│   ├── ModAutoComplete.bas           # ComboBox inteligente
│   └── ModUtils.bas                  # Funciones auxiliares
│
└── Formularios/
    ├── FrmMain.frm                   # Formulario principal
    ├── FrmReports.frm                # Formulario de reportes
    ├── FrmSearch.frm                 # Búsqueda avanzada
    └── FrmAddGasto.frm               # Agregar tipo de gasto
```

---

## 🎯 Guía de Uso Rápida

### Registrar una Factura
1. Abre el formulario principal (`Ctrl + Shift + I`)
2. Selecciona "Factura"
3. Completa los campos:
   - Fecha de emisión
   - RUC del proveedor
   - Nombre del proveedor
   - Número de factura
   - Tipo de gasto (con búsqueda inteligente)
   - Subtotal (IVA se calcula automáticamente)
   - Forma de pago
4. Clic en "Guardar"

### Registrar una Nota de Venta
1. Abre el formulario principal
2. Selecciona "Nota de Venta"
3. Completa los campos:
   - Fecha de emisión
   - Número de nota de venta
   - Proveedor
   - Emitido con (Cédula/RUC)
   - Tipo de gasto
   - Total
   - Forma de pago
4. Clic en "Guardar"

### Buscar Registros
1. Haz clic en "Buscar"
2. Ingresa criterios (puedes combinar múltiples filtros)
3. Los resultados aparecen en una tabla
4. Haz doble clic para editar

### Generar Reportes
1. Clic en "Reportes"
2. Selecciona el tipo:
   - Facturas
   - Notas de Venta
   - General
3. Aplica filtros si deseas
4. Exporta a Excel o PDF

### Ver Dashboard
1. Clic en "Dashboard"
2. Se actualizará automáticamente con los datos
3. Visualiza indicadores clave y gráficos

---

## ⚙️ Requisitos del Sistema

| Requisito | Versión |
|-----------|----------|
| **Excel** | 2016, 2019, 2021, Microsoft 365 |
| **VBA** | 7.1+ (incluido en Excel) |
| **Windows** | 7, 8, 10, 11 |
| **Espacio en disco** | ~5 MB |
| **Carpeta destino** | `C:\informe\` (se crea automáticamente) |

### Habilitar Macros
1. Abre Excel
2. Archivo → Opciones → Centro de confianza
3. Haz clic en "Configuración del Centro de confianza"
4. Macros de VBA → Selecciona "Habilitar todas las macros"

---

## 📊 Campos de Base de Datos

### Hojas del Libro
- **Registros**: Almacena todos los documentos (Facturas y Notas de Venta)
- **TiposGasto**: Lista de categorías de gastos (editable)
- **Reportes**: Generación automática de reportes
- **Dashboard**: Indicadores y gráficos en tiempo real

### Estructura de Registros

| Campo | Factura | Nota Venta |
|-------|---------|----------|
| Código | ✅ Obligatorio | ✅ Obligatorio |
| Tipo Documento | ✅ | ✅ |
| Fecha Emisión | ✅ Obligatorio | ✅ Obligatorio |
| RUC Proveedor | ✅ Obligatorio | ❌ Vacío |
| Número Factura | ✅ Obligatorio | ❌ Vacío |
| Número Nota Venta | ❌ Vacío | ✅ Obligatorio |
| Proveedor | ✅ Obligatorio | ✅ Obligatorio |
| Emitido con | ❌ Vacío | ✅ Obligatorio |
| Tipo Gasto | ✅ Obligatorio | ✅ Obligatorio |
| Subtotal | ✅ Obligatorio | ❌ Vacío |
| IVA | ✅ Obligatorio | ❌ Vacío |
| Total | ✅ Obligatorio | ✅ Obligatorio |
| Forma Pago | ✅ Obligatorio | ✅ Obligatorio |
| Observaciones | ❌ Opcional | ❌ Opcional |

---

## 🔍 Búsqueda Inteligente

El sistema permite buscar por:
- Código (Fac0001, NV0001, etc.)
- Tipo de documento (Factura, Nota Venta)
- Fecha (rango o exacta)
- RUC del proveedor
- Número de factura
- Número de nota de venta
- Nombre del proveedor
- Tipo de gasto
- Forma de pago

Puedes **combinar múltiples filtros simultáneamente**.

---

## 💾 Exportación de Reportes

### A Excel
```
C:\informe\Reportes\Reporte_20260612_1430.xlsx
```

### A PDF
```
C:\informe\Reportes\Reporte_20260612_1430.pdf
```

El PDF incluye:
- Nombre del reporte
- Fecha de generación
- Filtros aplicados
- Tabla completa
- Totales y resumen

---

## 📈 Dashboard - Indicadores

**Totales:**
- Total de documentos
- Total Facturas
- Total Notas de Venta
- Total Gastos

**Análisis:**
- Promedio por documento
- Total por proveedor
- Total por tipo de gasto
- Total por forma de pago
- Gasto del mes actual
- Gasto anual

**Gráficos:**
- Gastos por proveedor
- Gastos por tipo de gasto
- Gastos por mes
- Comparativo Facturas vs Notas de Venta
- Top 10 proveedores
- Formas de pago
- Evolución mensual

---

## ✅ Validaciones Integradas

El sistema **impide automáticamente:**
- ❌ Guardar registros incompletos
- ❌ Números de factura duplicados
- ❌ Números de nota venta duplicados
- ❌ Totales negativos
- ❌ Fechas inválidas o futuras
- ❌ Campos obligatorios vacíos
- ❌ IVA negativo

---

## 🎨 Interfaz de Usuario

- ✨ Diseño moderno y profesional
- 🎯 UserForms centrados y responsivos
- 🔘 Botones con iconos intuitivos
- 🌈 Esquema de colores profesional
- ⌨️ Navegación con teclado (Tab, Enter, Escape)
- 📱 Compatible con Excel 2016+
- 🖱️ ListView para visualización de datos

---

## 🔐 Seguridad y Validaciones

✅ **Integridad de datos:**
- Validación automática de campos obligatorios
- Prevención de duplicados
- Validación de rangos numéricos
- Validación de formatos de fecha

✅ **Registro y auditoría:**
- Todos los cambios se almacenan en la hoja Registros
- Cada operación genera un código único
- Historial completo de transacciones

---

## 📝 Tipos de Gasto Predefinidos

El sistema incluye categorías comunes:
- Servicios Generales
- Mantenimiento Equipos
- Mantenimiento Vehicular
- Utiles de Oficina
- Combustible
- Arrendamiento
- Seguros
- Servicios Profesionales
- Publicidad
- Viáticos

**Puedes agregar nuevas categorías dinámicamente** desde los formularios.

---

## 🚨 Solución de Problemas

### Las macros no se ejecutan
1. ✅ Habilita las macros en Centro de confianza de Excel
2. ✅ Guarda el archivo como `.xlsm` (macro habilitado)

### Las carpetas no se crean
1. ✅ Verifica que tengas permisos de escritura en `C:\`
2. ✅ Ejecuta Excel como administrador

### Los gráficos no aparecen en Dashboard
1. ✅ Abre el archivo en Excel y selecciona la hoja Dashboard
2. ✅ Hace clic en los gráficos para refrescar

### Error "División por cero" en reportes
1. ✅ Asegúrate de tener al menos un registro antes de generar reportes
2. ✅ Los reportes vacíos muestran 0 automáticamente

---

## 📞 Soporte y Contribuciones

### Reportar Issues
1. Abre una [issue en GitHub](https://github.com/Girney/sistema-control-gastos-excel/issues)
2. Describe el problema con detalle
3. Incluye versión de Excel y Windows

### Contribuir
```bash
# Fork el proyecto
# Crea una rama (git checkout -b feature/mejora)
# Commit cambios (git commit -m 'Agrega mejora')
# Push (git push origin feature/mejora)
# Abre Pull Request
```

---

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Ver archivo [LICENSE](LICENSE) para detalles.

---

## 🎓 Documentación Completa

Para una guía detallada paso a paso, consulta: **[GUIA_USO.md](GUIA_USO.md)**

---

## 👨‍💻 Autor

**Desarrollado como solución profesional de gestión empresarial**

---

## 🔄 Changelog

### v1.0.0 (2026-06-12)
- ✨ Release inicial
- ✅ CRUD completo de Facturas y Notas de Venta
- ✅ Búsqueda inteligente con filtros
- ✅ Reportes con agrupación
- ✅ Dashboard con 12+ indicadores
- ✅ Exportación a Excel y PDF
- ✅ ComboBox con autocompletado
- ✅ Validaciones completas

---

## 💡 Tips y Trucos

1. **Atajo de teclado principal**: `Ctrl + Shift + I` abre el formulario
2. **Búsqueda rápida**: Usa Enter para buscar automáticamente
3. **Editar registros**: Doble clic en resultado de búsqueda
4. **Tipos de gasto**: Se agregan automáticamente cuando escribes uno nuevo
5. **Reportes mensuales**: Filtra por mes en la sección de fechas
6. **Exportar todo**: Selecciona "Todos" en reportes para consolidado

---

**¡Gracias por usar Sistema de Control de Gastos! 💼✨**