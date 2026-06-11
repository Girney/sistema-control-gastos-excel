# 📊 Sistema de Control de Gastos - Excel VBA

Sistema profesional de control de gastos para Microsoft Excel utilizando VBA y UserForms. Permite registrar, consultar, editar, eliminar y reportar gastos mediante Facturas y Notas de Venta.

## 🎯 Características

### ✅ Gestión de Documentos
- **Facturas**: Códigos automáticos (Fac0001, Fac0002, etc.)
- **Notas de Venta**: Códigos automáticos (NV0001, NV0002, etc.)
- Formularios dinámicos según tipo de documento
- CRUD completo (Crear, Leer, Editar, Eliminar)
- Búsqueda inteligente

### 📊 Reportería Avanzada
- Reportes filtrados por: tipo documento, proveedor, tipo gasto, fecha
- Búsqueda inteligente con autocompletado
- Exportación a Excel
- Exportación a PDF

### 📈 Dashboard Dinámico
- Indicadores en tiempo real
- Gráficos de gastos por categoría
- Análisis por proveedor
- Top 10 proveedores

### 🔍 Búsqueda Inteligente
- ComboBox con filtrado automático
- Búsqueda sin distinción de mayúsculas/minúsculas
- Coincidencias parciales
- Creación de nuevos tipos de gasto

## 📋 Requisitos

- Microsoft Excel 2016, 2019, 2021 o Microsoft 365
- VBA habilitado
- Macros habilitadas
- Windows 10/11

## 🚀 Instalación Rápida

### Paso 1: Descargar Archivos
```bash
git clone https://github.com/Girney/sistema-control-gastos-excel.git
```

### Paso 2: Abrir Excel
1. Navega a la carpeta `/Template`
2. Abre `informe_factura.xlsm`
3. Habilita macros cuando se solicite

### Paso 3: Inicialización Automática
El sistema creará automáticamente:
- Carpeta: `C:\informe\`
- Archivo: `C:\informe\informe_factura.xlsx`
- Hojas: Registros, TiposGasto, Reportes, Dashboard

## 📁 Estructura del Proyecto

```
sistema-control-gastos-excel/
├── README.md
├── INSTALACION.md
├── VBA/
│   ├── Modulos/
│   │   ├── ModuloInitDB.bas         # Inicialización de BD
│   │   ├── ModuloCRUD.bas           # Operaciones CRUD
│   │   ├── ModuloReportes.bas       # Generación de reportes
│   │   ├── ModuloExportacion.bas    # Export Excel/PDF
│   │   ├── ModuloUtilidades.bas     # Funciones auxiliares
│   │   └── ModuloDashboard.bas      # Actualización dashboard
│   ├── Clases/
│   │   ├── ClaseDocumento.cls       # Clase documento
│   │   ├── ClaseGasto.cls           # Clase tipo gasto
│   │   └── ClaseReporte.cls         # Clase reporte
│   └── Formularios/
│       ├── FormularioPrincipal.frm  # Formulario principal
│       ├── FormularioReportes.frm   # Formulario reportes
│       └── FormularioBusqueda.frm   # Formulario búsqueda
├── Template/
│   └── informe_factura.xlsm
└── Documentacion/
    ├── GUIA_USUARIO.md
    ├── GUIA_DESARROLLADOR.md
    └── CAMBIOS.md
```

## 🎮 Uso Básico

### Iniciar el Sistema
1. Abrir Excel
2. Presionar `Ctrl + Alt + F` o usar menú desplegable
3. El formulario principal se abrirá automáticamente

### Registrar Factura
1. Seleccionar "Factura"
2. Completar campos (el código se genera automáticamente)
3. Hacer clic en "Guardar"

### Registrar Nota de Venta
1. Seleccionar "Nota de Venta"
2. Completar campos (el código se genera automáticamente)
3. Hacer clic en "Guardar"

### Ver Reportes
1. Hacer clic en botón "Reportes"
2. Aplicar filtros deseados
3. Exportar a Excel o PDF

## 🔐 Validaciones

El sistema valida automáticamente:
- ✅ Campos obligatorios no vacíos
- ✅ Números de factura únicos
- ✅ Números de Nota de Venta únicos
- ✅ Sin valores negativos
- ✅ Fechas válidas
- ✅ Importes numéricos correctos
- ✅ Total no puede ser cero

## 📋 Campos - Factura

- Código interno (automático: Fac0001, Fac0002, etc.)
- Fecha de emisión
- Proveedor
- Tipo de identificación (Cédula/RUC)
- Número de factura
- Tipo de gasto
- Subtotal
- IVA
- Total
- Forma de pago
- Observaciones

## 📋 Campos - Nota de Venta

- Código interno (automático: NV0001, NV0002, etc.)
- Fecha de emisión
- Proveedor
- Número de Nota de Venta
- Nombre de Proveedor
- Tipo de gasto
- Total
- Forma de pago
- Observaciones

## 💳 Formas de Pago

- Efectivo
- Transferencia
- Tarjeta de Crédito
- Tarjeta de Débito
- Cheque
- Otro

## 📊 Tipos de Gasto Iniciales

Alimentación | Transporte | Hospedaje | Materiales de oficina
Servicios básicos | Combustible | Mantenimiento | Equipos tecnológicos
Capacitación | Papelería | Internet | Telefonía
Software | Hardware | Publicidad | Impresiones
Repuestos | Herramientas | Seguridad | Otros

## 📈 Dashboard - Indicadores

### Métricas
- Total de documentos registrados
- Total de Facturas
- Total de Notas de Venta
- Total de gastos
- Total por proveedor
- Total por tipo de gasto

### Gráficos
1. **Gastos por Proveedor** - Gráfico de barras
2. **Gastos por Tipo de Gasto** - Gráfico circular
3. **Gastos por Mes** - Gráfico de líneas
4. **Facturas vs Notas de Venta** - Gráfico comparativo
5. **Top 10 Proveedores** - Ranking

## 🔄 Ciclo de Vida del Documento

```
┌─────────────┐
│   NUEVO     │  Crear nuevo registro
└──────┬──────┘
       │
       ▼
┌─────────────┐
│  EDITAR     │  Modificar datos
└──────┬──────┘
       │
       ▼
┌─────────────┐
│  GUARDAR    │  Validar y almacenar
└──────┬──────┘
       │
       ▼
┌─────────────┐
│  REGISTRADO │  En base de datos
└──────┬──────┘
       │
       ▼ (Opcional)
┌─────────────┐
│  ELIMINAR   │  Remover registro
└─────────────┘
```

## 📁 Carpetas Automáticas

El sistema crea automáticamente:

```
C:\informe\
├── informe_factura.xlsx      (Base de datos)
└── Reportes\                 (Exportaciones)
    ├── Excel_YYYYMMDD_HHmmss.xlsx
    └── PDF_YYYYMMDD_HHmmss.pdf
```

## 🛠️ Instalación Manual desde Código

Si prefieres instalar módulo por módulo:

1. Crear nuevo libro Excel (`.xlsm`)
2. Abrir Editor VBA (`Alt + F11`)
3. Importar archivos:
   - `/VBA/Modulos/*.bas`
   - `/VBA/Clases/*.cls`
   - `/VBA/Formularios/*.frm`
4. Ejecutar: `ModuloInitDB.InicializarSistema()`

## 🐛 Solución de Problemas

### Macros no se ejecutan
- Verificar que Excel esté en modo de "Contenido habilitado"
- Guardar como `.xlsm` (no `.xlsx`)

### Carpeta C:\informe no se crea
- Verificar permisos de escritura en C:\
- Ejecutar Excel como administrador

### Errores de validación
- Revisar que los campos obligatorios estén completados
- Verificar formato de fechas (DD/MM/YYYY)
- Asegurar valores numéricos positivos

## 📞 Soporte

Para reportar errores o sugerir mejoras:
- [Abrir Issue](https://github.com/Girney/sistema-control-gastos-excel/issues)
- Describir el problema detalladamente
- Adjuntar screenshot si es posible

## 📄 Licencia

Este proyecto está disponible bajo licencia MIT.

## 👨‍💻 Autor

Girney

---

**Última actualización**: Junio 2026
**Versión**: 1.0.0
