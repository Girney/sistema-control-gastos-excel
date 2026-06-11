# 🔧 Guía de Instalación Detallada

## Requisitos Previos

- ✅ Windows 10 o superior
- ✅ Microsoft Excel 2016 o superior
- ✅ .NET Framework 4.5+
- ✅ Acceso de escritura en `C:\`
- ✅ Macros habilitadas en Excel

## Instalación Paso a Paso

### Opción A: Instalación Rápida (Recomendado)

**Paso 1: Descargar**
```bash
git clone https://github.com/Girney/sistema-control-gastos-excel.git
cd sistema-control-gastos-excel
```

**Paso 2: Abrir Archivo Template**
1. Navega a carpeta `Template`
2. Haz doble clic en `informe_factura.xlsm`
3. Si aparece aviso de seguridad, haz clic en "Habilitar contenido"

**Paso 3: Inicialización Automática**
- El sistema se inicializará automáticamente
- Se crearán carpetas en `C:\informe\`
- Se generarán todas las hojas necesarias
- Verás confirmación en pantalla

### Opción B: Instalación Manual desde Código

**Paso 1: Crear Nuevo Libro**
1. Abre Excel
2. Crea un nuevo libro en blanco
3. Guarda como: `C:\Users\[TuUsuario]\informe_factura.xlsm`
4. **Importante**: Guarda como `.xlsm` (macro habilitado)

**Paso 2: Abrir Editor VBA**
1. Presiona `Alt + F11`
2. Se abrirá el editor Visual Basic

**Paso 3: Importar Módulos**

Para cada archivo en `/VBA/Modulos/`:
1. Haz clic derecho en la carpeta del proyecto (izquierda)
2. Selecciona "Importar archivo"
3. Navega a `/VBA/Modulos/[archivo].bas`
4. Repite para todos los módulos:
   - `ModuloInitDB.bas`
   - `ModuloCRUD.bas`
   - `ModuloReportes.bas`
   - `ModuloExportacion.bas`
   - `ModuloUtilidades.bas`
   - `ModuloDashboard.bas`

**Paso 4: Importar Clases**

Para cada archivo en `/VBA/Clases/`:
1. Haz clic derecho en la carpeta del proyecto
2. Selecciona "Importar archivo"
3. Navega a `/VBA/Clases/[archivo].cls`
4. Importa:
   - `ClaseDocumento.cls`
   - `ClaseGasto.cls`
   - `ClaseReporte.cls`

**Paso 5: Importar Formularios**

Para cada archivo en `/VBA/Formularios/`:
1. Haz clic derecho en carpeta "Formularios" en el árbol
2. Selecciona "Importar archivo"
3. Navega a `/VBA/Formularios/[archivo].frm`
4. Importa:
   - `FormularioPrincipal.frm`
   - `FormularioReportes.frm`
   - `FormularioBusqueda.frm`

**Paso 6: Ejecutar Inicialización**

1. Presiona `Ctrl + Home` para cerrar el editor
2. En Excel, presiona `Ctrl + Alt + F`
3. Se abrirá el formulario principal
4. El sistema se inicializará automáticamente

## Habilitación de Macros

### En Excel 2016/2019/2021

1. Abre Excel
2. Ve a: `Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza`
3. Selecciona: `Configuración de macros`
4. Elige: `Habilitar todas las macros`
5. OK

### En Microsoft 365

1. Ve a: `Archivo > Más configuración > Centro de confianza`
2. Selecciona: `Configuración de macros`
3. Marca: `Habilitar todas las macros`
4. Confirma

## Verificación de Instalación

### Checklist de Instalación

- [ ] El archivo `informe_factura.xlsm` está guardado
- [ ] Se puede abrir sin errores
- [ ] Presionar `Alt + F11` abre el editor VBA
- [ ] Existen 6 módulos en el proyecto
- [ ] Existen 3 clases en el proyecto
- [ ] Existen 3 formularios en el proyecto
- [ ] La carpeta `C:\informe\` existe
- [ ] El archivo `C:\informe\informe_factura.xlsx` existe
- [ ] Las hojas existen: Registros, TiposGasto, Reportes, Dashboard
- [ ] El formulario principal se abre con `Ctrl + Alt + F`

## Prueba de Funcionamiento

### Prueba 1: Abrir Formulario Principal
```
1. Presiona: Ctrl + Alt + F
2. Resultado esperado: Se abre formulario con dos opciones (Factura/NV)
```

### Prueba 2: Crear Factura
```
1. Selecciona: Factura
2. Completa campos:
   - Fecha: Hoy
   - Proveedor: "Prueba"
   - ID: Cédula
   - No. Factura: 001
   - Tipo Gasto: Alimentación
   - Subtotal: 100
   - IVA: 12
   - Total: 112
   - Pago: Efectivo
3. Clic: "Guardar"
4. Resultado: Mensaje de confirmación
```

### Prueba 3: Crear Nota de Venta
```
1. Selecciona: Nota de Venta
2. Completa campos:
   - Fecha: Hoy
   - Proveedor: "Prueba"
   - No. NV: 001
   - Nombre Proveedor: "Empresa Prueba"
   - Tipo Gasto: Transporte
   - Total: 50
   - Pago: Transferencia
3. Clic: "Guardar"
4. Resultado: Mensaje de confirmación
```

### Prueba 4: Ver Reportes
```
1. En formulario principal, clic: "Reportes"
2. Se abre formulario de reportes
3. Aplica filtro de tipo: Factura
4. Resultado: Muestra factura registrada
```

### Prueba 5: Ver Dashboard
```
1. Haz clic en la hoja "Dashboard"
2. Resultado esperado:
   - Indicadores muestran valores
   - Gráficos se visualizan
   - Todo se actualiza automáticamente
```

## Solución de Problemas Comunes

### Problema: "Macros deshabilitadas"
**Solución:**
1. Guardar archivo como `.xlsm`
2. Habilitar macros en Configuración de Excel
3. Reiniciar Excel

### Problema: "Error al crear carpeta C:\informe\"
**Solución:**
1. Ejecutar Excel como administrador
2. Verificar permisos en `C:\`
3. Si es Windows protegido, usar `C:\Users\[Usuario]\informe\`

### Problema: "No se crean las hojas"
**Solución:**
1. Verificar que el archivo es `.xlsm`
2. Ejecutar `ModuloInitDB.InicializarSistema()` manualmente
3. Presionar `Ctrl + Alt + F` para activar

### Problema: "Formulario no se abre"
**Solución:**
1. Verificar que `FormularioPrincipal` esté importado
2. Revisar que no haya errores de sintaxis (F7)
3. Ejecutar manualmente: `FormularioPrincipal.Show`

### Problema: "ComboBox de gastos vacío"
**Solución:**
1. Verificar que hoja `TiposGasto` existe
2. Revisar que datos de tipos estén presentes
3. Ejecutar: `ModuloUtilidades.CargarTiposGasto()`

## Actualización

### Actualizar a Nueva Versión

1. Descargar última versión: `git pull`
2. Reemplazar archivos en `/VBA/`
3. Importar módulos nuevamente
4. Guardar como macro habilitado
5. Ejecutar inicialización

## Desinstalación

### Para desinstalar el sistema:

1. Eliminar: `C:\informe\` (carpeta completa)
2. Eliminar: `informe_factura.xlsm`
3. Limpiar caché de Excel

## Soporte

Si encuentras problemas durante la instalación:

1. Verifica los requisitos previos
2. Revisa la sección "Solución de Problemas"
3. [Abre un Issue](https://github.com/Girney/sistema-control-gastos-excel/issues) con:
   - Versión de Excel
   - Sistema Operativo
   - Descripción del error
   - Captura de pantalla si es posible

---

**¿Necesitas ayuda?** Consulta la [Guía de Usuario](GUIA_USUARIO.md)
