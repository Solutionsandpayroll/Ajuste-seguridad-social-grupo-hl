# Ajuste Contable – Seguridad Social (Grupo HL)

Herramienta web interna de **Solutions & Payroll** para automatizar la generación de la hoja de ajuste contable de seguridad social a partir de los archivos de nómina CAUSADO y PAGADO.

---

## Tabla de contenidos

1. [¿Qué hace esta herramienta?](#qué-hace-esta-herramienta)
2. [Tecnologías](#tecnologías)
3. [Instalación y ejecución local](#instalación-y-ejecución-local)
4. [Flujo de trabajo paso a paso](#flujo-de-trabajo-paso-a-paso)
   - [Paso 1 – Generar CauPag](#paso-1--generar-caupag)
   - [Paso 2 – Generar Ajuste](#paso-2--generar-ajuste)
5. [Lógica de negocio detallada](#lógica-de-negocio-detallada)
   - [Agrupación por pivot](#agrupación-por-pivot)
   - [Merge de registros duplicados](#merge-de-registros-duplicados)
   - [Filtro final](#filtro-final)
   - [Integración con Base Empleados Conceptos](#integración-con-base-empleados-conceptos)
6. [Estructura del proyecto](#estructura-del-proyecto)
7. [Plantilla Excel](#plantilla-excel)

---

## ¿Qué hace esta herramienta?

Dado el archivo fuente de nómina (`INFRA - CAUSADO - PAGADO.xlsx`), la herramienta:

1. **Paso 1** – Lee las filas CAUSADO y PAGADO del archivo fuente y las escribe en la hoja `CauPag` de la plantilla, dejando las tablas dinámicas listas para que el usuario las actualice manualmente en Excel.
2. **Paso 2** – Toma el archivo generado en el Paso 1 (con las tablas dinámicas ya actualizadas), agrupa los valores por empleado y entidad, detecta diferencias entre causado y pagado, y escribe los registros resultantes en la hoja `Ajuste`. Opcionalmente puede incorporar datos de la **Base Empleados Conceptos** para ajustar el valor causado con descuentos adicionales.

---

## Tecnologías

| Librería | Uso |
|---|---|
| React 18 + Vite 5 | Interfaz de usuario |
| ExcelJS | Lectura del archivo XLSX (hoja CauPag) |
| JSZip | Escritura directa del XML dentro del XLSX (preserva tablas dinámicas) |
| xlsx (SheetJS) | Lectura de archivos `.xls` / `.xlsx` (Base Empleados Conceptos) |
| vite-plugin-node-polyfills | Compatibilidad con APIs de Node en el navegador |

---

## Instalación y ejecución local

```bash
# 1. Clonar el repositorio
git clone https://github.com/Solutionsandpayroll/Ajuste-seguridad-social-grupo-hl.git
cd Ajuste-seguridad-social-grupo-hl

# 2. Instalar dependencias
npm install

# 3. Iniciar servidor de desarrollo
npm run dev
```

Abrir el navegador en `http://localhost:5173`.

Para generar la versión de producción:

```bash
npm run build
```

Los archivos compilados quedan en la carpeta `dist/`.

---

## Flujo de trabajo paso a paso

### Paso 1 – Generar CauPag

1. En la aplicación selecciona **"Paso 1"**.
2. Escribe la **fecha** del período (se usa en el nombre del archivo descargado).
3. Sube el archivo fuente de nómina (p. ej. `INFRA - CAUSADO - PAGADO.xlsx`).  
   - Puedes arrastrarlo sobre la zona de carga o hacer clic para buscarlo.
4. Presiona **"Generar Paso 1"**.
5. Se descargará automáticamente el archivo `Ajuste Contable Paso1 - YYYY-MM-DD.xlsx`.
6. **Abre ese archivo en Excel** y actualiza las 6 tablas dinámicas de la hoja `Dinamica`  
   (clic derecho sobre cada tabla → *Actualizar*, o usa *Actualizar todo* en la pestaña Datos).
7. Guarda y cierra el archivo.

> **Importante:** No modifiques la hoja `CauPag` manualmente. Las tablas dinámicas leen sus datos de ahí.

---

### Paso 2 – Generar Ajuste

1. En la aplicación selecciona **"Paso 2"**.
2. Sube el archivo generado en el Paso 1 (con las tablas dinámicas ya actualizadas).
3. *(Opcional)* Sube el archivo **Base Empleados Conceptos** (`.xls` o `.xlsx`) para incluir ajustes de conceptos C396, C397, C398 en el valor causado.
4. Presiona **"Generar Paso 2"**.
5. Se descargará `Ajuste Contable Final.xlsx` con la hoja `Ajuste` completa y, si se subió la base, la hoja `Hoja1` con los conceptos.

---

## Lógica de negocio detallada

### Agrupación por pivot

La herramienta evalúa **6 grupos** definidos en `PIVOT_CONFIGS`:

| Grupo | Tipo en Ajuste | Col. Administradora | Col. Valor |
|---|---|---|---|
| AFP | PENSIÓN | 51 | 62 |
| EPS | SALUD | 64 | 68 |
| ARL | ARL | 75 | 82 |
| CCF | CCF | 84 | 87 |
| SENA | SENA | 84 | 90 |
| ICBF | ICBF | 84 | 92 |

Para cada grupo se agrupa por `NoId + Administradora`, sumando causado y pagado con `Math.round`.

---

### Merge de registros duplicados

Cuando un mismo `NoId` aparece más de una vez para el mismo tipo (p. ej. varias filas AFP del mismo empleado), se aplica la siguiente lógica en orden:

**Caso A – Admin única** (solo un registro después del agrupado): pasa directo a Ajuste con `Entidad Causada = Entidad Pagada = esa admin`.

**Caso B – Varios admins, uno con valor NINGUNA**:  
Se elimina el registro con admin `NINGUNA`. Si solo queda uno → Caso A. Si quedan varios → Caso D.

**Caso C – Todos los admins son `(en blanco)` o `NINGUNA`**: se omite el empleado.

**Caso D – Varios admins reales distintos**:  
Se fusionan en un único registro:
- `Valor Causado` = suma de todos los causados válidos.
- `Valor Pagado` = suma de todos los pagados válidos.
- `Entidad Causada` = admin del registro con **mayor** valor causado.
- `Entidad Pagada` = admin del registro con **mayor** valor pagado.

Ejemplo con 3 filas AFP:
```
1082866503  PROTECCION   causado=227.700  pagado=0
1082866503  PROTECCION   causado=0        pagado=357.200
1082866503  PORVENIR     causado=131.200  pagado=0
→ causadoTotal = 358.900  pagadoTotal = 357.200
→ Entidad Causada = PROTECCION (mayor causado = 227.700)
→ Entidad Pagada  = PROTECCION (único con pagado > 0)
```

Ejemplo con 2 filas:
```
17957880  COLPENSIONES  causado=0       pagado=78.200
17957880  PORVENIR      causado=78.000  pagado=0
→ causadoTotal = 78.000  pagadoTotal = 78.200
→ Entidad Causada = PORVENIR   (mayor causado)
→ Entidad Pagada  = COLPENSIONES (mayor pagado)
```

---

### Filtro final

**Después** de calcular todos los registros (y aplicar ajustes de Hoja1), se eliminan los registros que cumplan **ambas** condiciones:

- `Dif (Valor Pagado − Valor Causado) = 0`
- `Entidad Causada = Entidad Pagada`

Si las entidades son distintas aunque la diferencia sea cero (cruce de entidades), **el registro se conserva**.

---

### Integración con Base Empleados Conceptos

Si se sube este archivo opcional, la herramienta:

1. Lee desde la fila 11 del primer hoja, filtrando conceptos `C396`, `C397`, `C398`.
2. Renombra "FONDO SOLIDARIDAD" → `APORTE PENSION`.
3. Mapea los conceptos a tipo: APORTE PENSION/FONDO SOLIDARIDAD → `PENSIÓN`; APORTE SALUD → `SALUD`.
4. Escribe esas filas en la hoja `Hoja1` (columnas A–H) con fórmula en columna I: `=+C{n}*-1`.
5. En la hoja `Ajuste`, columna `Valor Causado`, añade a la fórmula las referencias a `Hoja1!I{fila}` sumadas al causado original:  
   `=+{causado}+Hoja1!I{fila1}+Hoja1!I{fila2}+...`

Esto permite que los descuentos de seguridad social adicionales se reflejen automáticamente en el valor causado final del ajuste.

---

## Estructura del proyecto

```
├── public/
│   └── plantilla-ajuste.xlsx      # Plantilla Excel con hojas: CauPag, Dinamica, Ajuste, Hoja1, Hoja2
├── src/
│   ├── App.jsx                    # Interfaz principal (pasos 1 y 2, drag & drop)
│   ├── App.css                    # Estilos corporativos
│   ├── main.jsx                   # Punto de entrada React
│   └── utils/
│       └── procesarExcel.js       # Toda la lógica de procesamiento Excel
├── index.html
├── vite.config.js
└── package.json
```

---

## Plantilla Excel

La plantilla `public/plantilla-ajuste.xlsx` contiene las siguientes hojas:

| Hoja | Descripción |
|---|---|
| `DatosPruebaEmp` | Hoja auxiliar con datos de referencia |
| `CauPag` | Destino de las filas CAUSADO/PAGADO del fuente (99 columnas) |
| `Dinamica` | 6 tablas dinámicas sobre CauPag (AFP, EPS, ARL, CCF, SENA, ICBF) |
| `Ajuste` | Hoja de ajuste contable generada automáticamente |
| `Hoja1` | Datos de Base Empleados Conceptos (C396/C397/C398) |
| `Hoja2` | Hoja auxiliar adicional |

> La plantilla **no debe modificarse** sin coordinar cambios en `procesarExcel.js`, ya que las columnas y posiciones están mapeadas con índices exactos.

## ✨ Características Incluidas

- ✅ **Header corporativo** con logo y bienvenida
- ✅ **Diseño profesional** con colores y estilos de S&P
- ✅ **Sección de ayuda colapsable** (opcional)
- ✅ **Sistema de cards** con animaciones suaves
- ✅ **Footer** corporativo
- ✅ **100% responsive** para móviles y desktop
- ✅ **Animaciones** de entrada elegantes
- ✅ **Variables CSS** fáciles de personalizar

## 🚀 Cómo Usar Este Template

### Opción 1: Copiar para Nuevo Proyecto

```bash
# 1. Copiar la carpeta completa
cp -r syp-react-template mi-nuevo-proyecto

# 2. Entrar al nuevo proyecto
cd mi-nuevo-proyecto

# 3. Instalar dependencias
npm install

# 4. Iniciar desarrollo
npm run dev
```

### Opción 2: Clonar y Modificar

```bash
# 1. Copiar todo el contenido
Copy-Item -Path "syp-react-template" -Destination "nuevo-proyecto" -Recurse

# 2. Cambiar nombre en package.json
# Edita la línea: "name": "tu-nombre-proyecto"

# 3. Instalar y ejecutar
cd nuevo-proyecto
npm install
npm run dev
```

## 📝 Estructura del Template

```
syp-react-template/
├── public/
│   └── Logo syp.png          # Logo corporativo S&P
├── src/
│   ├── App.jsx               # Componente principal (limpio)
│   ├── App.css               # Estilos completos
│   ├── index.css             # Estilos globales
│   └── main.jsx              # Entry point
├── index.html                # HTML base con favicon
├── package.json              # Dependencias mínimas
└── vite.config.js            # Configuración Vite
```

## 🎯 Personalización Rápida

### 1. Cambiar Título de la App

Edita `src/App.jsx` línea ~20:
```jsx
<p className="subtitle">Tu Nuevo Título</p>
```

### 2. Modificar Mensaje de Bienvenida

Edita `src/App.jsx` línea ~30:
```jsx
<span>Bienvenido, Tu Usuario</span>
```

### 3. Personalizar Colores

Edita `src/App.css`, variables CSS al inicio:
```css
:root {
  --primary: #2563eb;        /* Azul principal */
  --primary-dark: #1e40af;   /* Azul oscuro */
  /* ... más colores */
}
```

### 4. Agregar tu Lógica

En `src/App.jsx`, dentro del `<div className="card-body">`:
- Agrega tus estados con `useState`
- Crea tus funciones
- Añade tus componentes de formulario

## 📦 Agregar Dependencias

Según lo que necesites para tu proyecto:

```bash
# Para procesar archivos Excel
npm install xlsx exceljs file-saver

# Para formularios
npm install react-hook-form

# Para hacer requests
npm install axios

# Para routing
npm install react-router-dom

# etc...
```

## 🎨 Componentes Disponibles

### Sección de Ayuda Colapsable

Si no la necesitas, puedes eliminar todo el bloque:
```jsx
<div className="help-section">
  {/* ... */}
</div>
```

### Form Groups

```jsx
<div className="form-group">
  <label className="label">
    {/* Icono SVG */}
    Tu Label
  </label>
  <input className="select-input" />
</div>
```

### Botones

```jsx
<button className="btn-primary">
  {/* Icono SVG */}
  Texto del Botón
</button>
```

## 🌈 Estilos Predefinidos

Clases disponibles en `App.css`:
- `.card` - Contenedor con sombra
- `.form-section` - Espaciado de formularios
- `.form-group` - Grupo de campo
- `.label` - Label con icono
- `.select-input` - Input/Select estilizado
- `.btn-primary` - Botón principal
- `.btn-remove` - Botón eliminar
- `.drop-zone` - Zona drag & drop
- `.modal-overlay` - Overlay de modal
- `.help-section` - Sección colapsable

## 💡 Tips

1. **Mantén limpio el App.jsx** - Crea componentes separados si crece mucho
2. **Usa las variables CSS** - No modifiques los colores directamente
3. **Los SVG están inline** - Puedes cambiarlos fácilmente o usar íconos de librerías
4. **Las animaciones ya están configuradas** - Se activarán automáticamente

## 📚 Recursos

- [Documentación React](https://react.dev/)
- [Documentación Vite](https://vitejs.dev/)
- [Iconos SVG](https://feathericons.com/)
- [Colores](https://tailwindcss.com/docs/customizing-colors)

## 🔒 No Subir a Git

Si inicias Git en tu nuevo proyecto, asegúrate de tener `.gitignore`:
```
node_modules
dist
.env
```

## 📄 Licencia

© 2026 Solutions & Payroll. Template de uso interno.

---

**¡Listo para crear tu próximo proyecto!** 🚀
