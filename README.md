# ZipGrade Controller 📝

Biblioteca de **Google Apps Script** para sincronizar resultados de exámenes de ZipGrade, consolidar fórmulas automáticas y gestionar asignaturas/sesiones dinámicamente en Google Sheets, con una interfaz premium en HTML/JS.

---

## 🚀 Características Principales

- **Sincronización y Alineación Automática con ZipGrade**: Parsea y procesa exportaciones CSV de ZipGrade. Alinea automáticamente las notas con la lista maestra de estudiantes (`S1`) y detecta inconsistencias (estudiantes faltantes, estudiantes no registrados).
- **Gestor Interactivo de Asignaturas y Sesiones**: Modal HTML con arrastrar-y-soltar para reordenar asignaturas entre varias hojas de sesión, actualizar cantidades de preguntas y añadir/eliminar columnas de evaluación.
- **Hojas de Sesión Dinámicas**: Crea automáticamente las hojas de sesión (`S1`, `S2`, etc.) con formato estándar. Los nombres de las sesiones secundarias se vinculan dinámicamente a la lista maestra mediante fórmulas.
- **Fórmulas y Formato Automático**: Construye dinámicamente fórmulas de TOTAL y PROMEDIO con cebra (filas alternadas). Garantiza que las filas de resumen **TOTAL** y **PROMEDIO** siempre se escriban, incluso si la hoja tenía filas faltantes. Formato uniforme: fuente **Nunito**, tamaño 12 en datos, fila TOTAL azul y fila PROMEDIO celeste.
- **Herramientas del Consolidado**: Ordenar por TOTAL (de mayor a menor) o restaurar el orden alfabético con un clic; el formato, las fórmulas y los resaltados se vuelven a aplicar automáticamente tras reordenar.
- **Resaltado del Mejor Promedio**: Resalta al estudiante con el promedio general más alto.
- **Recálculo y Limpieza**: Recalcula todas las fórmulas con un clic y una utilidad "limpiar filas vacías" mantiene las hojas ordenadas y libres de errores `#REF!`.
- **Importar/Exportar Configuración JSON**: Permite guardar, respaldar e importar toda la estructura de evaluación (orden de asignaturas, sesiones y cantidad de preguntas) mediante JSON.

---

## 🧱 Reglas y Estructura del Libro (IMPORTANTE)

Para que el sistema funcione correctamente, el libro de cálculo debe cumplir estas reglas:

### 1. `S1` es la Lista Maestra de Estudiantes
- La hoja **`S1`** contiene la lista definitiva de estudiantes.
- **Columna A** = Apellidos, **Columna B** = Nombres, a partir de la **fila 2**.
- Es la **única** hoja donde los nombres se escriben directamente. En las demás sesiones y en el Consolidado los nombres son **fórmulas** que apuntan a `S1` (siempre dinámicos).
- El número de estudiantes se lee desde aquí: fila 2 de S1 = estudiante 1.

### 2. Hojas de Sesión (`S1`, `S2`, `S3`…)
- Los nombres deben seguir el patrón `S` + número (`S1`, `S2`, …). Cualquier otra hoja no se considera sesión.
- **Columnas A y B**: apellidos y nombres (en `S1` directos; en las demás, fórmulas hacia `S1`).
- **Columnas C en adelante**: una columna por pregunta. La fila 1 tiene los encabezados `#1`, `#2`, … (fondo azul `#35528C`, letra blanca, negrita).
- Pestaña de color intercalada: sesiones impares morado (`#8D7BA6`), sesiones pares rojo (`#B23A48`).

### 3. Hoja Consolidado (Contiene "Consolidado" en el nombre)
- **Fila 1**: encabezado de sesión combinado por bloques (`SESIÓN 1`, `SESIÓN 2`, …).
- **Filas 2–3**: encabezado combinado de cada asignatura (nombre + rango, y opcionalmente `Habilidades:`), fondo `#527FD9`, letra blanca.
- **Filas 4 en adelante**: estudiantes. **La fila 4 = estudiante 1 = fila 2 de S1** (regla de desplazamiento: `fila Consolidado = fila S1 + 2`).
- **Columna TOTAL** (fondo `#93ABD9`, letra blanca, formato `0`) y **columna PROMEDIO** (formato `0.00`) cierran las columnas de datos.
- **Fila TOTAL** (fila `4 + N° estudiantes`): suma de cada columna; fondo `#93ABD9`, letra blanca.
- **Fila PROMEDIO** (fila siguiente): promedio de cada columna; fondo `#D9E1F2`, letra negra, formato `0.00`.
- Pestaña azul oscuro (`#223459`).
- La fila **TOTAL** y la fila **PROMEDIO** deben estar SIEMPRE en su posición correcta (justo después del último estudiante); el sistema las corrige automáticamente al recalcular.

### 4. Hoja Fuente de Datos (Pegar el CSV de ZipGrade)
- Hoja preferida: **`Sheet1`**. Si no existe, se usa la primera hoja que **no** sea sesión (`S#`) ni Consolidado.
- El CSV se pega desde la **fila 2** (la fila 1 se ignora). Estructura: **A** = apellido, **B** = nombre, **C en adelante** = nota por pregunta.
- El sistema ordena los datos alfabéticamente por apellido antes de procesarlos.

### 5. Regla del Nombre del Archivo (Texto del Curso)
- La celda **A1 del Consolidado** muestra el curso de forma 100% dinámica.
- Se extrae automáticamente de lo que esté **después del carácter `|`** en el nombre del libro de cálculo (sin importar el prefijo que tenga antes, ya sea `R_PENSAR`, `E.S.A.S`, `ZipGrade`, etc.):
  - Ejemplo: `R_PENSAR | 3° - 5°` → A1 = `3° - 5°`.
  - Ejemplo: `E.S.A.S | 601` → A1 = `601`.
  - Ejemplo: `ZipGrade - 9°A | MATEMÁTICAS` → A1 = `MATEMÁTICAS`.
- Se actualiza automáticamente tanto al abrir el documento (`onOpen`) como al recalcular (`⚡ Recalcular y Corregir Fórmulas`).

### 6. Condiciones de Validación (si no se cumplen, el sistema avisa)
- Debe existir una hoja Consolidado y las filas TOTAL/PROMEDIO en su posición.
- Los nombres de estudiantes del Consolidado deben ser fórmulas hacia `S1`.
- Las filas TOTAL/PROMEDIO deben contener fórmulas (`=SUM…`, `=AVERAGE…`).

---

## 📋 Menús y Herramientas

Una vez instalado, aparecen dos menús:

**📝 ZipGrade** — envía los resultados de la hoja de sesión seleccionada (S1, S2, …) hacia el Consolidado (`B1`).

**⚙️ Configuración y Estructura**
- **⚡ Recalcular y Corregir Fórmulas** — reconstruye TOTAL/PROMEDIO y reaplica todo el formato (incluida la fuente Nunito).
- **🧹 Limpiar Filas Vacías** — elimina filas sobrantes al final de todas las hojas.
- **🏆 Resaltar Mejor Promedio** — resalta al estudiante con mejor promedio.
- **🔢 Ordenar por Total / Restaurar Alfabético** — alterna el orden por TOTAL o alfabético en el Consolidado.
- **📚 Gestión de Asignaturas** / **📂 Gestión de Sesiones** — gestores HTML con arrastrar-y-soltar.
- **🔍 Reporte de Inconsistencias de Estudiantes** — detecta faltantes y no registrados.
- **📥 Exportar Configuración JSON** / **📤 Importar Configuración JSON** — respaldo y restauración de la estructura.

---

## 🔢 Flujo de Uso (Paso a Paso)

1. **Prepara la lista maestra**: procesa el primer examen de ZipGrade en **S1** (menú 📝 ZipGrade → `S1`). Esto fija la lista definitiva de estudiantes.
2. **Procesa las demás sesiones**: pega el CSV en `Sheet1` y ejecuta 📝 ZipGrade → `S2`, `S3`, … Se alinean automáticamente con S1; los faltantes reciben nota `0` y se te notifican las inconsistencias.
3. **Estructura las asignaturas**: usa 📚 Gestión de Asignaturas para definir nombres, cantidad de preguntas y orden por sesión.
4. **Recalcula**: ejecuta ⚡ Recalcular y Corregir Fórmulas para consolidar TOTAL/PROMEDIO, formato y nombres dinámicos.
5. **Analiza**: ordena por TOTAL, resalta el mejor promedio, exporta JSON para respaldar.

---

## 🎨 Esquema de Colores (Consolidado)

- **Cabeceras**: fila 1 = bloques `SESIÓN`, filas 2–3 = asignaturas (fondo `#527FD9`).
- **Notas por semaforización** (formato condicional automático):
  - `0.5 – 3.49` → 🔴 Rojo (`#E06666`)
  - `3.5 – 3.99` → 🟠 Naranja (`#E69238`)
  - `4.0 – 4.49` → 🟡 Amarillo (`#F1C232`)
  - `4.5 – 5.0` → 🟢 Verde (`#93C47D`)
- **Filas de estudiantes**: fondo blanco/gris claro alternado (cebra).

---

## 📥 Esquema de Configuración JSON

Puedes respaldar o importar la configuración de asignaturas/sesiones con JSON. El importador crea hojas automáticamente y **preserva las notas ya existentes** de las columnas:

```json
[
  {
    "name": "MATEMÁTICAS",
    "count": 10,
    "session": "1"
  },
  {
    "name": "FÍSICA",
    "count": 15,
    "session": "1"
  },
  {
    "name": "QUÍMICA",
    "count": 8,
    "session": "2"
  }
]
```

---

## 🛠️ Instalación y Configuración (Como Biblioteca)

El proyecto está estructurado como **Biblioteca de Google Apps Script** para actualizar fácilmente múltiples libros de cálculo sin copiar 2800+ líneas de código.

### Paso 1: Desplegar la Biblioteca
1. Crea un proyecto nuevo de Google Apps Script.
2. Coloca `Code.js` y `appsscript.json` en el editor.
3. Despliega como biblioteca: **Implementar (Deploy)** > **Nueva implementación** > tipo **Biblioteca**.
4. Copia el **ID de script** del despliegue.

### Paso 2: Añadir la Biblioteca al Libro Destino
1. Abre el libro de Google Sheets de destino.
2. Abre su editor de Apps Script (**Extensiones** > **Apps Script**).
3. En el panel izquierdo, clic en `+` junto a **Bibliotecas**.
4. Pega el **ID de script**, clic en **Buscar**, elige la última versión, define el Identificador como `ZipGradeControllerMain` y clic en **Añadir**.

### Paso 3: Añadir el Código Puente del Cliente
1. Copia el contenido de [`client_code.js`](./client_code.js).
2. Pégalo en el `Code.gs` del libro de destino.
3. Guarda y recarga el libro: aparecerá automáticamente el menú `📝 ZipGrade Controller`.

---

## 💻 Desarrollo Local (con clasp)

Puedes escribir y modificar el código en tu computadora y sincronizarlo a Google Apps Script usando la utilidad oficial `clasp`.

### Requisitos
Node.js/NPM instalados.

### Comandos
1. **Activa la API de Apps Script**: https://script.google.com/home/usersettings
2. **Inicia sesión en Google**:
   ```bash
   npx @google/clasp login
   ```
3. **Vincula tu Script**:
   ```bash
   npx @google/clasp clone "TU_ID_DE_SCRIPT"
   ```
4. **Sube los cambios locales a Google**:
   ```bash
   npx @google/clasp push
   ```
5. **Descarga los cambios de la nube a local**:
   ```bash
   npx @google/clasp pull
   ```
