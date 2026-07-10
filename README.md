# 🧾 Sistema Automatizado de Generación de Recibos de Pago

Script en Python que automatiza la generación masiva de recibos de pago mensuales para empleados, a partir de una **plantilla Word (.docx)** y un **archivo XML** con los datos de cada trabajador.

Por cada empleado se genera **un único archivo `.docx`**, con un recibo por mes dentro del rango solicitado, cada uno en su propia página — conservando exactamente el formato, tablas, encabezados, pies de página, márgenes y estilos de la plantilla original.

```
Salida/
├── Juan Pérez López.docx        # Página 1: Enero · Página 2: Febrero · Página 3: Marzo ...
└── María Sofía Gómez.docx       # Un recibo por página, uno por cada mes del rango
```

---

## ✨ Características

- 📄 **Un archivo por empleado** con todos sus recibos del rango de meses, cada uno en una página independiente (salto de página automático).
- 🎨 **Formato 100% fiel a la plantilla**: tablas, imágenes, encabezados, pies de página, márgenes y estilos se preservan sin alteraciones.
- 🔎 **Reemplazo inteligente de texto**: sustituye marcadores de la plantilla (nombre, DPI, salario, fechas, etc.) sin romper el formato original, incluso si Word fragmentó el texto en múltiples "runs".
- 📊 **Lectura automática desde XML**: carga todos los empleados y sus datos (nombre, DPI, fecha de inicio, salario) desde un solo archivo.
- 🗓️ **Cálculo automático de fechas**: último día del mes, nombres de mes en español, formatos de fecha institucionales (`DD-MES-AAAA`).
- 💰 **Formato de moneda local** (Quetzales de Guatemala): `Q 2,200.00`.
- 📈 **Barra de progreso en consola** en tiempo real.
- ✅ **Validaciones de entrada**: rutas inexistentes, rangos de mes inválidos, XML vacío o corrupto, etc.

---

## 🧰 Requisitos

- Python 3.9 o superior
- Dependencias:
  - [`python-docx`](https://python-docx.readthedocs.io/) — lectura/escritura de archivos Word
  - [`docxcompose`](https://github.com/4teamwork/docxcompose) — fusión de los recibos mensuales en un único documento por empleado, preservando el formato

### Instalación

```bash
pip install python-docx docxcompose
```

---

## 📁 Formato del archivo XML de entrada

El script espera un XML con un nodo `<empleado>` por cada trabajador:

```xml
<empleados>
  <empleado>
    <nombre>Juan Pérez López</nombre>
    <dpi>1234 56789 0101</dpi>
    <fecha_inicio>2024-05-06</fecha_inicio>
    <salario>3500.50</salario>
  </empleado>
  <empleado>
    <nombre>María Sofía Gómez</nombre>
    <dpi>9876 54321 0202</dpi>
    <fecha_inicio>2023-01-15</fecha_inicio>
    <salario>4200</salario>
  </empleado>
</empleados>
```

| Campo           | Formato          | Descripción                              |
|-----------------|-------------------|-------------------------------------------|
| `nombre`        | texto             | Nombre completo del empleado              |
| `dpi`           | texto             | Documento Personal de Identificación      |
| `fecha_inicio`  | `AAAA-MM-DD`      | Fecha de inicio laboral                   |
| `salario`       | número decimal    | Salario mensual                           |

---

## 📝 Plantilla Word

La plantilla `.docx` debe contener los siguientes marcadores/textos de ejemplo, que el script localiza y reemplaza automáticamente por los datos reales de cada empleado y mes:

- `NOMBRE DEL TRABAJADOR:,CINTHYA JEANNETH GARCÍA MARROQUÍN`
- `3737 93944 0101` (DPI)
- `06-MAYO-2024` (fecha de inicio)
- `Del 1 al 30 de Abril 2026` (período de pago)
- `1 al 30 de abril de 2026` (período de percepción)
- `Q 2,200.00` (monto del salario)
- `30-4-2026` (fecha de recepción)
- `NOMBRE:,` y `FIRMA:,`
- Párrafo completo de recibido de pago

> El resto de la plantilla (logotipos, tablas, márgenes, tipografía) se conserva intacto: el script solo sustituye texto, nunca reconstruye el documento desde cero.

---

## ▶️ Uso

Ejecuta el script y responde las preguntas en consola:

```bash
python main.py
```

```
==================================================
  SISTEMA AUTOMATIZADO DE GENERACIÓN DE RECIBOS
==================================================

1. Ruta de la plantilla Word (.docx): plantilla.docx
2. Ruta del archivo XML: empleados.xml
3. Carpeta donde guardar los documentos: Salida
4. Año (ej. 2026): 2026
5. Mes inicial (1-12): 1
6. Mes final (1-12): 3

[Proceso] Analizando el archivo XML de origen...
[Éxito] Se cargaron correctamente 2 empleados.
[Proceso] Se generarán 6 recibos en total.

Progreso: |████████████████████████████████████████| 100.0% (6/6)

==================================================
 ¡Proceso completado de manera exitosa!
 Se generó 1 archivo .docx por empleado (2 en total),
 cada uno con 3 recibo(s), uno por página.
 Destino: /ruta/completa/Salida
==================================================
```

### Datos solicitados

| # | Dato                        | Ejemplo         |
|---|------------------------------|-----------------|
| 1 | Ruta de la plantilla Word    | `plantilla.docx`|
| 2 | Ruta del archivo XML         | `empleados.xml` |
| 3 | Carpeta de salida             | `Salida`        |
| 4 | Año                            | `2026`          |
| 5 | Mes inicial (1-12)            | `1`             |
| 6 | Mes final (1-12)              | `3`             |

---

## ⚙️ Cómo funciona internamente

1. **Lectura del XML** (`leer_empleados_xml`): parsea todos los nodos `<empleado>` y normaliza sus datos, incluyendo el formato de fecha institucional.
2. **Por cada empleado**, para cada mes del rango:
   - Se calcula el mapa de reemplazos específico del mes (`construir_mapa_reemplazos`): último día del mes, período de pago, monto formateado, etc.
   - Se genera un recibo individual temporal a partir de la plantilla, aplicando los reemplazos sobre párrafos y celdas de tablas (`procesar_documento_word` / `reemplazar_texto_en_parrafo`).
3. **Fusión de meses** (`combinar_recibos_mensuales`): los recibos mensuales temporales del empleado se combinan, en orden, en un único documento usando `docxcompose`, insertando un salto de página duro antes de cada mes (excepto el primero), sin tocar encabezados, pies de página ni estilos.
4. Los archivos temporales se descartan automáticamente; solo el `.docx` consolidado final queda en la carpeta de salida.
5. La barra de progreso se actualiza por cada recibo mensual procesado.

---

## 📂 Estructura del proyecto

```
.
├── main.py          # Script principal (lectura XML, generación y fusión de recibos)
└── README.md
```

---

## ⚠️ Notas y limitaciones

- Los nombres de archivo se derivan del nombre del empleado; los caracteres `/` y `\` se reemplazan por `_` para evitar rutas inválidas.
- Si el rango de meses es inválido (mes inicial mayor al final, o fuera de 1-12), el proceso se detiene con un mensaje de error antes de generar nada.
- Si el XML no contiene empleados o no puede leerse, el proceso se detiene con un mensaje explicativo.
- El script no valida el contenido de la plantilla; si los marcadores de texto no coinciden exactamente con los definidos en `construir_mapa_reemplazos`, esos campos simplemente no se reemplazan en el recibo.

---

## 📄 Licencia

Este proyecto puede distribuirse y adaptarse libremente según las necesidades de la institución que lo utilice.
