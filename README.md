# powerbi_bot.py — Bot de análisis académico Power BI + Banner

Automatiza la extracción de datos académicos desde Power BI (Playwright) y los cruza con SVAPROY en Banner INB para determinar qué cursos de un alumno pueden proyectarse en el siguiente período.

---

## Flujo general

```
CSV de IDs
    │
    ▼
[GENERAL]  →  selecciona ID en slicer → extrae nombre y código de programa
    │                                         └─ guarda en resultados_powerbi.xlsx
    ▼
[DETALLE DEL ÁREA]  →  scrollea tabla virtualizada → detecta cursos con valor "N"
    │
    ▼
[CURSOS PRINCIPALES]  →  filtra por programa académico → extrae mapa curso→prerrequisitos
    │                   →  evalúa lógica AND/OR de prerreqs → clasifica cada curso N:
    │                         · "Para proyectar"  → prerreqs cumplidos
    │                         · "No proyectar"    → falta al menos un prerreq
    │                         · "Revisar"         → expresión de prerreqs no evaluable
    │                   └─ guarda en analisis.xlsx
    │                   └─ cursos "Revisar" → para_revisar.xlsx
    ▼
[SVAPROY (Banner)]  →  rellena formulario (ID + periodo + plan de estudios)
    │               →  exporta vía SHIFT+F1 o lectura DOM
    │                    └─ cruza "Para proyectar" vs SVAPROY por código (Materia+Curso)
    │                         → analisis_svaproy.xlsx
    │               →  retorna lista de cursos NO encontrados en SVAPROY
    ▼
[Reporte de NRCs u Horario]  →  verifica PERIODO una vez (sin re-clic si ya está marcado)
                             →  por cada curso sin SVAPROY:
                                    limpia MATERIA con borrador ◇ → selecciona materia
                                    limpia NÚMERO DE CURSO con borrador ◇ → selecciona número
                                    si tabla tiene datos → NRC existe → ERROR DEL SISTEMA
                                    fail-fast: para en el primer NRC encontrado
                                        └─ guarda en analisis_total.xlsx (ID | ERROR_DEL_SISTEMA)
```

---

## Archivos de entrada y salida

| Archivo | Rol |
|---|---|
| `FR_70140526.csv` | Input: una columna con IDs de alumno (sin encabezado de datos, con header en fila 1) |
| `resultados_powerbi.xlsx` | Output: fila por alumno con nombre y datos del reporte GENERAL |
| `analisis.xlsx` | Output: por cada curso con N → código, prerrequisitos y estado `Para proyectar` / `No proyectar` / `Revisar` |
| `analisis_svaproy.xlsx` | Output: cursos "Para proyectar" cruzados con SVAPROY **por código** (`Materia+Curso`), no por descripción |
| `analisis_total.xlsx` | Output: una fila por alumno — `Sí` si al menos un curso sin SVAPROY tiene NRC disponible (error del sistema), `No` en caso contrario |
| `para_revisar.xlsx` | Output: cursos cuya expresión de prerrequisitos no pudo evaluarse (ej. formato `&` en vez de `AND`) — para revisión manual |
| `log_powerbi_YYYYMMDD_HHMMSS.txt` | Output: log detallado de ejecución |

---

## Configuración (parte superior del script)

```python
POWER_BI_URL    # URL del reporte Power BI (app + report + page)
INPUT_FILE      # CSV de IDs de entrada
WAIT_CARGA      # Segundos de espera tras seleccionar un ID en GENERAL/DETALLE/CURSOS (default: 14)
WAIT_NRC        # Segundos de espera tras filtrar MATERIA+NÚMERO en la sección NRC (default: 10)
TEST_MODE       # True → procesa solo 1 ID y no consume el CSV completo
TEST_IDS        # IDs forzados en TEST_MODE, ej. ["72986227"]. [] = toma el primer ID del CSV
BANNER_PERIODO  # Período a proyectar en SVAPROY y en slicer NRC, ej. "202610"
BANNER_URL      # URL del applicationNavigator de Banner INB
```

---

## Dependencias

```
playwright  (sync_api)
openpyxl
```

Instalar con:
```bash
pip install playwright openpyxl
playwright install chrome
```

---

## Ejecución

```bash
python powerbi_bot.py
```

1. Se abre Chrome con Power BI.
2. El script pausa y pide login manual → presionar **ENTER** al terminar.
3. Se abre Banner en segunda pestaña y navega a SVAPROY automáticamente.
4. El bot itera cada ID del CSV (o `TEST_IDS` si `TEST_MODE=True`).
5. Al finalizar, pausa nuevamente → presionar **ENTER** para cerrar el navegador.

---

## Módulos y funciones clave

### Slicers (Power BI)
| Función | Qué hace |
|---|---|
| `abrir_slicer(page, aria_label)` | Abre el desplegable del slicer correcto por `aria-label` (NFC + lowercase) |
| `seleccionar_en_slicer(page, aria_label, valor)` | Abre, busca y hace clic en un valor (sin limpiar antes) |
| `verificar_slicer_seleccionado(...)` | Verifica si el valor ya está marcado; solo hace clic si no lo está |
| `seleccionar_programa_unico(...)` | Limpia con doble-clic en "Seleccionar todo" y selecciona un solo valor |
| `limpiar_slicer_borrador(page, aria_label)` | Limpia el slicer con el botón ◇ "Borrar selecciones"; fallback a "Seleccionar todo" si el botón no está visible |

### Extracción Power BI
| Función | Qué hace |
|---|---|
| `extraer_datos_general(page, id)` | Extrae headers + fila de la tabla GENERAL via JS |
| `obtener_programa_code(data)` | Parsea el código de programa (ej. `PR08MEDH`) desde los datos de GENERAL |
| `extraer_cursos_N_con_scroll(page)` | Scrollea la tabla virtualizada de DETALLE DEL ÁREA y acumula cursos con valor `"N"` |
| `extraer_mapa_prereqs(page)` | Scrollea CURSOS PRINCIPALES y construye `{nombre_curso: prereqs}` + `{codigo: nombre}` |

### Lógica de proyección
| Función | Qué hace |
|---|---|
| `calcular_proyectar(prereq_str, mapa_codigos, cursos_n_set)` | Evalúa expresión `AND`/`OR` de códigos; devuelve `"Para proyectar"` / `"No proyectar"` / `"Revisar"` |
| `es_curso_excluido(nombre)` | Excluye electivos, minors y cursos de libre configuración |

### Banner / SVAPROY
| Función | Qué hace |
|---|---|
| `navegar_svaproy(banner_page)` | Usa el buscador lateral de Banner para abrir SVAPROY |
| `rellenar_svaproy(banner_page, id, periodo)` | Llena el formulario: ID → Tab → Periodo → F9 Plan de estudios → primer plan → OK → Ir |
| `exportar_svaproy(banner_page)` | Intenta SHIFT+F1 (descarga), luego popup, luego lectura DOM |
| `_parsear_export_svaproy(path)` | Parsea XLSX/CSV exportado de Banner; maneja fila-bloque-clave + encabezados opcionales |
| `guardar_analisis_svaproy(id, cursos_bot, filas_svaproy)` | Cruza por código (`Materia+Curso` normalizado); retorna lista de cursos **no** encontrados en SVAPROY |

### NRC
| Función | Qué hace |
|---|---|
| `buscar_nrc_cursos(page, cursos_sin_svaproy)` | Navega a "Reporte de NRCs u Horario", verifica PERIODO una vez (sin re-clic si ya está), limpia y filtra por cada curso; fail-fast |
| `verificar_nrc_existe(page)` | Evalúa `JS_TIENE_NRC`: si `.bodyCells` tiene celdas con texto, retorna `True` |
| `guardar_analisis_total(id, error_sistema)` | Graba `ID \| Sí/No` en `analisis_total.xlsx` |
| `guardar_para_revisar(filas)` | Graba cursos con prerreqs no evaluables en `para_revisar.xlsx` |

---

## Notas de diseño

- **Virtualización**: Power BI no renderiza todas las filas a la vez. Las funciones de extracción usan scroll programático con `JS_FIND_SCROLL_CONTAINER` para recorrer la tabla completa.
- **Unicode NFC**: los `aria-label` de Power BI pueden tener variantes de normalización. Todo el JS de búsqueda normaliza a NFC + lowercase antes de comparar (importante para `NÚMERO DE CURSO` con tilde).
- **Matching por código**: SVAPROY y Power BI usan descripciones diferentes para el mismo curso. El cruce se hace por código (`Materia+Curso` sin espacios, ej. `INVE01010`), derivado del campo `Main_Table.Curso` en CURSOS PRINCIPALES.
- **Borrador ◇**: `limpiar_slicer_borrador` dispara un `mouseover` sobre el visual container para forzar la aparición del botón, luego lo busca con `[title="Borrar selecciones"]` subiendo hasta 12 niveles en el DOM. El log indica si usó el botón o el fallback.
- **PERIODO en NRC**: se verifica una sola vez al entrar a la sección. Si ya está en `202610`, no hace clic y espera solo 2s (no los 10s de `WAIT_NRC`).
- **Fail-fast NRC**: en cuanto detecta un NRC para un curso, para y marca el alumno como error del sistema sin revisar el resto.
- **Lógica prerrequisitos**: `"BIET08001 [11] AND TEOL08002 [11]"` se transforma a `True/False` con `eval()`. Si la expresión tiene formato inesperado (ej. `&` en vez de `AND`), `eval` falla y el curso va a `para_revisar.xlsx`.
- **Banner INB**: el formulario SVAPROY vive en un `iframe`. La función `_svaproy_frame` lo detecta por presencia de `≥2 inputs de texto`. Los eventos de teclado críticos (Tab, F9, SHIFT+F1) se envían desde `banner_page` para garantizar que lleguen al navegador.

---

## Consideraciones al correr

- **Borrar `.xlsx` antes de la primera corrida**: los archivos generados por versiones anteriores no tienen la columna "Código" en `analisis.xlsx`. Si el bot hace append, el header queda inconsistente. Renombrar o borrar los `.xlsx` antes de usar esta versión.
- **Cursos sin código**: si un curso de DETALLE DEL ÁREA no aparece en CURSOS PRINCIPALES (electivo de otro plan), `codigo` queda vacío → no matchea en SVAPROY → va a la lista de candidatos NRC pero se omite por código inválido. Aparece en el log como `"código inválido, omitido"`.
- **Modo test activo**: `TEST_MODE = True` con `TEST_IDS = ["72986227"]`. Cambiar a `TEST_MODE = False` para producción, o `TEST_IDS = []` para tomar el primer ID del CSV.

---

## BUG PENDIENTE — Sección NRC no encuentra slicers (analizar mañana)

### Síntoma
```
Toggle 3 abrió: [None, None, None, ...] (buscaba: 'PERIODO')
...
Toggle 13 abrió: [None, None, None, ...] (buscaba: 'PERIODO')
No se pudo abrir slicer 'PERIODO'
No se pudo confirmar PERIODO 202610
[1] NRC revisados: 13 cursos   ← se logueó el conteo pero no chequeó ninguno
analisis_total.xlsx: 72986227 → OK   ← marcó OK en vez de verificar
```

### Diagnóstico
- Hay **13 `.slicerBody` en el DOM** pero todos devuelven `aria-label = null`.
- La sección GENERAL/DETALLE/CURSOS sí funciona — el problema es exclusivo de "Reporte de NRCs u Horario".
- El HTML que se inspeccionó a mano **sí muestra** `aria-label="PERIODO"` en `.slicerBody`, así que el atributo existe.

### Causas probables (en orden de probabilidad)
1. **Timing**: la sección NRC tarda más en renderizar. `navegar_seccion` espera 6s fijos, puede que los `aria-label` se asignen después de ese tiempo.
2. **Frame/contexto distinto**: si la página NRC carga en un sub-frame, `document.querySelectorAll` desde el frame equivocado devuelve elementos sin atributos.
3. **Slicer visual diferente**: la sección NRC podría usar una versión distinta del visual de slicer de Power BI donde `aria-label` se asigna en un elemento padre distinto al `.slicerBody`.

### Fixes a probar mañana
- **Fix 1 (timing)**: en `buscar_nrc_cursos`, antes de llamar a `verificar_slicer_seleccionado`, agregar:
  ```python
  page.wait_for_selector('.slicerBody[aria-label]', timeout=20000, state='attached')
  ```
  Esto espera explícitamente a que al menos un `.slicerBody` tenga `aria-label` asignado.
- **Fix 2 (fallback por posición)**: si el aria-label sigue sin aparecer, identificar PERIODO por ser el 1.er toggle de la sección en vez de por nombre.
- **Verificar en prueba**: imprimir `page.url` y el HTML completo de los primeros `.slicerBody` justo al entrar a la sección, para confirmar qué contiene el DOM realmente.

---

## Troubleshooting rápido

| Síntoma | Revisión |
|---|---|
| Slicer no se abre | Verificar que el `aria-label` exacto (con tildes/mayúsculas) coincida con el DOM — revisar log |
| `"Borrador no encontrado"` en log | El botón ◇ requiere hover; el fallback "Seleccionar todo" se activó automáticamente, no es error crítico |
| `"NÚMERO DE CURSO"` no selecciona | Verificar si el slicer tiene tilde (`Ú`) o no en el Power BI real; ajustar la cadena en `buscar_nrc_cursos` |
| Tabla GENERAL vacía | Aumentar `WAIT_CARGA`; la página puede tardar más en cargar |
| Frame SVAPROY no encontrado | Revisar que Banner no requiera login adicional; el frame tarda hasta 15 s |
| SHIFT+F1 no descarga | Banner puede no tener extract habilitado; el bot cae a lectura DOM automáticamente |
| Archivo `.xlsx` abierto en Excel | El bot genera un archivo alternativo con timestamp en el nombre |
| Cursos en `para_revisar.xlsx` | Los prerrequisitos tienen formato no estándar (ej. `&` en vez de `AND`); revisar manualmente y ajustar `calcular_proyectar` si el patrón se repite |
