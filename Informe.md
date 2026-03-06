# Informe del Proyecto: ONPE Consulta Electoral Masiva v2.0

## Descripción General
Aplicación de escritorio para consulta masiva de DNIs en el portal de la ONPE 
(https://consultaelectoral.onpe.gob.pe/inicio). Permite cargar una lista de DNIs 
desde archivos CSV, TXT o Excel, y consulta automáticamente cada uno para obtener 
datos electorales como local de votación, mesa de sufragio, si es miembro de mesa, etc.

---

## Stack Tecnológico

### Backend / Lógica
| Tecnología | Versión | Uso |
|---|---|---|
| **Python** | 3.13.2 | Lenguaje principal |
| **undetected-chromedriver** | 3.5.5 | Automatización de Chrome evitando detección de bots/reCAPTCHA |
| **Selenium** | 4.41.0 | WebDriver para interacción con elementos del DOM |
| **SQLite** | (built-in) | Base de datos local para persistencia de resultados |
| **openpyxl** | 3.1.5 | Lectura de archivos Excel (.xlsx) |

### Frontend / Interfaz Gráfica
| Tecnología | Uso |
|---|---|
| **Tkinter** (built-in) | Interfaz gráfica de escritorio |
| **ttk** (themed widgets) | Widgets con estilo visual mejorado (tema `clam`) |

### Infraestructura
| Componente | Descripción |
|---|---|
| **Google Chrome** | Navegador real controlado por undetected-chromedriver (v145) |
| **Perfil Chrome persistente** | Carpeta `chrome_profile/` con cookies/sesión entre ejecuciones |
| **Git** | Control de versiones local |

---

## Arquitectura del Proyecto

```
OPERACIONES/
├── onpe_consulta.py      # Código principal (875 líneas)
│   ├── Database           # Clase de acceso a datos (SQLite)
│   ├── ONPEScraper        # Clase de scraping con Chrome
│   └── App                # Interfaz gráfica Tkinter
├── onpe_consultas.db      # Base de datos SQLite
├── chrome_profile/        # Perfil de Chrome persistente
├── carga.xlsx             # Archivo de ejemplo con DNIs
├── requirements.txt       # Dependencias Python
├── instalar.bat           # Script de instalación
├── ejecutar.bat           # Script de ejecución
├── .gitignore             # Archivos excluidos de Git
└── CLAUDE.md              # Instrucciones de desarrollo
```

### Clases Principales

1. **`Database`** — Manejo de SQLite con operaciones CRUD:
   - `upsert()`: Insertar/actualizar registros
   - `pending_dnis()`: Obtener DNIs ya consultados
   - `export_csv()`: Exportar resultados a CSV
   - `stats()`: Estadísticas de consultas

2. **`ONPEScraper`** — Web scraping con anti-detección:
   - Usa `undetected-chromedriver` para evadir reCAPTCHA v3
   - Simula comportamiento humano (movimientos de mouse, typing natural)
   - Extrae datos del DOM con regex y CSS selectors
   - Reintentos automáticos ante errores del servidor

3. **`App`** — Interfaz gráfica completa:
   - Carga masiva de DNIs desde archivos
   - Progreso en tiempo real
   - Tabla de resultados con colores por estado
   - Log de actividad en tiempo real
   - Exportación a CSV

---

## Correcciones Realizadas (2026-03-06)

### Bug 1: Detección de versión de Chrome incompleta ⚠️ CRÍTICO
- **Problema**: Solo buscaba la versión en `HKLM` del registro de Windows, pero en muchas instalaciones Chrome está registrado en `HKCU`.
- **Impacto**: Usaba un valor hardcodeado (145) que dejará de funcionar cuando Chrome se actualice.
- **Solución**: Búsqueda secuencial en HKLM, HKCU y WOW6432Node. Además, si la versión detectada falla al crear el driver, reintenta sin forzar versión (auto-detect).

### Bug 2: Thread Safety con SQLite ⚠️ MODERADO
- **Problema**: Las conexiones SQLite se creaban sin `check_same_thread=False`, pero las operaciones de BD se ejecutan tanto desde el hilo principal (GUI) como desde el hilo worker.
- **Impacto**: En Python 3.x, SQLite por defecto lanza `ProgrammingError` si se accede desde un hilo diferente al que creó la conexión.
- **Solución**: Añadido `check_same_thread=False` en todas las conexiones.

### Bug 3: Borrado destructivo de BD al cargar archivo ⚠️ MODERADO
- **Problema**: Cada vez que se cargaba un archivo nuevo, se ejecutaba `self.db.clear()` automáticamente sin preguntar, borrando todos los resultados previos.
- **Impacto**: Pérdida accidental de datos al cargar un nuevo archivo.
- **Solución**: Se muestra un diálogo de confirmación preguntando si desea limpiar la BD o mantener los registros existentes.

### Bug 4: Progreso inexacto
- **Problema**: La barra de progreso calculaba `i / total * 100`, empezando en 0% para el primer DNI.
- **Solución**: Cambiado a `(i + 1) / total * 100` para reflejar correctamente el progreso.

### Bug 5: Logs de errores sin traceback
- **Problema**: Los errores críticos solo mostraban el mensaje de error, sin el traceback completo.
- **Solución**: Se añadió `traceback.format_exc()` para facilitar la depuración.

---

## Cómo Ejecutar

1. **Instalar dependencias** (primera vez):
   ```
   ejecutar instalar.bat
   ```

2. **Ejecutar la aplicación**:
   ```
   ejecutar ejecutar.bat
   ```

3. **O directamente**:
   ```
   .venv\Scripts\python.exe onpe_consulta.py
   ```
