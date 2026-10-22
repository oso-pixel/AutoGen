# Demo AutoGen — Análisis de ventas de una cafetería

Sistema multi-agente con **AutoGen 0.4+** (`autogen-agentchat`) que analiza las
ventas de una microempresa de café y deja evidencias en disco. Funciona con la
API compatible con OpenAI de **DeepSeek** o con **OpenAI**, según el `.env`.

## Equipo de agentes

| Agente | Rol |
|--------|-----|
| Analista | Escribe el código Python con pandas. **NO ejecuta código.** |
| Ejecutor | `CodeExecutorAgent` que recibe el código del Analista, lo ejecuta y devuelve la salida real. |
| Revisor | Valida las cifras del Ejecutor. Si todo está bien, resume y escribe `TERMINADO`. |

El flujo es **Analista → Ejecutor → Revisor**, con `RoundRobinGroupChat` y
`TextMentionTermination("TERMINADO") | MaxMessageTermination(20)`.
El Ejecutor usa `sources=["Analista"]`, así que solo ejecuta código escrito por
el Analista.

## Requisitos

- Python 3.10 o superior (en este equipo: Python 3.14, se invoca con `py`).
- Dependencias del archivo `requirements.txt`:
  ```powershell
  py -3.14 -m pip install -r requirements.txt
  ```

## Configuración

El proyecto no incluye claves. Copia el archivo de ejemplo y coloca la tuya:

```powershell
Copy-Item .env.example .env
```

En Linux/macOS: `cp .env.example .env`. Luego edita `.env`.

El script prioriza `OPENAI_API_KEY` cuando tiene una clave real; si sigue en el
placeholder, usa `DEEPSEEK_API_KEY`.

Opción OpenAI:
```env
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4o-mini
```

Opción DeepSeek (la que usa esta demo):
```env
DEEPSEEK_API_KEY=sk-...
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-flash
```

## Ejecutar

```powershell
py -3.14 demo_autogen.py
```

Al ejecutarlo, el script te pide las instrucciones para el equipo en una sola
línea: escribes la tarea y presionas Enter. Si presionas Enter directo, usa la
tarea de ejemplo. También puedes pasarlas como argumento:

```powershell
py -3.14 demo_autogen.py "Analiza las ventas y genera el reporte en evidencias/"
```

Solo se imprime el **último mensaje del Revisor**. El resto de la conversación
(código del Analista, salida del Ejecutor, etc.) no aparece en pantalla, pero
queda guardada en `evidencias/*.txt`.

En este equipo `python` no está en el PATH, por eso usamos `py -3.14`. Si tu
entorno sí tiene `python`, puedes ejecutar `python demo_autogen.py`.
Asegúrate de estar dentro de la carpeta `demo_autogen_cafe`.

> `demo_cafe.py` es la versión anterior, sin generación de evidencias.

## Evidencias

El script crea la carpeta `evidencias/` según lo que pidas:

- **Un solo mes**: genera únicamente `ventas_<mes>.xlsx` (por ejemplo,
  `ventas_mayo.xlsx`). No genera los archivos por producto, por mes ni de otros
  meses.
- **General o anual**: genera los tres archivos estándar, un Excel por mes y
  `ventas_anual.xlsx`.

En ambos casos guarda la conversación en los tres `.txt`.

Ejemplos:

- `"Quiero las ventas de mayo"` → solo `ventas_mayo.xlsx`.
- `"Solo quiero este mes"` → pregunta de qué mes y genera solo ese archivo.
- `"Genera el Excel de todo el año"` → los tres estándar, los seis mensuales y
  `ventas_anual.xlsx`.

| Archivo | Contenido |
|---------|-----------|
| `Analista.txt` | Código y mensajes del Analista. |
| `Ejecutor.txt` | Salida real de la ejecución. |
| `Revisor.txt` | Validación, resumen ejecutivo y recomendaciones. |
| `ventas_detallado.xlsx` | Las 18 filas con `ingresos_totales`. |
| `ventas_por_producto.xlsx` | Totales por producto, ordenados de mayor a menor por ingresos. |
| `ventas_por_mes.xlsx` | Totales mensuales en orden cronológico. |
| `ventas_enero.xlsx` … `ventas_junio.xlsx` | Un Excel por mes con sus ventas. |
| `ventas_anual.xlsx` | Consolidado del período con hojas `Detalle`, `Por producto` y `Por mes`. |

Los datos actuales cubren **enero–junio**; `ventas_anual.xlsx` consolida esos
seis meses. Si agregas más meses a la tabla, la misma lógica genera sus
archivos.

## Contexto compartido

Los agentes reciben el contenido de `contexto_negocio.md` en cada tarea:
empresa, productos, período, resultados de referencia y convención de archivos.
Puedes editar ese archivo para actualizar el "recorrido" que conocen los
agentes.

## Notas para la exposición

- El modelo se configura con `OpenAIChatCompletionClient` apuntando a
  `https://api.deepseek.com/v1`, porque DeepSeek es compatible con la API de OpenAI.
- Se pasa `model_info` manualmente: AutoGen no conoce los modelos de DeepSeek en
  su catálogo interno.
- DeepSeek resuelve el alias `deepseek-chat` como `deepseek-flash`. Usamos
  `deepseek-flash` directamente para evitar el warning de AutoGen y tener una
  estimación de tokens correcta.
- DeepSeek activa **thinking mode** por defecto y devuelve su razonamiento
  interno junto a la respuesta. El cliente envía
  `extra_body={"thinking": {"type": "disabled"}}` para desactivarlo y que los
  agentes muestren solo su respuesta final.
- `demo_autogen.py` usa `team.run(...)` y muestra únicamente el último mensaje
  del Revisor. La conversación completa (Analista, Ejecutor y Revisor) queda en
  `evidencias/*.txt`.
- El Ejecutor es un `CodeExecutorAgent` con
  `LocalCommandLineCodeExecutor(work_dir=".")`, para que `evidencias/` quede en
  la raíz del proyecto. El Analista no tiene herramientas de ejecución: solo
  produce el bloque de código.
- El código del Analista incluye `os.makedirs('evidencias')` y los tres
  `to_excel(...)`; por eso se requiere `openpyxl`.
- El Revisor no ejecuta nada: valida la salida real que publica el Ejecutor y,
  si las cifras son correctas, cierra con un resumen ejecutivo y `TERMINADO`.
  Si el usuario pide un mes concreto, responde con ese mes; si pide el año,
  muestra el consolidado.
- El script genera automáticamente los Excel mensuales y `ventas_anual.xlsx` a
  partir de `ventas_detallado.xlsx`, sin imprimir nada extra en la terminal.
- El código del Analista imprime un bloque de `CONTROL` con
  `CUADRE OK: True/False`. El Revisor usa ese cuadre y la salida del Ejecutor
  como evidencia; no recalcula sumas largas mentalmente, para evitar errores
  aritméticos del modelo.
- En Windows se fuerzan `PYTHONIOENCODING=utf-8` y `PYTHONUTF8=1` para que el
  subproceso del Ejecutor emita UTF-8. Así se evita el `UnicodeDecodeError`
  que aparecía con acentos como "Café" cuando el hijo usaba cp1252.
- El ejecutor local corre código generado por el modelo en esta máquina. Es
  adecuado para una demo controlada; para producción se recomienda
  `DockerCommandLineCodeExecutor`.
- Si quieres usar `deepseek-reasoner`, cambia `DEEPSEEK_MODEL` en el `.env`.
- **No compartas el archivo `.env`**: contiene tu API key. El `.gitignore` ya
  lo excluye de Git.
