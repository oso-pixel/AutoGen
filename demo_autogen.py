"""
demo_autogen.py
===============
Demo de AutoGen 0.4+ (autogen-agentchat) para una exposición.

Flujo con RoundRobinGroupChat:
    1. Analista -> escribe el código Python con pandas. NUNCA lo ejecuta.
    2. Ejecutor -> CodeExecutorAgent que corre el código y devuelve la salida real.
    3. Revisor  -> valida las cifras y cierra con un resumen ejecutivo y TERMINADO.

Además de la conversación en la terminal, el script deja evidencias:
    evidencias/Analista.txt
    evidencias/Ejecutor.txt
    evidencias/Revisor.txt

    Si la solicitud es de un solo mes:
        evidencias/ventas_<mes>.xlsx

    Si la solicitud es general o anual:
        evidencias/ventas_detallado.xlsx
        evidencias/ventas_por_producto.xlsx
        evidencias/ventas_por_mes.xlsx
        evidencias/ventas_<mes>.xlsx (uno por mes)
        evidencias/ventas_anual.xlsx

CÓMO SE RESOLVIÓ (para que todo funcione)
-----------------------------------------
1. El Analista NO ejecuta código: solo escribe el bloque. El Ejecutor
   (CodeExecutorAgent) lo recibe, lo corre y publica la salida real.
2. El Revisor valida la salida del Ejecutor y el bloque "CUADRE OK" que imprime
   el script. No hace sumas largas mentalmente: una versión anterior se
   equivocó con una suma y bloqueó la aprobación.
3. En Windows se fuerza UTF-8 en el subproceso del Ejecutor. Si el hijo usa
   cp1252, los acentos de "Café" rompen la decodificación de AutoGen.
4. DeepSeek se usa con el cliente OpenAI-compatible; se desactiva su "thinking
   mode" para que no muestre el razonamiento interno.
5. La solicitud se clasifica como "un solo mes" o "general/anual". En un solo
   mes se genera únicamente ventas_<mes>.xlsx; en general se genera todo.
6. Antes de ejecutar se borran los Excel de solicitudes anteriores para que la
   carpeta refleje solo lo que se pidió ahora.
7. La terminal muestra únicamente el último mensaje del Revisor; el resto de
   la conversación queda en evidencias/*.txt.
"""

# asyncio: motor async de AutoGen.
# os/sys: rutas, carpeta de evidencias y argumentos por línea de comandos.
# re/unicodedata: detectar el mes aunque haya acentos o mayúsculas.
# warnings: silenciar avisos esperados de esta demo controlada.
import asyncio
import os
import re
import sys
import unicodedata
import warnings

# pandas: leer y escribir los Excel de evidencia.
import pandas as pd
# dotenv: cargar OPENAI_API_KEY / DEEPSEEK_API_KEY desde .env.
from dotenv import load_dotenv

# --- Componentes de AutoGen 0.4+ -------------------------------------------
from autogen_agentchat.agents import AssistantAgent, CodeExecutorAgent
from autogen_agentchat.conditions import MaxMessageTermination, TextMentionTermination
from autogen_agentchat.teams import RoundRobinGroupChat
from autogen_core.models import ModelFamily
from autogen_ext.code_executors.local import LocalCommandLineCodeExecutor
from autogen_ext.models.openai import OpenAIChatCompletionClient

# Carga las variables de entorno del archivo .env de la raíz del proyecto.
# Lo normal es ejecutar el script desde demo_autogen_cafe/.
load_dotenv()

# ---------------------------------------------------------------------------
# CORRECCIÓN PARA WINDOWS: UTF-8 EN EL SUBPROCESO DEL EJECUTOR
# ---------------------------------------------------------------------------
# LocalCommandLineCodeExecutor lanza el código del Analista con sys.executable
# y captura stdout/stderr como bytes. Después AutoGen hace stdout.decode()
# asumiendo UTF-8. En Windows, un subproceso Python puede escribir con cp1252;
# entonces "Café" produce el byte 0xE9, que no es UTF-8 válido, y la demo falla
# con UnicodeDecodeError.
# El ejecutor copia os.environ para el hijo, así que estas variables se
# heredan y obligan al subproceso a emitir UTF-8.
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

# ---------------------------------------------------------------------------
# AVISOS QUE NO APORTAN EN ESTA DEMO
# ---------------------------------------------------------------------------
# 1) El ejecutor local avisa que corre código en la máquina. Es intencional:
#    es una demo local con código generado por el propio equipo.
warnings.filterwarnings(
    "ignore",
    message="Using LocalCommandLineCodeExecutor may execute code on the local machine.*",
    category=UserWarning,
)
# 2) CodeExecutorAgent avisa que no hay aprobación humana. La demo aprueba
#    automáticamente para poder correr sin intervención.
warnings.filterwarnings(
    "ignore",
    message="No approval function set for CodeExecutorAgent.*",
    category=UserWarning,
)
# 3) work_dir="." está deprecado, pero lo usamos a propósito para que las rutas
#    "evidencias/..." queden en la raíz del proyecto.
warnings.filterwarnings(
    "ignore",
    message="Using the current directory as work_dir is deprecated.",
    category=DeprecationWarning,
)

# ---------------------------------------------------------------------------
# DATOS DE ENTRADA (no se modifican).
# Esta tabla es la única fuente de verdad: se incluye en la tarea para que el
# Analista la copie en su código. Además, datos_desde_tabla() la vuelve a leer
# como respaldo para garantizar los Excel aunque el modelo cometa un error.
# ---------------------------------------------------------------------------
TABLA_VENTAS = """
| mes    | producto          | unidades_vendidas | precio_unitario_cop |
|--------|-------------------|-------------------|---------------------|
| Enero  | Café Molido 500 g | 120               | 18000               |
| Enero  | Café en Grano 1kg | 80                | 42000               |
| Enero  | Café Instantáneo  | 200               | 12000               |
| Febrero| Café Molido 500 g | 135               | 18000               |
| Febrero| Café en Grano 1kg | 85                | 42000               |
| Febrero| Café Instantáneo  | 190               | 12000               |
| Marzo  | Café Molido 500 g | 150               | 18000               |
| Marzo  | Café en Grano 1kg | 90                | 42000               |
| Marzo  | Café Instantáneo  | 210               | 12000               |
| Abril  | Café Molido 500 g | 145               | 18000               |
| Abril  | Café en Grano 1kg | 95                | 42000               |
| Abril  | Café Instantáneo  | 220               | 12000               |
| Mayo   | Café Molido 500 g | 160               | 18000               |
| Mayo   | Café en Grano 1kg | 100               | 42000               |
| Mayo   | Café Instantáneo  | 230               | 12000               |
| Junio  | Café Molido 500 g | 175               | 18000               |
| Junio  | Café en Grano 1kg | 110               | 42000               |
| Junio  | Café Instantáneo  | 240               | 12000               |
"""

# Instrucciones por defecto si el usuario no escribe ninguna.
TAREA_POR_DEFECTO = (
    "Analiza las ventas mensuales de la microempresa. "
    "Quiero el reporte de ventas ordenado de mayor a menor. "
    "Genera los archivos de evidencia en la carpeta evidencias/."
)

# MESES interpreta el lenguaje natural del usuario ("mayo", "Mayo", etc.).
# MESES_DISPONIBLES limita la demo a los meses que existen en la tabla.
# ORDEN_MESES fija el orden cronológico al escribir los archivos mensuales y
# la hoja "Por mes" del consolidado anual.
MESES = {
    "enero": "Enero",
    "febrero": "Febrero",
    "marzo": "Marzo",
    "abril": "Abril",
    "mayo": "Mayo",
    "junio": "Junio",
    "julio": "Julio",
    "agosto": "Agosto",
    "septiembre": "Septiembre",
    "setiembre": "Septiembre",
    "octubre": "Octubre",
    "noviembre": "Noviembre",
    "diciembre": "Diciembre",
}
MESES_DISPONIBLES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"]
ORDEN_MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# ---------------------------------------------------------------------------
# MENSAJES DE SISTEMA
# ---------------------------------------------------------------------------
# El prompt del Analista es clave:
# - Le prohíbe ejecutar: la ejecución es responsabilidad del Ejecutor.
# - Le exige crear la carpeta y guardar el Excel según el modo solicitado.
# - Le pide el bloque CUADRE OK para que el Revisor valide sin hacer sumas.
# - Le prohíbe inventar "esperados" externos que no estén en la tabla.
ANALISTA_PROMPT = """
Escribes código Python con pandas. NUNCA ejecutes código tú mismo. Solo devuelve el bloque de código.

El código que escribas DEBE incluir al inicio:
    import os
    os.makedirs('evidencias', exist_ok=True)

Al final guarda los Excel según el MODO DE LA SOLICITUD que aparece en la tarea:

- Modo "UN SOLO MES (<Mes>)":
    df_mes = df[df['mes'] == '<Mes>']
    df_mes.to_excel('evidencias/ventas_<mes>.xlsx', index=False)
  NO generes ventas_detallado.xlsx, ventas_por_producto.xlsx,
  ventas_por_mes.xlsx ni archivos de otros meses.

- Modo "GENERAL / ANUAL":
    df.to_excel('evidencias/ventas_detallado.xlsx', index=False)
    por_producto.sort_values('ingresos_totales', ascending=False).to_excel('evidencias/ventas_por_producto.xlsx', index=False)
    por_mes.to_excel('evidencias/ventas_por_mes.xlsx', index=False)
  Ordena SIEMPRE los resultados de MAYOR a MENOR por ingresos antes de guardar.

Contexto: eres el Analista de una microempresa de café.
- Ten en cuenta el CONTEXTO DEL NEGOCIO que aparece en la tarea (productos,
  período, archivos y resultados de referencia).
- Adapta el análisis a las instrucciones del usuario: si pide un mes concreto,
  imprime el detalle de ese mes; si pide el año o todas las ventas, imprime el
  consolidado.
- El sistema se encarga de los archivos adicionales según el modo: en modo de
  un solo mes solo debe existir ventas_<mes>.xlsx; en modo general se generan
  además los archivos mensuales y ventas_anual.xlsx.
- Embebe los 18 registros de la tabla en el script (no leas archivos externos
  ni uses internet).
- Crea `df` con las columnas: mes, producto, unidades_vendidas,
  precio_unitario_cop e ingresos_totales.
- Crea `por_producto` con al menos las columnas producto, unidades_totales e
  ingresos_totales.
- Crea `por_mes` con las columnas mes e ingresos_totales, en orden cronológico.
- Imprime: ingresos por producto, ingresos totales por mes, producto más
  vendido, mes con mayores ingresos y variación porcentual enero-junio.
- Imprime un bloque de CONTROL con el total de unidades, el total de ingresos y
  la comprobación de que la suma por producto sea igual a la suma por mes y al
  total general. Imprime `CUADRE OK: True` o `CUADRE OK: False`.
- Imprime el nombre del archivo Excel que generaste.
- No inventes "valores esperados" externos: todo se calcula con las 18 filas.
- Incluye comentarios en español y usa print() para mostrar los resultados.
- Después del bloque de código, explica brevemente qué hace el script.
- Si el Revisor pide correcciones, entrega el código completo corregido.
- No cierres la conversación; eso le corresponde al Revisor.
- No muestres tu razonamiento interno ni notas de proceso; responde directo.
"""

# El prompt del Revisor corrige un problema real de la versión anterior: el
# modelo intentaba recalcular sumas mentales y se equivocaba (por ejemplo,
# 80+85+90+95+100+110 = 560, no 660), bloqueando la aprobación. Ahora:
# - La salida real del Ejecutor es la fuente de verdad.
# - Valida el bloque CUADRE OK que calcula Python.
# - No recalcula ni inventa valores esperados.
# - Responde según el modo y cierra con TERMINADO.
REVISOR_PROMPT = """
Validas las cifras. Si todo está bien, escribes un resumen ejecutivo con total de
ingresos, producto con mayores ingresos, mes con mayores ingresos, variación
enero-junio y 2 recomendaciones. Terminas con la palabra TERMINADO.

Responde según el MODO DE LA SOLICITUD que aparece en la tarea:
- Modo "UN SOLO MES (<Mes>)": valida y responde únicamente sobre ese mes.
  Menciona `evidencias/ventas_<mes>.xlsx` y NO menciones
  ventas_detallado.xlsx, ventas_por_producto.xlsx, ventas_por_mes.xlsx, los
  archivos mensuales de otros meses ni ventas_anual.xlsx. No incluyas el total
  del período ni cifras de otros meses; usa el control solo para validar.
- Modo "GENERAL / ANUAL": incluye el resumen ejecutivo completo (total,
  producto con mayores ingresos, mes con mayores ingresos, variación y 2
  recomendaciones) y menciona los archivos estándar, los mensuales y
  `evidencias/ventas_anual.xlsx`.

Reglas de validación:
- La salida real del Ejecutor es la fuente de verdad.
- NO rehagas sumas largas mentalmente y NO inventes "valores esperados" que no
  estén en la tabla. El Ejecutor ya hizo los cálculos.
- No declares que el entorno está corrupto si el Ejecutor no muestra un traceback.

Verifica:
1. Que la ejecución no haya producido errores.
2. Que el código use las 18 filas de la tabla y las columnas correctas.
3. Que las fórmulas sean correctas: ingresos_totales = unidades_vendidas *
   precio_unitario_cop, y variación = ((junio - enero) / enero) * 100.
4. Que la salida incluya el bloque de CONTROL y que `CUADRE OK` sea `True`.
5. Que se haya generado el archivo Excel que corresponde al modo:
   - un solo mes: `ventas_<mes>.xlsx`;
   - general/anual: los tres archivos estándar.
6. Que aparezcan los cinco resultados pedidos: ingresos por producto, ingresos
   por mes, producto más vendido, mes con mayores ingresos y variación
   enero-junio.

Si todos los puntos se cumplen, APRUEBA usando las cifras de la salida; no
recalculles las sumas por tu cuenta.
Si falta algo, pide una corrección concreta y NO apruebes.

Al aprobar: escribe "APROBADO", el resumen ejecutivo con las cifras de la salida
y 2 recomendaciones para la microempresa, y TERMINADO en la última línea, solo,
sin comillas.
No reescribas el código completo. No muestres tu razonamiento interno ni notas
de proceso; responde directo.
"""


# El "recorrido" es un archivo editable por el usuario. Se inyecta en la tarea
# para que Analista y Revisor compartan el mismo contexto del negocio.
def cargar_contexto_negocio() -> str:
    """Carga el 'recorrido' del negocio desde contexto_negocio.md."""
    ruta = "contexto_negocio.md"
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            return f.read().strip()
    return (
        "Microempresa de café con tres productos. Ventas mensuales de enero a "
        "junio. Los reportes se guardan en la carpeta evidencias/."
    )


# Los usuarios escriben "MAYO", "mayo", "maYo"... Normalizamos para comparar.
def normalizar_texto(texto: str) -> str:
    """Pasa a minúsculas y quita acentos para detectar meses con tolerancia."""
    texto = texto.lower()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def detectar_mes_solicitado(texto: str) -> str | None:
    """
    Devuelve el mes si la solicitud menciona exactamente uno.
    Devuelve None si es general/anual o si menciona varios meses.
    """
    t = normalizar_texto(texto)

    # Si piden el año completo, todos los meses, el semestre, etc., NO es una
    # solicitud de un solo mes: se trata como general/anual.
    claves_general = [
        "anual", "todos los meses", "semestre", "completo", "general",
        "todo el periodo",
    ]
    if any(clave in t for clave in claves_general) or re.search(r"\bano\b", t):
        return None

    # Buscamos nombres de mes como palabras completas para no confundir, por
    # ejemplo, "mayo" con una palabra que lo contenga.
    encontrados = [
        mes for clave, mes in MESES.items()
        if re.search(rf"\b{re.escape(clave)}\b", t)
    ]
    encontrados = list(dict.fromkeys(encontrados))

    # Exactamente un mes -> modo de un solo mes.
    # Cero meses o varios meses -> solicitud general/anual.
    if len(encontrados) == 1:
        return encontrados[0]
    return None


def resolver_mes_solicitado(instrucciones: str) -> str | None:
    """
    Determina si la solicitud es de un solo mes. Si el usuario dice
    "este mes" sin nombrarlo, le pregunta cuál.
    """
    # Primero intentamos detectar un mes nombrado en la instrucción.
    mes = detectar_mes_solicitado(instrucciones)
    if mes:
        # Validamos que ese mes exista en los datos (enero-junio).
        if mes not in MESES_DISPONIBLES:
            print(
                f"No hay datos de {mes}. "
                f"Meses disponibles: {', '.join(MESES_DISPONIBLES)}."
            )
            return None
        return mes

    # Si no hay mes nombrado, pero la frase sugiere uno ("este mes",
    # "solo un mes", etc.), preguntamos cuál para no generar todos.
    t = normalizar_texto(instrucciones)
    frases_mes = [
        "este mes", "solo un mes", "un solo mes", "mes actual", "solo este mes",
    ]
    if not any(frase in t for frase in frases_mes):
        return None

    # Repetimos hasta obtener un mes válido o hasta que el usuario cancele
    # con Enter (en ese caso se usa el modo general).
    while True:
        print(f"¿De qué mes quieres el reporte? ({', '.join(MESES_DISPONIBLES)})")
        try:
            respuesta = input("> ").strip()
        except EOFError:
            return None
        if not respuesta:
            return None
        mes = detectar_mes_solicitado(respuesta)
        if mes in MESES_DISPONIBLES:
            return mes
        if mes:
            print(
                f"No hay datos de {mes}. "
                f"Meses disponibles: {', '.join(MESES_DISPONIBLES)}."
            )
        else:
            print(
                "No reconocí el mes. Escribe el nombre del mes o presiona "
                "Enter para el reporte general."
            )


def obtener_instrucciones() -> str:
    """
    Obtiene el texto de las instrucciones del usuario.

    - Si se pasan como argumentos de línea de comandos, se usan directamente.
    - Si no, se piden por teclado en una sola línea.
    - Si el usuario presiona Enter sin escribir nada, se usa TAREA_POR_DEFECTO.
    """
    # Prioridad 1: instrucciones pasadas como argumentos. Sirve para demos
    # rápidas o pruebas sin interacción:
    #   py demo_autogen.py "Quiero las ventas de mayo"
    if len(sys.argv) > 1:
        instrucciones = " ".join(sys.argv[1:]).strip()
    else:
        # Prioridad 2: pedirlas por teclado en una sola línea.
        print("Escribe las instrucciones para el equipo.")
        print("Presiona Enter directo para usar la tarea de ejemplo.")
        try:
            instrucciones = input("> ").strip()
        except EOFError:
            instrucciones = ""

    # Si no escribió nada, usamos la tarea de ejemplo.
    if not instrucciones:
        instrucciones = TAREA_POR_DEFECTO
    return instrucciones


def construir_tarea(instrucciones: str, mes: str | None) -> str:
    """
    Construye la tarea con el modo, el contexto y la tabla de ventas.

    El bloque "MODO DE LA SOLICITUD" es el contrato que leen los prompts del
    Analista y del Revisor para saber qué archivos generar y qué responder.
    """
    contexto = cargar_contexto_negocio()
    if mes:
        # Modo estricto: solo el Excel del mes pedido. Así no aparecen los
        # archivos por producto, por mes ni de otros meses.
        modo = (
            f"UN SOLO MES ({mes}). Genera únicamente "
            f"evidencias/ventas_{mes.lower()}.xlsx con las filas de {mes}. "
            "NO generes ventas_detallado.xlsx, ventas_por_producto.xlsx, "
            "ventas_por_mes.xlsx ni archivos de otros meses."
        )
    else:
        # Modo completo: los tres estándar y, al final, los mensuales y el anual.
        modo = (
            "GENERAL / ANUAL. Genera los tres archivos estándar "
            "(ventas_detallado.xlsx, ventas_por_producto.xlsx, "
            "ventas_por_mes.xlsx). El sistema generará además los archivos "
            "mensuales y ventas_anual.xlsx."
        )
    # La tarea final contiene lo que pidió el usuario, el modo, el contexto
    # compartido y los datos que el Analista debe usar.
    return (
        f"{instrucciones}\n\n"
        f"MODO DE LA SOLICITUD:\n{modo}\n\n"
        f"CONTEXTO DEL NEGOCIO:\n{contexto}\n\n"
        f"Datos de ventas:\n{TABLA_VENTAS}"
    )


def crear_modelo():
    """
    Crea el cliente del modelo.

    Prioridad:
      1. OPENAI_API_KEY (si tiene una clave real).
      2. DEEPSEEK_API_KEY (configuración que ya funcionaba en el proyecto).
    """
    # Leemos las dos claves posibles. El .env puede tener una o ambas.
    openai_key = os.getenv("OPENAI_API_KEY", "").strip()
    deepseek_key = os.getenv("DEEPSEEK_API_KEY", "").strip()

    # Prioridad 1: si hay una clave real de OpenAI, usamos OpenAI normal.
    if openai_key and openai_key != "tu_api_key_aqui":
        return OpenAIChatCompletionClient(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            api_key=openai_key,
        )

    # Prioridad 2: DeepSeek. Usamos el mismo cliente OpenAI porque DeepSeek
    # expone una API compatible; solo cambia la URL base.
    if deepseek_key and deepseek_key != "tu_api_key_aqui":
        return OpenAIChatCompletionClient(
            model=os.getenv("DEEPSEEK_MODEL", "deepseek-flash"),
            api_key=deepseek_key,
            # base_url redirige el protocolo OpenAI hacia DeepSeek.
            base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1"),
            # DeepSeek activa "thinking mode" por defecto y devuelve su cadena
            # de razonamiento junto a la respuesta. Lo desactivamos para que los
            # agentes muestren solo su respuesta final.
            extra_body={"thinking": {"type": "disabled"}},
            # AutoGen no conoce "deepseek-flash" en su catálogo interno, así que
            # declaramos sus capacidades manualmente.
            model_info={
                "vision": False,
                "function_calling": True,
                "json_output": True,
                "structured_output": False,
                "family": ModelFamily.UNKNOWN,
            },
        )

    # Sin claves: main() mostrará el mensaje de configuración.
    return None


def datos_desde_tabla() -> pd.DataFrame:
    """
    Construye el DataFrame desde la tabla fija (fuente de respaldo).

    Sirve para garantizar los Excel aunque el código generado por el LLM falle
    o no cree el detalle. La tabla es la misma que ve el Analista en la tarea.
    """
    filas = []
    # Recorremos la tabla markdown línea por línea.
    for linea in TABLA_VENTAS.strip().splitlines():
        linea = linea.strip()
        if not linea.startswith("|"):
            continue
        celdas = [c.strip() for c in linea.strip("|").split("|")]

        # Ignorar el encabezado ("mes") y la línea separadora ("|---|---|").
        if len(celdas) < 4 or celdas[0].lower() == "mes":
            continue
        if all(set(c) <= set("-: ") for c in celdas):
            continue

        filas.append(
            {
                "mes": celdas[0],
                "producto": celdas[1],
                "unidades_vendidas": int(celdas[2]),
                "precio_unitario_cop": int(celdas[3]),
            }
        )

    df = pd.DataFrame(filas)
    # La tabla no trae ingresos: se calculan aquí.
    df["ingresos_totales"] = df["unidades_vendidas"] * df["precio_unitario_cop"]
    return df


def limpiar_excels_evidencias(permitido: str | None = None) -> None:
    """
    Borra los Excel generados en evidencias/ para que la carpeta refleje
    solo la solicitud actual. Si se indica `permitido`, conserva ese archivo.
    """
    carpeta = "evidencias"
    if not os.path.isdir(carpeta):
        return
    # Solo borramos los Excel "ventas_*.xlsx"; los .txt de la conversación
    # se conservan (se van acumulando entre ejecuciones).
    for archivo in os.listdir(carpeta):
        if archivo.startswith("ventas_") and archivo.endswith(".xlsx"):
            # `permitido` se usa en modo de un solo mes para conservar el
            # archivo del mes solicitado y borrar cualquier extra.
            if permitido and archivo == permitido:
                continue
            os.remove(os.path.join(carpeta, archivo))


def generar_reportes(mes: str | None) -> None:
    """
    Genera los Excel según la solicitud:
    - Un solo mes: únicamente ventas_<mes>.xlsx.
    - General/anual: un Excel por mes y ventas_anual.xlsx.
    """
    # Aseguramos que exista la carpeta antes de escribir cualquier archivo.
    os.makedirs("evidencias", exist_ok=True)

    # -----------------------------------------------------------------------
    # MODO DE UN SOLO MES
    # Genera únicamente ventas_<mes>.xlsx. No crea los archivos estándar,
    # por producto, por mes, de otros meses ni el anual.
    # -----------------------------------------------------------------------
    if mes:
        df = datos_desde_tabla()
        grupo = df[df["mes"] == mes]
        grupo.to_excel(f"evidencias/ventas_{mes.lower()}.xlsx", index=False)
        return

    # -----------------------------------------------------------------------
    # MODO GENERAL / ANUAL
    # Preferimos el detalle real que dejó el Ejecutor (ventas_detallado.xlsx).
    # Si no existe o no tiene las columnas esperadas, usamos la tabla fija.
    # -----------------------------------------------------------------------
    ruta_detalle = os.path.join("evidencias", "ventas_detallado.xlsx")
    columnas = {"mes", "producto", "unidades_vendidas", "precio_unitario_cop"}
    if os.path.exists(ruta_detalle):
        df = pd.read_excel(ruta_detalle)
        if not columnas.issubset(df.columns):
            df = datos_desde_tabla()
    else:
        df = datos_desde_tabla()

    if "ingresos_totales" not in df.columns:
        df["ingresos_totales"] = df["unidades_vendidas"] * df["precio_unitario_cop"]

    # Un Excel por cada mes presente en los datos (enero-junio en esta demo).
    for nombre_mes in ORDEN_MESES:
        grupo = df[df["mes"] == nombre_mes]
        if not grupo.empty:
            grupo.to_excel(f"evidencias/ventas_{nombre_mes.lower()}.xlsx", index=False)

    # Consolidado anual con tres hojas:
    #   Detalle      -> las 18 filas
    #   Por producto -> totales ordenados de mayor a menor
    #   Por mes      -> totales en orden cronológico
    por_producto = (
        df.groupby("producto", as_index=False)
        .agg(
            unidades_totales=("unidades_vendidas", "sum"),
            ingresos_totales=("ingresos_totales", "sum"),
        )
        .sort_values("ingresos_totales", ascending=False)
    )
    por_mes = df.groupby("mes", as_index=False)["ingresos_totales"].sum()
    meses_presentes = [m for m in ORDEN_MESES if m in set(df["mes"])]
    # Categorical + ordered fuerza el orden cronológico; si no, pandas podría
    # ordenar los meses alfabéticamente.
    por_mes["mes"] = pd.Categorical(
        por_mes["mes"], categories=meses_presentes, ordered=True
    )
    por_mes = por_mes.sort_values("mes")

    with pd.ExcelWriter("evidencias/ventas_anual.xlsx") as writer:
        # Hoja 1: detalle completo (las 18 filas).
        df.to_excel(writer, sheet_name="Detalle", index=False)
        # Hoja 2: totales por producto, de mayor a menor ingreso.
        por_producto.to_excel(writer, sheet_name="Por producto", index=False)
        # Hoja 3: totales mensuales en orden cronológico.
        por_mes.to_excel(writer, sheet_name="Por mes", index=False)


# ---------------------------------------------------------------------------
# GUARDRAIL DURO (determinista, sin LLM)
# ---------------------------------------------------------------------------
# El Revisor es un modelo opinando; esta validación es código puro y por eso
# no se puede "convencer". Revisa los Excel que quedaron en evidencias/.
def validar_evidencias(mes: str | None) -> list[str]:
    """Devuelve la lista de problemas encontrados (vacía = todo correcto)."""
    problemas: list[str] = []
    carpeta = "evidencias"
    referencia = datos_desde_tabla()

    if mes:
        ruta = os.path.join(carpeta, f"ventas_{mes.lower()}.xlsx")
        if not os.path.exists(ruta):
            return [f"no existe {ruta}"]
        df = pd.read_excel(ruta)
        filas = referencia[referencia["mes"] == mes]
        if len(df) != len(filas):
            problemas.append(f"{ruta}: {len(df)} filas, se esperaban {len(filas)}")
        elif int(df["ingresos_totales"].sum()) != int(filas["ingresos_totales"].sum()):
            problemas.append(f"{ruta}: los ingresos no cuadran")
        return problemas

    requeridos = [
        "ventas_detallado.xlsx", "ventas_por_producto.xlsx",
        "ventas_por_mes.xlsx", "ventas_anual.xlsx",
    ]
    for nombre in requeridos:
        if not os.path.exists(os.path.join(carpeta, nombre)):
            problemas.append(f"falta {nombre}")
    if problemas:
        return problemas

    detalle = pd.read_excel(os.path.join(carpeta, "ventas_detallado.xlsx"))
    if len(detalle) != len(referencia):
        problemas.append(
            f"ventas_detallado.xlsx: {len(detalle)} filas, "
            f"se esperaban {len(referencia)}"
        )
    elif int(detalle["ingresos_totales"].sum()) != int(
        referencia["ingresos_totales"].sum()
    ):
        problemas.append("ventas_detallado.xlsx: el total no cuadra con la tabla")

    anual = pd.read_excel(
        os.path.join(carpeta, "ventas_anual.xlsx"), sheet_name=None
    )
    for hoja in ("Detalle", "Por producto", "Por mes"):
        if hoja not in anual:
            problemas.append(f"ventas_anual.xlsx: falta la hoja '{hoja}'")
    return problemas


def sabotear_evidencia(mes: str | None) -> None:
    """Solo para la demo: borra la última fila de un Excel para probar el guardrail."""
    nombre = f"ventas_{mes.lower()}.xlsx" if mes else "ventas_detallado.xlsx"
    ruta = os.path.join("evidencias", nombre)
    df = pd.read_excel(ruta)
    df.drop(index=df.index[-1]).to_excel(ruta, index=False)


async def main() -> None:
    # 1) Crear el cliente del modelo (OpenAI o DeepSeek, según el .env).
    model_client = crear_modelo()
    if model_client is None:
        print("⚠️  No hay API key configurada.")
        print("   Agrega OPENAI_API_KEY (o DEEPSEEK_API_KEY) en el archivo .env.")
        return

    # 2) Leer lo que pide el usuario y decidir el modo.
    #    - Un mes nombrado ("mayo") -> modo de un solo mes.
    #    - "este mes" sin nombre -> se pregunta cuál.
    #    - "año", "completo", "todos los meses", etc. -> modo general.
    instrucciones = obtener_instrucciones()
    mes_solicitado = resolver_mes_solicitado(instrucciones)
    tarea = construir_tarea(instrucciones, mes_solicitado)

    # 3) Limpiar Excel de ejecuciones anteriores para que la carpeta refleje
    #    solo la solicitud actual. Los .txt de la conversación se conservan.
    os.makedirs("evidencias", exist_ok=True)
    limpiar_excels_evidencias()

    ejecutor_local = None
    try:
        # 4) Crear los tres agentes del flujo.
        #    Analista: solo escribe el bloque de código. Sin herramientas de
        #    ejecución, para garantizar que la ejecución sea del Ejecutor.
        analista = AssistantAgent(
            name="Analista",
            model_client=model_client,
            system_message=ANALISTA_PROMPT,
        )

        #    Ejecutor: corre el código con el directorio actual como work_dir,
        #    para que las rutas "evidencias/..." queden en la raíz del proyecto.
        ejecutor_local = LocalCommandLineCodeExecutor(work_dir=".")
        ejecutor = CodeExecutorAgent(
            name="Ejecutor",
            code_executor=ejecutor_local,
            # Solo ejecuta bloques de código escritos por el Analista.
            sources=["Analista"],
        )

        #    Revisor: valida la salida real y cierra con TERMINADO.
        revisor = AssistantAgent(
            name="Revisor",
            model_client=model_client,
            system_message=REVISOR_PROMPT,
        )

        # 5) Orquestación: turnos fijos Analista -> Ejecutor -> Revisor.
        #    TextMentionTermination corta cuando el Revisor escribe TERMINADO;
        #    MaxMessageTermination(20) evita un bucle infinito si algo falla.
        termination = (
            TextMentionTermination("TERMINADO")
            | MaxMessageTermination(max_messages=20)
        )
        team = RoundRobinGroupChat(
            [analista, ejecutor, revisor],
            termination_condition=termination,
        )

        # 6) Ejecutar el equipo. team.run devuelve TaskResult con todos los
        #    mensajes; no usamos Console ni imprimimos la conversación.
        result = await team.run(task=tarea)

        # 7) Generar los Excel según el modo.
        #    En modo de un solo mes borramos cualquier Excel extra que haya
        #    creado el Analista y conservamos solo ventas_<mes>.xlsx.
        if mes_solicitado:
            nombre_mes = f"ventas_{mes_solicitado.lower()}.xlsx"
            limpiar_excels_evidencias(permitido=nombre_mes)
        generar_reportes(mes_solicitado)

        # 8) Imprimir ÚNICAMENTE el último mensaje del Revisor.
        #    El resto de la conversación (código del Analista, salida del
        #    Ejecutor, etc.) no se muestra en pantalla.
        mensajes_revisor = [msg for msg in result.messages if msg.source == "Revisor"]
        if mensajes_revisor:
            ultimo_mensaje = mensajes_revisor[-1]
            contenido = ultimo_mensaje.content
            print(contenido if isinstance(contenido, str) else str(contenido))
        else:
            print("⚠️ El Revisor no emitió ningún mensaje.")

        # 8b) GUARDRAIL DURO sobre las evidencias. Con DEMO_FALLO=1 se
        #     corrompe un Excel a propósito para que el guardrail falle.
        if os.getenv("DEMO_FALLO") == "1":
            sabotear_evidencia(mes_solicitado)

        problemas = validar_evidencias(mes_solicitado)
        if problemas:
            print("\n❌ GUARDRAIL FALLIDO:")
            for problema in problemas:
                print(f"   - {problema}")
        else:
            print("\n✅ GUARDRAIL PASADO: evidencias verificadas.")

        # -------------------------------------------------------------------
        # 9) BLOQUE POST-EJECUCIÓN: guardar la conversación en .txt.
        # -------------------------------------------------------------------
        # Crear carpeta de evidencias (idempotente).
        os.makedirs("evidencias", exist_ok=True)

        # Guardar lo que dijo cada agente en su propio archivo. Se abre en
        # modo "a" (append) para no perder ejecuciones anteriores.
        for msg in result.messages:
            nombre = msg.source  # "Analista", "Ejecutor", "Revisor" o "user"
            if nombre in ["Analista", "Ejecutor", "Revisor"]:
                archivo = f"evidencias/{nombre}.txt"
                with open(archivo, "a", encoding="utf-8") as f:
                    f.write(f"===== {nombre} =====\n")
                    f.write(str(msg.content) + "\n\n")
    finally:
        # 10) Cerrar el cliente y detener el ejecutor aunque haya un error.
        #     model_client.close() evita warnings de conexión abierta.
        await model_client.close()
        if ejecutor_local is not None:
            await ejecutor_local.stop()


# Punto de entrada: asyncio.run inicia el flujo async de AutoGen.
if __name__ == "__main__":
    asyncio.run(main())
