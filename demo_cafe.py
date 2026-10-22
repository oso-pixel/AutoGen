"""
demo_cafe.py
============
Demo de AutoGen 0.4+ (autogen-agentchat) para una exposición.

Arquitectura con RoundRobinGroupChat:
    1. Analista -> escribe el código Python con pandas. NO lo ejecuta.
    2. Ejecutor -> CodeExecutorAgent que recibe el código del Analista,
                   lo ejecuta y devuelve la salida real.
    3. Revisor  -> valida las cifras del Ejecutor y, si todo está bien,
                   responde con un resumen ejecutivo y la palabra TERMINADO.

El modelo se consume a través de la API compatible con OpenAI de DeepSeek.
"""

import asyncio
import os
import warnings
from pathlib import Path

from dotenv import load_dotenv

# --- Componentes de AutoGen 0.4+ -------------------------------------------
from autogen_agentchat.agents import AssistantAgent, CodeExecutorAgent
from autogen_agentchat.base import TaskResult
from autogen_agentchat.conditions import MaxMessageTermination, TextMentionTermination
from autogen_agentchat.teams import RoundRobinGroupChat
from autogen_core.models import ModelFamily
from autogen_ext.code_executors.local import LocalCommandLineCodeExecutor
from autogen_ext.models.openai import OpenAIChatCompletionClient

# Carga el .env que está junto a este script, sin depender del directorio actual.
load_dotenv(Path(__file__).with_name(".env"))

# El Ejecutor corre código generado por el modelo en esta máquina. Es adecuado
# para una demo controlada; silenciamos únicamente su aviso de seguridad.
warnings.filterwarnings(
    "ignore",
    message="Using LocalCommandLineCodeExecutor may execute code on the local machine.*",
    category=UserWarning,
)
# En esta demo el Ejecutor aprueba automáticamente el código del Analista.
warnings.filterwarnings(
    "ignore",
    message="No approval function set for CodeExecutorAgent.*",
    category=UserWarning,
)

# Carpeta de trabajo donde el Ejecutor corre los scripts del Analista.
CARPETA_EJECUCION = Path(__file__).with_name("sandbox")

# ---------------------------------------------------------------------------
# Datos de ejemplo: microempresa de café (precios en pesos colombianos).
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

# Tarea inicial para el equipo. No contiene la palabra de terminación.
TAREA_INICIAL = f"""
Somos el equipo de análisis de una microempresa de café. La dueña del negocio
necesita un informe con cifras reales para tomar decisiones.

Analiza esta tabla de ventas mensuales (unidades vendidas y precio unitario en COP):

{TABLA_VENTAS}

Objetivos del análisis:
1. Calcular los ingresos por producto (unidades_vendidas * precio_unitario_cop).
2. Calcular el total de ingresos por mes.
3. Identificar el producto más vendido y el mes con mayores ingresos.
4. Calcular la variación porcentual de ingresos entre enero y junio.

Flujo de trabajo:
- Analista: escribe el código completo con pandas; NO lo ejecutes.
- Ejecutor: ejecuta el código del Analista y devuelve la salida real.
- Revisor: valida las cifras y cierra con un resumen ejecutivo y recomendaciones.
"""

# ---------------------------------------------------------------------------
# Mensajes de sistema: cada agente tiene un rol y límites claros.
# ---------------------------------------------------------------------------
ANALISTA_PROMPT = """
Escribes código Python con pandas. NO ejecutes código. Solo devuelve el bloque de código.

Contexto: eres el Analista de una microempresa de café.
- Entrega SIEMPRE un único bloque ```python ... ``` autocontenido.
- Embebe los datos de la tabla en el script (no leas archivos externos ni uses
  internet).
- Crea un DataFrame con las columnas mes, producto, unidades_vendidas y
  precio_unitario_cop.
- Calcula e imprime: ingresos por producto, ingresos totales por mes, producto
  más vendido, mes con mayores ingresos y variación porcentual enero-junio.
- Incluye un bloque de CONTROL que imprima el total de unidades, el total de
  ingresos y compruebe que la suma de ingresos por producto sea igual a la suma
  por mes y al total general. Imprime una línea `CUADRE OK: True` o
  `CUADRE OK: False` según el resultado.
- No inventes "valores esperados" externos: todos los controles deben calcularse
  a partir de las 18 filas de la tabla.
- Incluye comentarios en español y usa print() para mostrar cada resultado.
- Después del bloque de código, explica brevemente qué hace el script.
- Si el Revisor pide correcciones, entrega el código completo corregido.
- No cierres la conversación; eso le corresponde al Revisor.
- No muestres tu razonamiento interno ni notas de proceso; responde directo.
"""

REVISOR_PROMPT = """
Validas que las cifras del Ejecutor sean correctas. Si todo está bien, respondes
con un resumen ejecutivo y la palabra TERMINADO.

Reglas de validación:
- La salida real del Ejecutor es la fuente de verdad.
- NO rehagas sumas largas mentalmente y NO inventes "valores esperados" que no
  estén en la tabla. El Ejecutor ya hizo los cálculos.
- No declares que el entorno está corrupto ni supongas errores de ejecución si
  el Ejecutor no muestra un traceback.

Verifica:
1. Que la ejecución no haya producido errores.
2. Que el código use las 18 filas de la tabla y las columnas correctas.
3. Que las fórmulas sean correctas: ingresos = unidades_vendidas *
   precio_unitario_cop, y variación = ((junio - enero) / enero) * 100.
4. Que la salida incluya el bloque de CONTROL y que `CUADRE OK` sea `True`.
5. Que aparezcan los cinco resultados pedidos: ingresos por producto, ingresos
   por mes, producto más vendido, mes con mayores ingresos y variación
   enero-junio.

Si los cinco puntos se cumplen, APRUEBA usando las cifras de la salida; no
recalculles las sumas por tu cuenta.
Si falta algo, pide una corrección concreta y NO apruebes.

Al aprobar: escribe "APROBADO", un resumen ejecutivo breve con las cifras de la
salida y 2 recomendaciones para la microempresa, y TERMINADO en la última línea,
solo, sin comillas.
No reescribas el código completo. No muestres tu razonamiento interno ni notas
de proceso; responde directo.
"""


def mostrar_mensaje(mensaje) -> None:
    """Imprime un mensaje del equipo indicando qué agente lo emitió."""
    agente = getattr(mensaje, "source", "Sistema")
    titulo = "Ejecutor — salida real del código" if agente == "Ejecutor" else agente
    contenido = mensaje.content if isinstance(mensaje.content, str) else str(mensaje.content)
    print(f"\n{'=' * 78}")
    print(f"  {titulo}")
    print(f"{'=' * 78}")
    print(contenido)


async def main() -> None:
    # 1) Leer la API key de DeepSeek desde el entorno (.env).
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key or api_key == "tu_api_key_aqui":
        print("⚠️  No hay API key configurada.")
        print("   Abre el archivo .env y reemplaza 'tu_api_key_aqui' por tu clave de DeepSeek.")
        return

    # 2) Cliente del modelo (API compatible con OpenAI de DeepSeek).
    model_client = OpenAIChatCompletionClient(
        model=os.getenv("DEEPSEEK_MODEL", "deepseek-flash"),
        api_key=api_key,
        base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1"),
        # DeepSeek activa thinking mode por defecto y expone su razonamiento.
        # Lo desactivamos para que cada agente entregue solo su respuesta final.
        extra_body={"thinking": {"type": "disabled"}},
        # DeepSeek no está en el catálogo interno de AutoGen.
        model_info={
            "vision": False,
            "function_calling": True,
            "json_output": True,
            "structured_output": False,
            "family": ModelFamily.UNKNOWN,
        },
    )

    # 3) Ejecutor local de código y agente que lo controla.
    CARPETA_EJECUCION.mkdir(exist_ok=True)
    ejecutor_local = LocalCommandLineCodeExecutor(
        timeout=60,
        work_dir=CARPETA_EJECUCION,
    )

    try:
        # 4) Crear los tres agentes del flujo.
        analista = AssistantAgent(
            name="Analista",
            model_client=model_client,
            system_message=ANALISTA_PROMPT,
            # Sin herramientas de ejecución: el Analista SOLO escribe código.
        )
        ejecutor = CodeExecutorAgent(
            name="Ejecutor",
            code_executor=ejecutor_local,
            # Solo ejecuta bloques de código escritos por el Analista.
            sources=["Analista"],
        )
        revisor = AssistantAgent(
            name="Revisor",
            model_client=model_client,
            system_message=REVISOR_PROMPT,
        )

        # 5) Condición de parada: el Revisor cierra con TERMINADO; el límite
        #    de 20 mensajes evita un bucle infinito si algo falla.
        condicion_termino = (
            TextMentionTermination("TERMINADO")
            | MaxMessageTermination(max_messages=20)
        )

        # 6) Turnos: Analista escribe -> Ejecutor corre -> Revisor valida.
        equipo = RoundRobinGroupChat(
            participants=[analista, ejecutor, revisor],
            termination_condition=condicion_termino,
        )

        # 7) Ejecutar en vivo para mostrar el flujo completo durante la demo.
        print("🚀 Iniciando el equipo de análisis de ventas de café...")
        resultado: TaskResult | None = None
        async for elemento in equipo.run_stream(task=TAREA_INICIAL):
            if isinstance(elemento, TaskResult):
                resultado = elemento
            else:
                mostrar_mensaje(elemento)

        # 8) Mostrar el motivo de finalización.
        if resultado is not None:
            print(f"\n✅ Motivo de finalización: {resultado.stop_reason}")
            print(
                "   Si fue por 'TERMINADO', el Revisor aprobó; "
                "si no, se alcanzó el límite de 20 mensajes."
            )
    finally:
        # 9) Cerrar el cliente y limpiar el ejecutor local.
        await model_client.close()
        await ejecutor_local.stop()


if __name__ == "__main__":
    asyncio.run(main())
