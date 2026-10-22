# Contexto del negocio

## Empresa
Microempresa de café con tres productos:
- Café Molido 500 g — precio unitario $18.000 COP.
- Café en Grano 1kg — precio unitario $42.000 COP.
- Café Instantáneo — precio unitario $12.000 COP.

## Datos disponibles
Ventas mensuales de enero a junio (18 registros: 6 meses × 3 productos).
Columnas: mes, producto, unidades_vendidas, precio_unitario_cop.
Los ingresos se calculan como unidades_vendidas * precio_unitario_cop.

## Resultados de referencia
- Ingresos del período: $54.930.000 COP.
- Producto con mayores ingresos: Café en Grano 1kg ($23.520.000 COP).
- Producto más vendido en unidades: Café Instantáneo (1.290 unidades).
- Mes con mayores ingresos: Junio ($10.650.000 COP).
- Variación enero–junio: +34,47%.

## Archivos de evidencia
- ventas_detallado.xlsx: las 18 filas con ingresos.
- ventas_por_producto.xlsx: totales por producto, de mayor a menor.
- ventas_por_mes.xlsx: totales mensuales en orden cronológico.
- ventas_enero.xlsx ... ventas_junio.xlsx: un archivo por mes.
- ventas_anual.xlsx: consolidado del período, con hojas "Detalle", "Por producto" y "Por mes".
- Analista.txt, Ejecutor.txt, Revisor.txt: conversación completa.

## Instrucciones generales
- Responder de forma específica a lo que pida el usuario.
- Si pide un solo mes, generar únicamente ventas_<mes>.xlsx; no generar los
  archivos estándar, por producto, por mes ni de otros meses.
- Si pide el año o todas las ventas, generar los tres archivos estándar, un
  archivo por mes y ventas_anual.xlsx.
- Los datos actuales cubren enero–junio; el consolidado anual contiene esos seis meses.
