# Automatización de procesamiento de facturas — Prueba práctica

## Qué hace

Un script de Python (`scripts/procesar\_facturas.py`) reemplaza el trabajo manual de:
abrir cada factura PDF, copiar sus datos a Excel, revisar los cálculos y anotar
observaciones. El script:

1. **Extrae** de cada PDF: folio, cliente, RFC, fecha, concepto, subtotal, IVA y total
(usando `pdfplumber` + expresiones regulares — las facturas tienen un formato de texto
consistente, así que regex es suficiente y más barato/confiable que un LLM para este paso).
2. **Valida** cada factura: campos faltantes, y que `IVA == 16% del subtotal` y
`total == subtotal + IVA`. Esto detecta automáticamente:

   * **F-1003**: RFC faltante.
   * **F-1004**: el total impreso ($17,800.00) no coincide con subtotal + IVA
($15,400.00 + $2,464.00 = **$17,864.00**) — diferencia de $64.00. Es una
inconsistencia de captura/impresión en la factura original, no un error del script.
3. **Incorpora el correo del cliente** (`Correo\_Cliente\_F1003.txt`): busca un folio
(patrón `F-####`) y un RFC con el patrón estándar mexicano en el texto del correo, y si
el campo estaba vacío en la factura, lo completa y lo marca como "actualizado vía correo".
4. **Consolida** todo en `data/Registro\_Consolidado.xlsx`, con el mismo layout que
`Registro\_Manual\_Actual.xlsx` (mismas columnas), más:

   * Relleno de color: verde = sin incidencias, naranja = con incidencias.
   * Columna `Estatus` (OK / Con incidencias) y `Observaciones` con el detalle.
   * Fila de totales calculada con fórmulas `SUM()` (no valores fijos), para que se
recalcule si se edita alguna fila.
5. **Genera un resumen ejecutivo** (`data/Resumen\_Ejecutivo.txt`): documentos procesados,
incidencias encontradas, correcciones aplicadas y acciones sugeridas.

## Cómo se ejecuta

```bash
pip install pdfplumber openpyxl pandas   # si no están instalados

python scripts/procesar\_facturas.py \\
    --facturas\_dir /ruta/a/las/facturas \\
    --correos\_dir  /ruta/a/los/correos \\
    --salida       data/Registro\_Consolidado.xlsx \\
    --resumen      data/Resumen\_Ejecutivo.txt
```

Sin argumentos, usa por defecto `/mnt/user-data/uploads` (donde están los archivos de
esta prueba) y escribe en `data/`. Espera facturas nombradas `Factura\_\*.pdf` y correos
`\*.txt` en las carpetas indicadas.

## Resultado en esta prueba

* 5 facturas procesadas, 3 sin incidencias, 2 con incidencias.
* F-1003: RFC resuelto automáticamente con el correo del cliente.
* F-1004: incidencia de total que **requiere confirmación humana** con el emisor (el
script señala el problema pero no "corrige" un monto sin que alguien lo confirme).
* Ver `data/Resumen\_Ejecutivo.txt` para el detalle completo.

## Decisiones y por qué

* **Regex sobre LLM para la extracción**: las facturas son un formato fijo y repetitivo.
Un LLM aquí sería más caro, más lento y menos determinista sin aportar precisión
adicional. El LLM se usó donde sí aporta valor: entender el
proceso completo, decidir las reglas de validación de negocio, redactar el resumen
ejecutivo y diseñar la propuesta de automatización — no para "leer" campos con formato fijo.
* **openpyxl con fórmulas** en vez de valores fijos en la fila de totales, para que el
archivo siga siendo confiable si alguien edita una fila a mano después.
* **No sobrescribir el total de F-1004 automáticamente**: un desfase de $64 puede
deberse a un error de captura en cualquiera de los tres campos (subtotal, IVA o total).
Sin más contexto, la acción correcta es señalarlo y pedir confirmación, no adivinar
cuál de los tres números está mal.
* **Reglas simples de texto para el correo** (folio + patrón de RFC) en vez de un LLM,
por la misma razón: es un caso de un campo estructurado (RFC) dentro de texto libre,
y una regex es suficiente, auditable y gratis. Si los correos fueran más variados o
ambiguos en producción, ahí sí un LLM aportaría valor real (ver siguiente sección).

## Qué falta para producción / cómo lo completaría

Esta prueba resuelve el caso con 5 facturas y 1 correo en un solo lote. Para uso diario
real, propondría:

1. **Disparador automático**: un flujo en n8n o Power Automate que monitoree una carpeta
de red / bandeja de correo / carpeta de Drive y dispare el procesamiento cuando llegue
una factura o un correo nuevo (en vez de correrlo manualmente).
2. **OCR para facturas escaneadas**: si algunas facturas llegan como imagen (no como PDF
con texto), agregar un paso de OCR (ej. Tesseract o un servicio en la nube) antes de la
extracción. El validador de consistencia (subtotal+IVA=total) seguiría funcionando igual
y serviría también para detectar errores de OCR.
3. **LLM para correos ambiguos**: la regla de texto actual asume que el correo menciona
el folio explícitamente y un RFC con formato estándar. En producción, los correos de
clientes son mucho más variados ("la factura de la semana pasada", sin folio explícito,
con errores de redacción). Ahí un LLM (con el registro actual como contexto) sí aporta
valor real: interpretar la intención del correo y proponer el cambio, dejando siempre
una persona que confirme antes de escribir al registro.
4. **Base de datos en vez de un Excel plano**: para volumen diario, mover el registro a
una base de datos (o Google Sheets vía API) con historial de cambios, en vez de
sobrescribir un archivo Excel cada vez.
5. **Validaciones de negocio adicionales**: folios duplicados, RFCs con formato inválido
(longitud/checksum), fechas fuera de rango, límites de monto que requieran aprobación
extra.
6. **Notificación automática**: cuando el script detecta una incidencia, enviar un correo
o mensaje (Slack/Teams) a la persona responsable en vez de solo dejarlo en un archivo.
7. **Pruebas automatizadas**: casos de prueba con facturas "buenas" y "con error" conocidos,
para evitar que un cambio futuro al script rompa la extracción o la validación en silencio.

Con más tiempo, el siguiente paso concreto sería envolver `procesar\_facturas.py` en un
flujo de n8n: nodo "Watch folder" → nodo "Execute Command" (llama al script) → nodo
"Send email" si `Con incidencias` > 0.

## Archivos de este entregable

```
README.md                          este archivo
scripts/procesar\_facturas.py       script de extracción, validación y consolidación
data/Registro\_Consolidado.xlsx     entregable: registro consolidado
data/Resumen\_Ejecutivo.txt         entregable: resumen ejecutivo
```

