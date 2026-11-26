# Problemas y jugadas

Cada problema real abstraído, el razonamiento y la solución que se eligió («la jugada»). El patrón común: **cuando el dato depende de un tercero, no se arregla al tercero; se guarda la evidencia y se decide con ella.**

---

## 1. El dato desaparece justo cuando más importa

**Síntoma:** los aéreos que llegaban no salían en el correo de arribados. Tampoco algunos marítimos.

**Causa, en tres capas:**
1. El reporte de tránsito borra la línea cuando la carga llega al terminal.
2. El Order Match suelta el pedido (433) cuando la línea sale del tránsito.
3. El correo solo dejaba pasar embarques con pedido, y la mayoría de lo arribado era stock.

**Razonamiento:** no se puede pedir a la fuente que no borre. Lo que sí se puede es fotografiar a diario lo que la fuente muestra y guardar la versión.

**Jugada:**
- Un **historial versionado por línea** (Key × versión, con `visto_desde` y `visto_hasta`) y una copia diaria del cruce 422 → 433.
- La línea «salió» cuando su última versión ya no está en el corte de hoy.
- Un **rescate** desde los reportes semanales para los huecos.
- El correo deja pasar todo lo arribado. El pedido es visibilidad, no filtro.

**Evidencia:** 13 aéreos (28 líneas) en una semana eran invisibles entre el terminal y el almacén; ahora salen con su aviso, su ruta y su pedido.

## 2. El radar perdía el buque que ya había llegado

**Síntoma:** un marítimo en el puerto, sin salida al almacén, no aparecía en el radar.

**Causa:** el nombre del buque solo existía en el reporte del día. El historial guardaba el BL, pero no el buque.

**Jugada:**
- El historial guarda el buque, el contenedor y la `revised_eta`. Una copia anterior las recibe vacías, con aviso, sin romperse.
- El radar suma lo que la fuente ya soltó y marca el arribo con la regla del marítimo: el ATA del forwarder o la ETA vencida, porque el aviso no es arribo.
- El pedido sale del historial cuando el Order Match de hoy ya no lo tiene, con el mismo cruce que el correo.

## 3. Un cambio de ETA no dejaba rastro

**Síntoma:** un buque que se atrasaba 9 días simplemente dejaba de verse.

**Razonamiento:** la regla de negocio es que el compromiso se mide sobre la ETA más reciente. Para avisar que cambió hay que conservar también la anterior.

**Jugada:** un cambio de ETA abre una versión nueva del historial. La alerta compara la ETA nueva contra la anterior y avisa cuando la diferencia pasa de 5 días. Con datos sintéticos sale 1 alerta, la correcta, y ningún falso positivo.

## 4. Los terceros escriben como quieren

| Tercero y formato | Riesgo | Validación |
|---|---|---|
| Fecha ISO `2026-09-08` leída día primero | Se convertía en 9 de agosto | ISO = año-mes-día; el resto, día/mes. ATD a ±1 día en 12 de 12 |
| Aviso de llegada con PDF escaneado | No se registraba | Cuenta como evidencia, con la fecha del correo como llegada |
| El mismo PDF escaneado, cada día | Se descargaba y releía durante 30 días | Se recuerda el correo ya intentado |
| Respuestas y reenvíos | Duplicaban el aviso | Se ignoran; la llave es HAWB + id del correo (idempotente) |
| Aviso marítimo | Llega cuando la carga sale de planta | No cuenta como arribo; el arribo es el ATA o la ETA vencida |
| Embarque con dos modos (aéreo y marítimo) | El SLA quedaba vacío y nunca alertaba | Toma el SLA más estricto |
| Un país nuevo en el reporte | Tumbaba toda la cadena | Queda fuera del alcance y deja un aviso |

## 5. Fallas silenciosas

**Síntoma:** si la copia de pedidos fallaba (por ejemplo, un archivo bloqueado por la sincronización en la nube), el orquestador igual decía «OK» y se perdía ese día.

**Jugada:** hay tres resultados posibles: **OK**, **AVISO** y **FALLO**.
- Un AVISO no detiene la cadena, porque el correo igual tiene que salir, pero el resumen lo marca y la tarea programada lo ve.
- Si el historial no se puede leer, la cadena aborta. Nunca sigue con datos parciales.
- El reemplazo atómico de archivos reintenta 3 veces ante bloqueos breves.

## 6. Lo que pasa en el día no se veía hasta mañana

**Síntoma:** al refrescar el radar, las coordenadas quedaban en caché y no aparecía lo que el forwarder entregó hoy ni lo que el 3PL recibió.

**Razonamiento:** el radar es de lectura, así que no debería poder ensuciar la base.

**Jugada:** el radar horario copia la base a una carpeta temporal, procesa ahí solo los correos nuevos, arma el radar y borra la copia.
- **Prueba:** la base no cambia ni un byte y una recepción de las 13:10 se ve a las 14:00.
- **AIS:** la caché baja de 6 a 2 h. Se mantienen las pausas al azar de 8 a 15 s, el tope de 20 consultas y el freno ante 403/429/503, un ritmo parecido al de una persona.

## 7. Nueve scripts que se usaban entre sí

**Síntoma:** tres «jobs» eran a la vez librerías de otros. La tabla del correo se calculaba dos veces, el cruce de pedidos de 4 a 5 veces por día y había 5 versiones de la misma normalización de IDs.

**Jugada:**
- Se pasó a 5 módulos más un orquestador, con una sola configuración y la tabla calculada una vez.
- Las variantes que parecían duplicadas pero eran **reglas distintas a propósito** (estricto en el reporte diario, tolerante en el semanal) **no se unificaron**: unificarlas habría cambiado reglas validadas.
- No había ninguna función muerta: el volumen de código son reglas, no relleno.

## 8. Una consulta cara que casi no aportaba

**Jugada:** el modo `--sin-lake` salta la consulta de bultos, kg y m³ al data lake. El conteo de bultos del aviso se mantiene y todo lo demás sale igual (13/13).
