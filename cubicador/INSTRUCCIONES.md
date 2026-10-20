# Cubicador automático – instrucciones

## Instalación (una sola vez)
1. Abrir el cubicador `.xlsb` (ese formato admite macros).
2. `Alt + F11` → si ya existe un módulo anterior del cubicador, borrarlo (clic derecho → Quitar).
3. **Insertar → Módulo** y pegar todo el contenido de `CubicadorAutomatico_para_pegar.txt`.
4. Al inicio del código, llenar los destinatarios (correos o nombres como salen en Outlook, separados por `;`):
   - `PARA_TRANSPORTE`, `CC_TRANSPORTE`: correo de solicitud de vehículo.
   - `PARA_BODEGA`, `CC_BODEGA`: correo del corte a bodega.
   - `CARPETA_TXT` (opcional): carpeta donde SAP deja los TXT de remisión. Si se deja vacío, la macro la pregunta la primera vez y la recuerda.
5. Guardar.

## Uso diario
| Paso | Macro | Qué hace |
|---|---|---|
| 1 | `CubicarExportSAP` | Pide el export de SAP, el texto del corte (ej. `1 Corte Octubre 14`) y dos filtros: **desde qué fecha** tomar las pendientes para vehículo (vacío = todas) y **qué remisiones van a bodega** (1 = por fecha desde/hasta, 2 = posteriores a la última registrada, 3 = sin filtro). Crea `CUBICAR HOY`, `RESUMEN TRANSPORTE`, `ALERTAS` y la **vista previa** de los correos. |
| 2 | (revisar) | **`PREVIA BODEGA`**: asunto, texto, destinatarios y todas las líneas que irán a bodega, agrupadas por cliente, con columna TXT (OK/FALTA). **`PREVIA VEHICULOS`**: un bloque por vehículo con zona, destinos, descargues, unidades, peso, m³, % de capacidad y todas sus líneas. Se editan las celdas amarillas: Pedir (SI/NO), Tipo, Doc. Transporte, Nota adicional, Nota inicial; se pueden borrar líneas. Avisa "EXCEDE DESCARGUES" o "SOBRE CAPACIDAD". Si se filtra `CUBICAR HOY`, ejecutar `ArmarVistaPrevia` para rehacer las vistas. |
| 3 | `EnviarCorteBodega` | Arma los borradores a bodega exactamente con lo que hay en `PREVIA BODEGA` y adjunta los TXT. Al final pregunta si registra las remisiones como enviadas (`HISTORIAL BODEGA`). |
| 4 | `SolicitarVehiculo` | Arma un borrador por cada vehículo con Pedir = SI en `PREVIA VEHICULOS` (recalcula totales desde las líneas; no arma los que excedan descargues o tengan tipo no válido). Al final pregunta si marca las remisiones en `CRUCE`. |

Ningún correo se envía solo: todos quedan como borrador en Outlook, con la firma del usuario.

## Reglas
- **Remisiones pendientes** = entregas del export con cantidad > 0 que no están en `CRUCE`. Las que no van en un vehículo siguen saliendo los días siguientes.
- **Remisiones nuevas para bodega** = pendientes que no están en `HISTORIAL BODEGA`. La primera vez, el historial se llena con las remisiones que ya están en `CUBICAJE BODEGA`.
- **Cubicaje** = m³/und de `MATERIAL` × cantidad. **Valor** = `VALO MATERIAL` × cantidad. **Zona** = ciudad en `llanta_zona`.
- **Máximo de descargues** (puntos de entrega = direcciones distintas, columna Dirección): TURBO 13, SENCILLO 10, PATINETA 1, TRACTOMULA 1. Se cambia en la función `MaxDescargues`.
- **Regla de rutas**: si hay carga para Cúcuta, Cúcuta y Bucaramanga van solas en su(s) vehículo(s), sin Bogotá ni otras ciudades. Sin carga para Cúcuta, Bucaramanga se puede unir con Bogotá.
- **Saturación** = m³ ÷ m³ máximo de `CUBICAJE BODEGA!C3:C6` (ej. Sencillo 36). El nombre del vehículo en el correo usa la etiqueta de la columna A (ej. "SENCILLO 42").
- **Tablas de los correos**: las líneas salen agrupadas por Código SAP (o por Dirección si no hay código), en el orden en que aparece cada cliente y, dentro de cada uno, por número de remisión.
- **Peso** del correo = suma de pesos; **Destino** = ciudades distintas en el orden de las filas; **Doc. Transporte** sugerido = primera remisión visible (se puede cambiar).
- Los TXT se reconocen por el texto `Remisión: 29150…` dentro del archivo; si hay varios para la misma remisión se adjunta el más reciente.
- Si hay remisiones de PIRELLI y PROMETEON en el mismo corte, se crea un correo de bodega por empresa.

## Validación hecha
- Lógica de cubicaje (replicada en Python) con el export y el cubicador compartidos: mismas 38 remisiones / 123 líneas y mismos totales (85,9716 m³, 698 und, 9.296,04 kg, $226.306.752).
- Correo de vehículo No 2 del 14-oct: con las mismas líneas da 291 unidades, 4.196 kg y 38,71 m³ → 108 % sobre 36 m³, igual que el correo original.
- La macro VBA no se ha ejecutado en Excel desde aquí: probar primero en una copia.
