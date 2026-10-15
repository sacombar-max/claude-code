# Cubicador automático – instrucciones

## Instalación (una sola vez)
1. Abrir el cubicador (`.xlsb`; ese formato ya admite macros).
2. `Alt + F11` → menú **Archivo → Importar archivo…** → seleccionar `CubicadorAutomatico.bas`.
3. En el módulo, completar `CORREO_TRANSPORTADORA` (y `CORREO_COPIA` si aplica).
4. Opcional: crear botones en una hoja (Programador → Insertar → Botón) y asignarles las 3 macros.
5. Guardar el archivo.

## Uso diario
| Paso | Macro | Qué hace |
|---|---|---|
| 1 | `CubicarExportSAP` | Pide el export de SAP y el texto de la fecha de corte. Toma solo las entregas que **no** están en `CRUCE`, calcula cubicaje, valor y zona, y crea las hojas `CUBICAR HOY`, `RESUMEN TRANSPORTE` y `ALERTAS`. |
| 2 | (revisar) | Revisar `ALERTAS` y `RESUMEN TRANSPORTE` (saturación por vehículo y vehículo sugerido, por zona y por ciudad). |
| 3 | `GenerarCorreoTransporte` | Abre un borrador de Outlook con el resumen. **No lo envía**: se revisa y se envía a mano. |
| 4 | `MarcarRemisionesCubicadas` | Después de enviar la solicitud, agrega las remisiones a `CRUCE` para que no se vuelvan a cubicar. |

## Reglas que aplica
- **Remisiones nuevas** = entregas del export con cantidad > 0 que no están en `CRUCE`.
- **Cubicaje** = m³/und de la hoja `MATERIAL` × cantidad. **Valor** = valor de `VALO MATERIAL` × cantidad.
- **Zona** = ciudad buscada en `llanta_zona` (ignora tildes, comas y puntos: "BOGOTA D.C." = "BOGOTA, D.C.").
- **Vehículo sugerido** = el más pequeño cuya saturación sea ≤ `TOPE_SATURACION` (100 % por defecto). Capacidades leídas de `CUBICAJE BODEGA!B3:C6`. Si no cabe en el más grande, indica cuántos se necesitan y marca "revisar".
- Solo considera volumen (m³), no peso máximo por vehículo.

## Validación hecha
Con el export del 13-oct-2026 y el cubicador compartido, la lógica (replicada en Python) dio las mismas 38 remisiones / 123 líneas que el cubicador y los mismos totales: 85,9716 m³, 698 unidades, 9.296,04 kg y $226.306.752. La macro VBA en sí no se ha ejecutado todavía en Excel: probarla primero en una copia del archivo.
