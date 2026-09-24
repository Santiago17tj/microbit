# Monitor de compostaje con micro:bit → Excel

```
DS18B20 → micro:bit P0      FC-28 → micro:bit P1
micro:bit → USB/Serial (115200) → Python → Excel (tabla + gráfica en tiempo real)
```

La micro:bit envía cada ~15-20 s una línea: `TEMP:33,HUM:58,RAW:430`

## Archivos

| Archivo | Para qué |
|---|---|
| `main.ts` + `pxt.json` | Programa MakeCode (usa la extensión DS18B20 `dstemp`, `github:bsiever/microbit-dstemp`) |
| `DataStream_compost.py` | **El que hay que ejecutar.** Adaptación del `DataStream_V2_var2.py` del profesor para temperatura + humedad |
| `verificar_instalacion.py` / `1_VERIFICAR.bat` | Instala pyserial/pywin32, respalda el archivo del profesor, detecta el COM, prueba serial y Excel |
| `2_INICIAR_COMPOST.bat` | Lanza `DataStream_compost.py` con doble clic |
| `cargar_hex.py` | Copia el `.hex` a la unidad MICROBIT y comprueba que no haya `FAIL.TXT` |

El archivo del profesor **no se modifica**. `verificar_instalacion.py` guarda una copia en `respaldo/`.

## Pasos en Windows

1. Copia esta carpeta a tu PC junto a `DataStream_V2_var2.py`.
2. **Cargar el programa en la micro:bit** (sin WebUSB):
   - Opción A: en MakeCode → *Importar* → *Importar URL* → pega la URL de este repositorio de GitHub. La extensión DS18B20 se instala sola. Luego *Descargar* → si sale "No se pudo conectar", pulsa **Descargar como archivo**.
   - Opción B: descarga `compostaje-microbit.hex` desde la pestaña *Actions* de GitHub (artefacto `compostaje-microbit-hex`).
   - Arrastra el `.hex` a la unidad **MICROBIT**, o ejecuta `python cargar_hex.py archivo.hex`.
3. Doble clic en `1_VERIFICAR.bat`. Debe acabar en "Todo listo".
4. Doble clic en `2_INICIAR_COMPOST.bat`. Se abre Excel con la hoja *Compostaje*, la tabla y la gráfica.

Opciones: `python DataStream_compost.py --puerto COM5` (COM manual), `--simular` (probar Excel sin micro:bit), `--sin-excel`.
También puedes fijar el COM en la línea `PUERTO_MANUAL = None` del script.

Los datos se guardan en `datos_compost/compost_FECHA.xlsx` (cada 10 filas y al pulsar Ctrl+C) y en un `.csv` de respaldo que se escribe fila a fila.

## Problemas típicos

- **MakeCode: "No se pudo conectar / está siendo utilizado por otra aplicación"**: cierra el Python, Tera Term, el monitor serie u otras pestañas de MakeCode. Solo un programa puede usar la micro:bit a la vez. O usa *Descargar como archivo*.
- **"No se pudo abrir COMx"**: MakeCode (WebUSB) u otro programa tiene el puerto abierto.
- **Temperatura = ERROR SENSOR** (la micro:bit envía `-Infinity`): revisa el DS18B20. Rojo → 3V, negro → GND, amarillo → P0, **resistencia de 4.7 kΩ entre amarillo y 3V**.
- **Humedad siempre 100 %**: la fórmula provisional es `map(RAW, 1023→0 %, 250→100 %)`, así que cualquier RAW ≤ 250 da 100 %. Mira la columna *Humedad RAW* y el resumen que imprime el script al cerrarse (RAW min/max/media). Mide el RAW con el sensor en seco (aire) y en agua, y cambia `1023` y `250` en `main.ts` por esos valores.
