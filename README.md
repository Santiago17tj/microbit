# Monitor de compostaje con micro:bit → Excel

```
DS18B20 → micro:bit P0      FC-28 → micro:bit P1
micro:bit → USB/Serial (115200) → Python → Excel (tabla + gráfica en tiempo real)
                                          → Panel web en vivo (http://localhost:8765)
```

La micro:bit envía cada 10 s una línea: `TEMP:33,HUM:58,RAW:430` (`INTERVALO_MS` en `main.ts`).
Al conectarse, Python le pide un dato (`DATO?`) para que el primero llegue enseguida.

## Archivos

| Archivo | Para qué |
|---|---|
| `main.ts` + `pxt.json` | Programa MakeCode (usa la extensión DS18B20 `dstemp`, `github:bsiever/microbit-dstemp`) |
| `DataStream_compost.py` | **El que hay que ejecutar.** Adaptación del `DataStream_V2_var2.py` del profesor para temperatura + humedad |
| `verificar_instalacion.py` / `1_VERIFICAR.bat` | Instala pyserial/pywin32, respalda el archivo del profesor, detecta el COM, prueba serial y Excel |
| `2_INICIAR_COMPOST.bat` | Lanza `DataStream_compost.py` con doble clic (abre Excel **y** el panel web) |
| `3_DEMO_PANEL.bat` | Demo del panel con datos simulados, sin micro:bit ni Excel (para presentar) |
| `panel_web.py` + `panel/index.html` | Panel web local en tiempo real (solo librería estándar, funciona sin internet) |
| `ventanas.py` | Coloca el panel (izquierda) y Excel (derecha) lado a lado al iniciar |
| `cargar_hex.py` | Copia el `.hex` a la unidad MICROBIT y comprueba que no haya `FAIL.TXT` (lo usa también `DataStream_compost.py`) |
| `diagnostico_ds18b20/` → `diagnostico-ds18b20.hex` | Busca el DS18B20 en todos los pines. Pantalla: `P0 24` si responde, ❌ si no |
| `provisional_temp_interna/` → `PROVISIONAL-temp-interna.hex` | **Provisional:** usa el sensor interno de la micro:bit (mide el chip, no el compost) mientras se cambia el DS18B20 |

Los tres `.hex` se compilan solos en GitHub Actions (artefacto `compostaje-microbit-hex`).
`PROVISIONAL-temp-interna.hex` y `compostaje-microbit.hex` también están en el repositorio, para que la carga automática funcione en PC sin Node (p. ej. en el colegio). Si cambias un `main.ts`, vuelve a compilarlo (el script lo hace solo si hay Node) y súbelo.

El archivo del profesor **no se modifica**. `verificar_instalacion.py` guarda una copia en `respaldo/`.

## Carga automática del programa

Al iniciar, `DataStream_compost.py` espera a que se conecte la micro:bit y le pregunta qué programa tiene: envía `ID?` y ella responde, por ejemplo, `ID:PROVISIONAL-temp-interna v1`.

- Si ya tiene el programa correcto, no se toca y empieza a leer datos.
- Si tiene otro programa, o una versión anterior sin identificación, le carga `HEX_A_CARGAR` en la unidad MICROBIT, comprueba que no aparezca `FAIL.TXT` y verifica que ahora responda bien.

`HEX_A_CARGAR` está al principio del script. Ahora es `PROVISIONAL-temp-interna.hex`; cuando se reemplace el DS18B20, cámbialo por `compostaje-microbit.hex`.

El `.hex` se busca en esta carpeta. Si `main.ts` es más nuevo que el `.hex` (por ejemplo, después de calibrar el FC-28), se recompila solo con `npx makecode build` (necesita Node) y se carga. Si no está el `.hex` y no hay Node, se busca en *Descargas* el artefacto de GitHub Actions.

**Si modificas un `main.ts`, sube la versión en `PROGRAMA_ID`** (`v1` → `v2`) para que las micro:bit que ya tienen la versión anterior se actualicen.

Opciones: `--sin-cargar` (no comprobar ni cargar nada), `--hex compostaje-microbit.hex` (usar otro programa).

## Pasos en Windows

1. Copia esta carpeta a tu PC junto a `DataStream_V2_var2.py`.
2. **Cargar el programa en la micro:bit**: normalmente no hace falta, porque `2_INICIAR_COMPOST.bat` lo carga solo (ver arriba). A mano (sin WebUSB):
   - Opción A: en MakeCode → *Importar* → *Importar URL* → pega la URL de este repositorio de GitHub. La extensión DS18B20 se instala sola. Luego *Descargar* → si sale "No se pudo conectar", pulsa **Descargar como archivo**.
   - Opción B: descarga `compostaje-microbit.hex` desde la pestaña *Actions* de GitHub (artefacto `compostaje-microbit-hex`).
   - Arrastra el `.hex` a la unidad **MICROBIT**, o ejecuta `python cargar_hex.py archivo.hex`.
3. Doble clic en `1_VERIFICAR.bat`. Debe acabar en "Todo listo".
4. Doble clic en `2_INICIAR_COMPOST.bat`. Se abren a la vez el panel web (mitad izquierda de la pantalla) y Excel con la hoja *Compostaje*, la tabla y la gráfica (mitad derecha).

Opciones: `python DataStream_compost.py --puerto COM5` (COM manual), `--simular` (probar sin micro:bit), `--sin-excel`, `--sin-panel`, `--puerto-web 8080`, `--sin-acomodar` (no colocar las ventanas).
También puedes fijar el COM en la línea `PUERTO_MANUAL = None` del script.

Los datos se guardan en `datos_compost/compost_FECHA.xlsx` (cada 10 filas y al pulsar Ctrl+C) y en un `.csv` de respaldo que se escribe fila a fila.

## Panel web

Al iniciar, el script abre `http://localhost:8765` en el navegador. Muestra:

- Temperatura y humedad actuales, con la **fase del compost** (mesófila, termófila, termófila alta, peligro) y la recomendación correspondiente. Usa los mismos umbrales que `main.ts`.
- Gráfica en vivo de las dos variables con las zonas de fase marcadas (últimos 15 min, 1 h o toda la sesión). Pasa el ratón por encima para ver cada lectura.
- Estadísticas de la sesión (mín/máx/media), las últimas lecturas y un registro de eventos: conexión, errores del DS18B20 y cambios de fase.
- Botón para descargar el CSV de la sesión. Tema claro/oscuro.

No necesita instalar nada más ni tener internet. Se abre como ventana propia de Edge o Chrome (sin pestañas) en la mitad izquierda de la pantalla, con Excel en la derecha (el zoom de Excel se ajusta para que quepan la tabla y la gráfica). Si vuelves a iniciar el programa, se reutiliza la misma ventana del panel. Si la cierras, abre `http://localhost:8765` mientras el script siga en marcha.

Si Excel está ocupado al abrirse (por ejemplo, con el aviso *Error de activación de productos*), el script no lo abandona: guarda los datos en el CSV y crea la hoja y coloca las ventanas en cuanto Excel responda.

## Problemas típicos

- **`2_INICIAR_COMPOST.bat` no hace nada**: casi siempre es que `python` es el acceso directo de Microsoft Store (no hace nada y no da error). Los `.bat` ahora prueban `python` y `py -3`, y si no hay ninguno lo dicen. Instala Python de python.org marcando *Add python.exe to PATH*, o desactiva `python.exe` en *Configuración → Aplicaciones → Alias de ejecución de aplicaciones*. Si ves `... esperando la primera linea de la micro:bit`, Python funciona y el problema es la micro:bit (carga el `.hex`).
- **"La llamada fue rechazada por el destinatario" / "Excel está ocupado"**: Excel no acepta datos mientras editas una celda o hay un aviso abierto (p. ej. el de activación de Office). Pulsa Esc o cierra el aviso y no escribas en la hoja mientras se registra. El programa reintenta y escribe las filas pendientes con la siguiente lectura; el CSV siempre las tiene.
- **Excel dice "Error de activación de productos"**: Office no está activado. Puede abrir avisos que bloquean a Excel y, pasado un tiempo, no dejar editar ni guardar. Inicia sesión en Office con la cuenta que tiene la licencia (colegio/universidad) o usa `--sin-excel` y abre después el `.csv`.
- **MakeCode: "No se pudo conectar / está siendo utilizado por otra aplicación"**: cierra el Python, Tera Term, el monitor serie u otras pestañas de MakeCode. Solo un programa puede usar la micro:bit a la vez. O usa *Descargar como archivo*.
- **"No se pudo abrir COMx"**: MakeCode (WebUSB) u otro programa tiene el puerto abierto.
- **Temperatura = ERROR SENSOR** (la micro:bit envía `-Infinity`): revisa el DS18B20. Rojo → 3V, negro → GND, amarillo → P0, **resistencia de 4.7 kΩ entre amarillo y 3V**.
- **Humedad siempre 100 %**: la fórmula provisional es `map(RAW, 1023→0 %, 250→100 %)`, así que cualquier RAW ≤ 250 da 100 %. Mira la columna *Humedad RAW* y el resumen que imprime el script al cerrarse (RAW min/max/media). Mide el RAW con el sensor en seco (aire) y en agua, y cambia `1023` y `250` en `main.ts` por esos valores.
