// VERSION PROVISIONAL: usa el sensor de temperatura INTERNO de la micro:bit
// (mide el chip, NO el compost) mientras se reemplaza el DS18B20.
// Mismo formato serial que la version final, para probar Python y Excel.
// FC-28   -> P1 (salida analogica A0)
// Envia por USB/serial (115200) lineas: TEMP:33,HUM:58,RAW:430
// cada INTERVALO_MS, independientemente de lo que muestre la pantalla.
let tempC = 0
let humAnalog = 0
let humPct = 0
let pedido = false
let ultimoEnvio = 0
const INTERVALO_MS = 10000   // cada cuanto se envia una lectura por USB
// Identificacion: DataStream_compost.py envia "ID?" y la micro:bit responde
// "ID:<programa>" para saber si ya tiene el programa correcto o hay que cargarlo.
// Si cambias este programa (p. ej. al calibrar el FC-28), sube la version (v2 -> v3).
const PROGRAMA_ID = "PROVISIONAL-temp-interna v2"

function leerSensores() {
    tempC = input.temperature()
    humAnalog = pins.analogReadPin(AnalogPin.P1)
    humPct = Math.constrain(
        Math.map(humAnalog, 1023, 250, 0, 100),
        0,
        100
    )
}

// Ordenes del PC: "ID?" -> identificacion, "DATO?" -> una lectura ya mismo
// (asi Python no espera un ciclo entero al conectarse)
serial.onDataReceived(serial.delimiters(Delimiters.NewLine), function () {
    const orden = serial.readUntil(serial.delimiters(Delimiters.NewLine)).trim()
    if (orden == "ID?") {
        serial.writeLine("ID:" + PROGRAMA_ID)
    } else if (orden == "DATO?") {
        pedido = true
    }
})

leerSensores()

// Medir y enviar: el unico bucle que lee los sensores
basic.forever(function () {
    if (pedido || input.runningTime() - ultimoEnvio >= INTERVALO_MS) {
        pedido = false
        ultimoEnvio = input.runningTime()
        leerSensores()
        serial.writeLine(
            "TEMP:" + Math.round(tempC) +
            ",HUM:" + Math.round(humPct) +
            ",RAW:" + humAnalog
        )
    }
    basic.pause(100)
})

// Pantalla: muestra la ultima lectura (no retrasa el envio de datos)
basic.forever(function () {
    basic.showString("T:")
    basic.showNumber(Math.round(tempC))
    basic.pause(500)
    if (tempC > 65) {
        basic.showString("PELIGRO CALOR")
        basic.showIcon(IconNames.Skull)
    } else if (tempC > 60) {
        basic.showString("TERMOFILA ALTA")
        basic.showIcon(IconNames.Surprised)
    } else if (tempC >= 40) {
        basic.showString("TERMOFILA")
        basic.showIcon(IconNames.Happy)
    } else {
        basic.showString("MESOFILA")
        basic.showIcon(IconNames.Sad)
    }
    basic.pause(1000)
    basic.showString("H:")
    basic.showNumber(Math.round(humPct))
    basic.pause(500)
    if (humPct <= 30) {
        basic.showString("SECO")
        basic.showIcon(IconNames.Sad)
    } else if (humPct <= 70) {
        basic.showString("HUMEDAD OK")
        basic.showIcon(IconNames.Yes)
    } else {
        basic.showString("HUMEDO")
        basic.showIcon(IconNames.Surprised)
    }
    basic.pause(1000)
})
