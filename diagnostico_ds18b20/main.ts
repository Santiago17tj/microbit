// Diagnostico: busca el DS18B20 en todos los pines libres y dice donde responde.
// Pantalla: nombre del pin si lo encuentra, X si no responde en ninguno.
let pines = [DigitalPin.P0, DigitalPin.P1, DigitalPin.P2, DigitalPin.P8, DigitalPin.P12, DigitalPin.P13, DigitalPin.P14, DigitalPin.P15, DigitalPin.P16]
let nombres = ["P0", "P1", "P2", "P8", "P12", "P13", "P14", "P15", "P16"]
basic.forever(function () {
    let encontrado = false
    for (let i = 0; i < pines.length; i++) {
        let t = dstemp.celsius(pines[i])
        if (t > -55 && t < 125) {
            serial.writeLine("DS18B20 RESPONDE en " + nombres[i] + ": " + t + " C")
            basic.showString(nombres[i] + " " + Math.round(t))
            encontrado = true
        }
    }
    if (!encontrado) {
        serial.writeLine("DS18B20 NO responde en ningun pin (P0,P1,P2,P8,P12-P16)")
        basic.showIcon(IconNames.No)
    }
    basic.pause(2000)
})
