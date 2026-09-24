// Monitor de compostaje
// DS18B20 -> P0 (con resistencia pull-up de 4.7 kOhm entre DATA y 3V)
// FC-28   -> P1 (salida analogica A0)
// Envia por USB/serial (115200) lineas: TEMP:33,HUM:58,RAW:430
let tempC = 0
let humAnalog = 0
let humPct = 0
basic.forever(function () {
    tempC = dstemp.celsius(DigitalPin.P0)
    humAnalog = pins.analogReadPin(AnalogPin.P1)
    humPct = Math.constrain(
        Math.map(humAnalog, 1023, 250, 0, 100),
        0,
        100
    )
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
    serial.writeLine(
        "TEMP:" + Math.round(tempC) +
        ",HUM:" + Math.round(humPct) +
        ",RAW:" + humAnalog
    )
    basic.pause(1000)
})
