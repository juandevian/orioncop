

# Multas por pago extemporáneo.


* A una factura sólo se le debe cobrar una multa.
* La multa se debe cobrar a quienes no hayan pagado después de la fecha de vencimiento
* Se debe tener un servicio para las M. P. E. (multa por pago extemporáneo)


## Campos espesíficos en BD. 
- orifacturas
    * FechaMulta: date: NotNull: "1900-01-01" 
    * Multada: tinyint (1): NotNull: 0

- oriservicios
    * DiasPExtemporaneo


## CASOS DE USO

### Caso 001
---
Hacer recibo de caja con fecha igual o anterior  al fecha de vencimiento de la cuenta de cobro NO GENERA multa.
---

### Caso 002
---
Hacer recibo de caja con fecha después de fecha de vencimiento, DEBE VERIFICAR parámetros del servicio para determinar si cobrar o no la multa.
si (fecha_recibo >= fecha_vencimiento + dias_pago_extemporaneo)

Ej.
Si el servicio tiene el cobro de la multa 1 día después de la fecha de vencimiento, debe cobrar intereses a partir de la fecha de vencimiento + 1
Fecha de vencimiento 10 del mes
Dias pago extemporáneo 1
Fecha de inicio de cobro de multa = 11
---

### Caso 003
---
Hacer recibo de caja por servivicio de identificación de un mes anterior (Nó administración) NO DEBE GENERAR multa.

---
NOTA: En el código para la validación de la fecha de cobro de la multa está: '
'''
Cobra multa si Fecha Recibo > a FechaVencimiento + Dias Pago Extemporáneo.

Cobra multa si Fecha Recibo > o = a FechaVencimiento + Dias Pago Extemporáneo.
'''



Multa por pago extemporáneo (M. P. E.)

Sentencia para actualizar facturas.

UPDATE orifacturas SET Multada = TRUE WHERE FechaFactura <= "2026-03-31 00:00:00";
