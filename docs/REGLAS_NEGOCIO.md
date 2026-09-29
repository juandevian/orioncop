# Reglas de negocio - orioncop

Este documento consolida reglas de negocio funcionales encontradas en el código de OrionCop.

## Referencias base

- `.github/PLAN_MAESTRO.md`
- `.github/ESTRATEGIA_REPOSITORIOS_GITHUB.md`
- `docs/PLAN_MAESTRO.md`
- `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`

## Reglas

1. **Anulación de recibo de caja:** solo es anulable si el documento esta en periodo actual, cumple política de fecha, y no tiene anticipo reintegrado.
2. **Cobro sin deuda:** si no hay deuda pendiente y no es recibo solo de anticipo, se informa que no hay deuda para pago.
3. **Saldo no negativo:** el saldo calculado del recibo nunca puede quedar en negativo (se normaliza a 0).
4. **Anticipo por sobrepago:** cuando pago + descuentos supera deuda, el excedente se convierte en anticipo.
5. **Aplicacion total de pago:** el valor de pago debe quedar totalmente aplicado; si queda diferencia se lanza error.
6. **Estado de factura por fechas:** si hoy supera fecha de vencimiento, la factura pasa a `PeriodoGracia` o `Vencida` segun fecha de gracia.
7. **Fechas en factura automatica:** la fecha de vencimiento/gracia se toma por minima entre items; la fecha de gracia no puede ser menor que vencimiento.
8. **Descuento por valor en nota credito:** solo aplica si existe deuda del tipo (capital/interes/IVA), no permite repetir mismo tipo, y el valor debe ser > 0 y <= deuda.
9. **Anulacion fuera de periodo:** cuando se exige fecha de hoy, no se permite anular factura por fuera del periodo actual.
10. **Vencimiento/gracia de servicio:** se calcula por reglas de fin de mes o dias configurados sobre fecha de facturacion/vencimiento.
11. **Ajuste de cuotas administracion:** solo aplica para cuotas de administracion anuales, no ajustadas, y del ano actual.
12. **Concepto unico en servicios permanentes:** el concepto debe ser unico para servicios permanentes.
13. **EsAjuste restringido:** solo puede ser `True` cuando el tipo de servicio es anual.
14. **Tope del interés de mora (Ley 675 de 2001, art. 30):** el interés de mora no puede superar 1,5 veces el interés bancario corriente (IBC) certificado. La parametrización lo impide: si se digita un valor mayor se avisa y se deja el máximo. En modo Fijo la tasa deseada se conserva y la tasa efectiva es `Min(deseada, 1,5 x IBC)`, que se recalcula con cada IBC.
15. **Modo Variable:** tasa = IBC x factor, con factor definido por el usuario en el rango (0, 1,5]. Modo *Sin parametrizar* (por defecto) no consulta la API ni bloquea el cierre.
16. **Sincronización del IBC en el Cierre de mes:** antes de cerrar el mes y de causar intereses se actualiza `OriTasasMora` (la causación lee la tasa del día anterior a su fecha, por eso se usa el certificado vigente ese día). Si no se puede sincronizar (API caída y sin certificado local vigente) no se cierra el mes ni se causan intereses. La sincronización manual se puede hacer en cualquier momento.
17. **Tasa de mora anual simple:** `TasaMora` se guarda como fracción anual simple; la mora se calcula `deuda x (tasa / 365|366) x días`, y el equivalente mensual mostrado es `anual / 12`.

## Evidencia en codigo

- `OrionCopLsIbcCalculo.vb`, `OrionCopLsIbcSincroniza.vb`, `OrionCopLsOrionCop.vb` (FblnSincronizaIbc)
- `OrionCopLsItemFactura.vb` (FdecIntereseMora), `OrionCopLsCentroUtilidadOriCop.vb` (FdblTasaMoraFecha)
- `OrionCopL\clsReciboCaja.vb:246-267`
- `OrionCopL\clsReciboCaja.vb:453-455`
- `OrionCopL\clsReciboCaja.vb:470-472`
- `OrionCopL\clsReciboCaja.vb:500-505`
- `OrionCopL\clsReciboCaja.vb:701-703`
- `OrionCopL\clsFactura.vb:271-278`
- `OrionCopL\clsFactura.vb:600-626`
- `OrionCopL\clsNotaCr.vb:835-894`
- `OrionCopL\clsNotaCr.vb:1062-1064`
- `OrionCopL\clsServicio.vb:347-350`
- `OrionCopL\clsServicio.vb:370-389`
- `OrionCopL\clsServicio.vb:849-859`
- `OrionCopL\clsServicio.vb:1491-1493`
- `OrionCopL\clsServicio.vb:1524-1526`
