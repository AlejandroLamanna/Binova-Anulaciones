Sub Main
stop

'Funcion de Anulacion de Transaccion Sencilla 
set Tr = BinovaAnulaciones( self, "AnulaTransaccionSencilla", "", "", "", "", False, False, False, False, False, False, "", "", "", True )
'Funcion de Anulación de Transacción de Egreso de Valor
set xEgresoAnulado = BinovaAnulaciones( Transaccion, "anulaEgresoValor", "TP04", "", "", "EDAF0231-A22C-4FA4-88D4-5F7803F69637", False, False, True, False, True, False, "", "", "", "", False, "" )


End Sub



