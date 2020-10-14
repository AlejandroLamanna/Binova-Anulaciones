
Function AnulaTransaccionSencilla(transaccion,idFlagAnulado,RevierteAsiento, Mensaje)
		set AnulaTransaccionSencilla = Nothing
        Dim xelem , xview , xpend , xAsiento , xNuevoAsiento 
        Dim ok 
        set xws = transaccion.workspace
        ok = False
        If idFlagAnulado <> "" Then
            If transaccion.Flag Is Nothing Then
                ok = True
            Else
                If transaccion.Flag.ID <> idFlagAnulado Then
                    ok = True
                End If
            End If
        Else
            ok = True
        End If
        If Not ok Then
            MsgBox("Esta transacción ya ha sido anulada")
        End If
        If ok Then
            'Verifico que el pendiente esté sin cancelar nada.

            set xview = NewCompoundView(transaccion, "PENDIENTE", xws, Nothing, False)
            xview.addfilter(NewFilterSpec(xview.columnfrompath("TRANSACCION"), "=", transaccion)) 'Filtro los pendientes generados por la transaccion en la que estoy posicionado
            For Each xitem In xview.viewitems
                For Each xitempend In xitem.bo.items
                    If ok Then
                        If xitempend.cantpend2_cantidad < xitempend.cantori2_cantidad Or xitempend.valorpe2_importe < xitempend.valoror2_importe Then
                            ok = False
                            'Verifico que no se haya saldado contra si mismo.
                            For Each xitemCancelacion In xitempend.cancelaciones
                                If xitemCancelacion.itemorigen.id = xitempend.id Then 'Se saldó con click derecho 
                                    ok = True
                                End If
                            Next
                        End If
                    End If
                Next
            Next
        End If
        If Not ok Then
            MsgBox("Esta transacción ha generado transacciones posteriores, debe anularlas previamente.")
        End If

        If ok Then
            'esta parte remueve la cancelación del pendinete de la tr origen
            For Each xitemtransaccion In transaccion.ItemsTransaccion
                If Not xitemtransaccion.itempendiente Is Nothing Then
                    For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                        If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                            xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                            xcancelacion.Delete()
                        End If
                    Next
                End If
            Next
            '---------------------------------------------------------------------------------------------------
            'esta parte remueve los pendientes generados por la transaccion
            For Each xItemTR In transaccion.ItemsTransaccion
                set xelem = xItemTR
                set xview = NewCompoundView(transaccion, "ITEMTR", xws, Nothing, True)
                xview.addfilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", xelem))
                For Each xelem In xview.viewitems
                    set xpend = xelem.bo
                    SaldarPendiente(xpend)
                Next
            Next
            If RevierteAsiento Then
                set xAsiento = GetAsiento(transaccion)
                If Not xAsiento Is Nothing Then
                    set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, transaccion.FECHAACTUAL)
                End If
            End If
            If idFlagAnulado <> "" Then
                set transaccion.Flag = ExisteBO(transaccion, "FLAG", "ID", idFlagAnulado, Nothing, True, False, "=")
            End If
            transaccion.Detalle = transaccion.Detalle + "El usuario " + NombreUsuario() + " canceló este comprobante. " & Now()
            If Mensaje Then
                MsgBox("EL comprobante: " & transaccion.Name & " ha sido cancelado ")
            End If
            set AnulaTransaccionSencilla = Transaccion
        Else
            set AnulaTransaccionSencilla = Nothing
        End If
End Function


Function AnulaTRCompra(atransaccion, xtipotrgenerar,xcodigoTRDesimputacion,ID_Flag_Anulado,RemuevePendientes,RevierteAsiento)
        Dim xAsiento, xview , xview2 
        Dim xtrdesimputacion , xFactura 
        Dim xImporte, xabortar, rta , xmensaje 
        Dim anio , mes , dia , fecha 
        Dim xws 
        set xws = atransaccion.workspace
        Dim CalipsoFunctions 
        Dim xUtil 
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)

        If classname(atransaccion) = "TRFACTURACOMPRA" Or classname(atransaccion) = "TRDEBITOCOMPRA" Or classname(atransaccion) = "TRCREDITOCOMPRA" Then
            If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
                xImporte = 0
                xabortar = False
                If atransaccion.compromisopago.SALDO2_IMPORTE <> atransaccion.compromisopago.IMTOTAL2_IMPORTE Then
                    If atransaccion.compromisopago.IMTOTAL2_IMPORTE > atransaccion.Compania.PARAMETROS.TOLERANCIACP Then
                        If classname(atransaccion.compromisopago) = "CPDEBITO" Or classname(atransaccion.compromisopago) = "CPFACTURA" Then
                            set xview2 = NewCompoundView(atransaccion, "ITEMORDENPAGO", xws, Nothing, True)
                            xview2.AddFilter(NewFilterSpec(xview2.ColumnFromPath("REFERENCIA"), "=", atransaccion.compromisopago))
                            xmensaje = ""
                            For Each xop In xview2.ViewItems
                                xmensaje = xmensaje & xop.bo.placeowner.nombre & Chr(13)
                                xabortar = True
                            Next
                            If xabortar Then
                                Call MsgBox("Esta Transaccion ya ha sido pagada total o parcialmente." & Chr(13) & "Debe anular o desimputar primero todos los pagos imputados a esta factura." & Chr(13) & "Transacciones Imputadas:" & Chr(13) & xmensaje, 0, "Proceso de Anulacion")
                                Exit Function
                            End If
                        End If
                        'si llego hasta aca significa que fue imputada a un credito.

                        rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Proveedor.", 1, "Proceso de Anulación")
                        If rta = 2 Then
                            Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                            Exit Function
                        End If
                        set xtrdesimputacion = DesimputarTransaccion( atransaccion,  xcodigoTRDesimputacion)
                        If xtrdesimputacion Is Nothing Then
                            Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                            Exit function 
                        End If
                    End If
                End If
                set xFactura = atransaccion
                anio = VB.DatePart("yyyy", VB.Now)
                mes = VB.DatePart("m", VB.Now)
                dia = VB.DatePart("d", VB.Now)
                fecha = anio & Right("00" & mes, 2) & Right("00" & dia, 2)
                '     xfactura.external_id = "Anulada - " & fecha
                Dim xVector , xBucket , xFactura_Anulada , xObjeto , xpend , xNuevoAsiento 
                Dim xNumero , xtabla 
                set xVector = NewVector()
                set xBucket = NewBucket()
                set xFactura_Anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
                xFactura_Anulada.NUMERODOCUMENTO = xFactura.NUMERODOCUMENTO
                xFactura_Anulada.VINCULOTR = xFactura
                xFactura_Anulada.BOExtension.compromisopagotranulada = xFactura.compromisopago
                xFactura_Anulada.Destinatario = xFactura.Destinatario
                xFactura_Anulada.Originante = xFactura.Originante
                xFactura_Anulada.Detalle = "Anulación de " & xFactura.nombre
                xFactura_Anulada.centrocostos = xFactura.centrocostos
                xFactura_Anulada.FECHAACTUAL = xFactura.FECHAACTUAL
                xFactura_Anulada.NOTA = xFactura.NOTA
                xFactura_Anulada.mediotransporte = xFactura.mediotransporte
                xFactura_Anulada.domicilio = xFactura.domicilio
                xFactura_Anulada.Total.unidadvalorizacion = xFactura.Total.unidadvalorizacion
                xFactura_Anulada.cotizacion = xFactura.cotizacion
                xFactura_Anulada.tipopago = xFactura.tipopago
                xFactura_Anulada.domicilioentrega = xFactura.domicilioentrega
                xFactura_Anulada.fecharegistro = xFactura.fecharegistro
                If classname(atransaccion) = "TRDEBITOCOMPRA" Or _
                   classname(atransaccion) = "TRCREDITOCOMPRA" Or _
                   classname(atransaccion) = "TRDEBITOVENTA" Or _
                   classname(atransaccion) = "TRCREDITOVENTA" Then
                    xFactura_Anulada.GENERADAPORDIFCAMBIO = xFactura.GENERADAPORDIFCAMBIO
                End If
                xNumero = "9" & Right(atransaccion.NUMERODOCUMENTO, 11)
                xtabla = classname(atransaccion)
                xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xNumero & "' WHERE ID = '" & atransaccion.ID & "'"
                xVector.Add(xBucket)
                ExecutarSQL xVector, "distrobj", "Anulada - " & fecha, xws, -1

                For Each Item In xFactura.ItemsTransaccion
                    xObjeto = CrearItemTransaccion(xFactura_Anulada)
                    xObjeto.Referencia = Item.Referencia
                    xObjeto.unidadmedida = Item.unidadmedida
                    xObjeto.cantidad.cantidad = Item.cantidad.cantidad
                    xObjeto.UNIDADMEDIDANOLINEAL = Item.UNIDADMEDIDANOLINEAL
                    xObjeto.cantidadnolineal.cantidad = Item.cantidadnolineal.cantidad
                    xObjeto.valor.Importe = Item.valor.Importe
                    xObjeto.bultos = Item.bultos
                    xObjeto.PORCENTAJEBONIFICACION = Item.PORCENTAJEBONIFICACION
                    xObjeto.centrocostos = Item.centrocostos
                    xObjeto.imputacioncontable = Item.imputacioncontable
                    xObjeto.Detalle = Item.Detalle
                    xObjeto.observacion.Memo = Item.observacion.Memo
                    For Each ximpuestoori In Item.IMPUESTOSITEMTRANSACCION
                        For Each ximpuestodest In xObjeto.IMPUESTOSITEMTRANSACCION
                            If ximpuestoori.definicionimpuesto.impuesto.ID = ximpuestodest.definicionimpuesto.impuesto.ID Then
                                ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                            End If
                        Next
                    Next
                Next
                For Each ximpuestoori In xFactura.IMPUESTOSTRANSACCION
                    For Each ximpuestodest In xFactura_Anulada.IMPUESTOSTRANSACCION
                        If ximpuestoori.definicionimpuesto.impuesto.nombre = ximpuestodest.definicionimpuesto.impuesto.nombre Then
                            ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                        End If
                    Next
                Next
                'saldo los pendientes que queden de esta transaccion
                For Each xItemTR In xFactura.ItemsTransaccion
                    xview = NewCompoundView(xFactura, "ITEMTR", xws, Nothing, True)
                    'Set xColumnaRelacion = xview.ColumnFromPath("RELACIONTRORIG")
                    'Set xFilter = NewFilterSpec(xColumnaRelacion, " = ", "{89C23657-3F01-11D5-86AD-0080AD403F5F}")
                    'xview.AddFilter (xFilter)
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", xItemTR.ID))
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("CANCELADO"), " = ", False))
                    For Each xpendiente In xview.ViewItems
                        xpend = xpendiente.bo
                        SaldarPendiente(xpend)
                    Next
                Next
                If RemuevePendientes Then
                    'Libero los pendientes que puedan haber generado la transaccion
                    For Each xitemtransaccion In atransaccion.ItemsTransaccion
                        If Not xitemtransaccion.itempendiente Is Nothing Then
                            For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                                If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                    xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                    xcancelacion.Delete()
                                End If
                            Next
                        End If
                    Next
                End If
                If RevierteAsiento Then
                    set xAsiento = GetAsiento(atransaccion)
                    If Not xAsiento Is Nothing Then
                        set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, xFactura_Anulada.fecharegistro)

                    End If
                End If
                ShowBO(xFactura_Anulada)
            Else

                MsgBox("Esta Transaccion ya fue anulada")
            End If

        End If

End Function


Function AnulaOrdenPago(atransaccion,xtipotrgenerar, xcodigoTRDesimputacion , ID_Flag_Anulado,aCodigoConceptoContableAnulacion,xtipoTrEgresoAGenerar,reemite)
        Dim xview , xtrdesimputacion, xitem 
        Dim xImporte, xAbortar , rta 
        Dim xws
        set xws = atransaccion.workspace
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)

        If classname(atransaccion) = "TRORDENPAGO" Then
            If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
                xImporte = 0
                xAbortar = False
                If Not atransaccion.compromisopago Is Nothing Then
                    If atransaccion.compromisopago.SALDO2_IMPORTE <> atransaccion.compromisopago.IMTOTAL2_IMPORTE Then
                        If atransaccion.compromisopago.IMTOTAL2_IMPORTE > atransaccion.Compania.PARAMETROS.TOLERANCIACP Then
                            rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Proveedor.", 1, "Proceso de Anulación")
                            If rta = 2 Then
                                Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                Exit Function
                            End If
                            set xtrdesimputacion = DesimputarTransaccion( atransaccion, xcodigoTRDesimputacion)
                            If xtrdesimputacion Is Nothing Then
                                Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                Exit Function
                            End If
                        End If
                    End If
                End If
                Dim fecha, xtransaccion_anulada 
                fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
                Dim xVector , xBucket , xtabla 
                set xVector = NewVector()
                set xBucket = NewBucket()
                xtabla = classname(atransaccion)
                xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "' WHERE ID = '" & atransaccion.ID & "'"
                xVector.Add(xBucket)
                ExecutarSQL xVector, "distrobj", "Anulada - " & fecha, xws, -1

                'Me fijo que no haya ningun pendiente generado por esta orden de pago.
                'Esta vista busca todos los pendientes que se pueden haber generado para un Egreso de Valores.
                'Los pendientes de asiento y de imputacion no se cancelan porque esos son anulados con contramovimientos
                Dim xTipoTr, xTipoTr_OTP , xObjeto 
                Dim xtrdelgada 
                set xTipoTr = NewColumnSpec("TIPOTRANSACCION", "ID", "")
                set xTipoTr_OTP = NewColumnSpec("TIPOTRANSACCION", "OTP", "")

                set xview = NewCompoundView(atransaccion, "PENDIENTE", xws, Nothing, True)
                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("TRANSACCION"), "=", atransaccion.ID))
                xview.AddJoin(NewJoinSpec(xview.ColumnFromPath("TIPOTRGENERAR"), xTipoTr, False))
                xview.AddFilter(NewFilterSpec(xTipoTr_OTP, "=", "TREGRESOVALORES"))
                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("SALDADA"), "=", False))

                For Each xitem In xview.ViewItems
                    SaldarPendienteCompleto(xitem.bo)
                Next

                '            If xsaldada Then
                set xtransaccion_anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
                xtransaccion_anulada.NUMERODOCUMENTO = atransaccion.NUMERODOCUMENTO
                xtransaccion_anulada.VINCULOTR = atransaccion
                xtransaccion_anulada.BOExtension.compromisopagotranulada = atransaccion.compromisopago
                xtransaccion_anulada.Destinatario = atransaccion.Destinatario
                xtransaccion_anulada.Originante = atransaccion.Originante
                xtransaccion_anulada.Detalle = "Anulación de " & atransaccion.nombre
                xtransaccion_anulada.centrocostos = atransaccion.centrocostos
                xtransaccion_anulada.FECHAACTUAL = atransaccion.FECHAACTUAL
                xtransaccion_anulada.NOTA = atransaccion.NOTA
                If xtransaccion_anulada.tipotransaccion.TDDef Is Nothing Then
                    xtrdelgada = False
                Else
                    xtrdelgada = True
                End If
                If xtrdelgada Then
                    xtransaccion_anulada.total_unidadvalorizacion = atransaccion.Total.unidadvalorizacion
                Else
                    xtransaccion_anulada.Total.unidadvalorizacion = atransaccion.Total.unidadvalorizacion
                End If
                xtransaccion_anulada.cotizacion = atransaccion.cotizacion
                xtransaccion_anulada.tipopago = Nothing 'Pongo el tipo de pago en Null para que no aparezca el default del proveedor.
                For Each Item In atransaccion.ItemsTransaccion
                    set xObjeto = CrearItemTransaccion(xtransaccion_anulada)
                    set xObjeto.Referencia = ExisteBO(atransaccion, "CONCEPTOCONTABLE", "CODIGO", aCodigoConceptoContableAnulacion, Nothing, True, False, "=")
                    If xtrdelgada Then
                        xObjeto.Cantidad_Cantidad = 1
                        xObjeto.Valor_Importe = Item.valor.Importe
                    Else
                        xObjeto.cantidad.cantidad = 1
                        xObjeto.valor.Importe = Item.valor.Importe
                    End If
                    xObjeto.centrocostos = Item.centrocostos
                    If Not Item.Referencia Is Nothing Then
                        xObjeto.Detalle = Item.Referencia.Name
                    End If
                Next
                ShowBO(xtransaccion_anulada)
                'Ahora hay que anular los egresos de valores que se pudieron haber generado.
                '                Set xrelacion = NewColumnSpec("RELACIONTRANSACCION", "ID", "")
                '                Set xrelacion_TipoTRGenerar = NewColumnSpec("RELACIONTRANSACCION", "TRANSACCIONDESTINO", "")

                '                Set xtipotr = NewColumnSpec("TIPOTRANSACCION", "ID", "")
                '                Set xtipotr_otp = NewColumnSpec("TIPOTRANSACCION", "OTP", "")
                Dim xitem_pendiente, xitem_pendiente_itemtr , xitem_op_placeowner , xitem_op 

                set xitem_pendiente = NewColumnSpec("ITEMTR", "ID", "")
                set xitem_pendiente_itemtr = NewColumnSpec("ITEMTR", "ITEMTRANSACCION", "")

                set xitem_op = NewColumnSpec("ITEMORDENPAGO", "ID", "")
                set xitem_op_placeowner = NewColumnSpec("ITEMORDENPAGO", "PLACEOWNER", "")


                set xview = NewCompoundView(atransaccion, "ITEMCANCELACIONFINANZAS", xws, Nothing, True)
                xview.AddJoin(NewJoinSpec(xview.ColumnFromPath("ITEMPENDIENTE"), xitem_pendiente, False))
                xview.AddJoin(NewJoinSpec(xitem_pendiente_itemtr, xitem_op, False))
                Dim xproceso , xDic , xtrdestino , xanulacionev 
                xview.AddFilter(NewFilterSpec(xitem_op_placeowner, "=", atransaccion))
                set xproceso = Nothing
                set xDic = CreateObject("SCRIPTING.DICTIONARY")
                'Este diccionario controla que no se anule dos veces la misma tr
                For Each Item In xview.ViewItems
                    set xtrdestino = Item.bo.Place.Owner
                    If Not xDic.Exists(xtrdestino.ID) Then
                        xDic.Add xtrdestino.ID, ""
                        'lo hacemos con una funcion para que se pueda anular solo el egreso.
                        set xanulacionev = anulaEgresoValor(xtrdestino, xtipoTrEgresoAGenerar, reemite, ID_Flag_Anulado, False, "", False, False, False, False)
                        If Not xanulacionev Then
                            rta = MsgBox("No se pudo anular el Egreso de Valores" & Chr(13) & "¿Continúa de todas formas?", vbCritical + vbYesNo)
                            If rta = vbNo Then
                                MsgBox("Sen cancelará la anulación de la orden de pago")
                                If xws.InTransaction Then
                                    xws.RollBack()
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
End Function


Public Function AnulaEgresoValor(atransaccion ,  xtipotragenerar ,  reemite ,  ID_Flag_Anulado , GeneraCP , xcodigoTRDesimputacion , RemuevePendientes ,  RevierteAsiento , Cambianumero , RetornaTransaccion ) 
 Set xegreso = atransaccion
 If xegreso.estado = "C" Then
   If xegreso.EXTERNAL_ID = "" Then
      If GeneraCP Then
            If atransaccion.compromisopago Is Nothing Then
               MsgBox "La transaccion todavia no genero el compromiso de pago. Todavia no se puede anular"
               Exit Function
            Else
               If atransaccion.compromisopago.SALDO2_IMPORTE <> atransaccion.compromisopago.IMTOTAL2_IMPORTE Then
                              rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Proveedor.", 1, "Proceso de Anulación")
                              If rta = 2 Then
                                     Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                     Exit Function
                              End If
                              Set xcp = NewColumnSpec("COMPROMISOPAGO", "ID", "CP")
                              Set xORIGcp = NewColumnSpec("COMPROMISOPAGO", "TRORIGINANTE", "CP")
                              Set xview = NewCompoundView(atransaccion, "ITEMTRIMPUTACION", atransaccion.WorkSpace, nil, True)
                              Set xColumn = xview.ColumnFromPath("ORIGINANTE")
                              Set xjoin = NewJoinSpec(xColumn, xcp, False)
                              xview.addjoin (xjoin)
                              Set xFilter = NewFilterSpec(xORIGcp, "=", atransaccion.id)
                              xview.addfilter (xFilter)
              
                              Set xview2 = CNewCompoundView(atransaccion, "ITEMTRIMPUTACION", atransaccion.WorkSpace, nil, True)
                              Set xColumn = xview2.ColumnFromPath("DESTINATARIO")
                              Set xjoin = NewJoinSpec(xColumn, xcp, False)
                              xview2.addjoin (xjoin)
                              Set xFilter = NewFilterSpec(xORIGcp, "=", atransaccion.id)
                              xview2.addfilter (xFilter)
                              
                              Set xview.Union = xview2
                              xview.UnionAll = True
                              Set imputaciones = newcontainer()
                              Set d = CreateObject("Scripting.Dictionary")
                              For Each xitemimputacion In xview.viewitems
                                  Set xitemdes = ExisteBO(atransaccion, "ITEMTRDESIMPUTACION", "REFERENCIA", xitemimputacion.bo.placeowner.id, nil, True, False, "=")
                                  If xitemdes Is Nothing Then
                                     If Not d.Exists(xitemimputacion.bo.placeowner.id) Then
                                        imputaciones.Add (xitemimputacion.bo.placeowner)
                                        d.Add xitemimputacion.bo.placeowner.id, ""
                                     End If
                                  End If
                              Next
                              If imputaciones.Size > 0 Then
                                 Set xtrdesimputacion = CrearTransaccion(xcodigoTRDesimputacion, atransaccion.UnidadOperativa)
                                 xtrdesimputacion.destinatario = atransaccion.destinatario
                                 For Each xtrimputacion In imputaciones
                                     Set xitem = CrearItemTransaccion(xtrdesimputacion)
                                     Set xitem.referencia = xtrimputacion
                                 Next
                                 ShowBO xtrdesimputacion
                              Else
                                 MsgBox "No se encontro la imputacion correspondiente, verifique las imputaciones antes de continuar"
                                 Exit Function
                              End If
                     End If
               End If
         End If
         Set xTrAnulacion = CrearTransaccion(xtipotragenerar, xegreso.UnidadOperativa)
         Set xTrAnulacion.VINCULOTR = xegreso
         xID = xegreso.id
         mes = DatePart("m", Date)
         anio = DatePart("yyyy", Date)
         dia = DatePart("d", Date)
         fecha = anio & String(2 - Len(mes), "0") & mes & String(2 - Len(dia), "0") & dia
         xtabla = classname(xegreso)
         If Cambianumero Then
            xnumero = Right("9" & Left(atransaccion.numerodocumento, Len(atransaccion.numerodocumento) - 1), Len(atransaccion.numerodocumento))
            xsql = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xnumero & "' WHERE ID = '" & xegreso.id & "'"
         Else
            xsql = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "' WHERE ID = '" & xegreso.id & "'"
         End If
         Set xVector = NewVector()
         Set xBuc = NewBucket()
         xBuc.Value = xsql
         xVector.Add (xBuc)
         ExecutarSQL xVector, "DistrObj", False, xegreso.WorkSpace, -1
   '' ver que pasa con la tarjeta
         If classname(xegreso.originante) = "CHEQUERA" Or classname(xegreso.originante) = "CHEQUERAPAGODIFERIDO" Then
            Set xTrAnulacion.destinatario = xegreso.originante.CUENTABANCARIA
          Else
            Set xTrAnulacion.destinatario = xegreso.originante
         End If
         xTrAnulacion.nota = xegreso.nota
         xTrAnulacion.numerodocumento = xegreso.numerodocumento
         Set xTrAnulacion.originante = xegreso.destinatario
         xTrAnulacion.detalle = "Anulación de " & xegreso.nombre
         If xegreso.Attributes("centrocostos").isassigned Then
            Set xTrAnulacion.centrocostos = xegreso.centrocostos
         End If
         If GeneraCP Then
            Set  xTrAnulacion.boextension.compromisopagotranulada = xegreso.compromisopago
         End If
         If reemite Then
                xTrAnulacion.boextension.reemite_valor = True
         End If
         If RevierteAsiento Then
                Set xAsiento = GetAsiento(xTrAnulacion)
                If Not xAsiento Is Nothing Then
                         Set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, xTrAnulacion.fechaactual)
                End If
         End If
         xTrAnulacion.fechaactual = xegreso.fechaactual
         Set xTrAnulacion.total.unidadvalorizacion = xegreso.total.unidadvalorizacion
         xTrAnulacion.cotizacion = xegreso.cotizacion
         Set itemsxegreso = xegreso.itemstransaccion
         For Each elem In itemsxegreso
            Set xObjeto = CrearItemTransaccion(xTrAnulacion)
            xtipo = elem.referencia.tipovalor.referenciatipovalor.descripcion
            If xtipo = "EFECTIVO" Or xtipo = "CHEQUE PROPIO" Or xtipo = "CHEQUE DIFERIDO PROPIO" Or xtipo = "DOCUMENTO FINANCIERO" Or xtipo = "CHEQUE TERCERO" Or xtipo = "CHEQUE DIFERIDO TERCERO" Or xtipo = "TARJETA" Then
               Set xview = NewCompoundView(xegreso, "ITEMVALOR", xegreso.WorkSpace, elem.destino.valores, True)
               Set xColumn1 = xview.ColumnFromPath("ITEMTRANSACCION")
               Set xFilter = NewFilterSpec(xColumn1, " = ", elem.id)
               xview.addfilter (xFilter)
               For Each xitemvalor In xview.viewitems
                  If Not xitemvalor.bo.pasado Then
                     MsgBox "El valor todavia no esta conciliado. Debe Conciliarlo antes de realizar la anulación."
                     anulaEgresoValor = False
                   Else
                     Set xVALOR = xitemvalor.bo
                  End If
               Next
               Set xObjeto.origen = elem.destino
               If xtipo = "CHEQUE PROPIO" Or xtipo = "CHEQUE DIFERIDO PROPIO" Then
                  Set xObjeto.referencia = xVALOR
                ElseIf xtipo = "CHEQUE TERCERO" Or xtipo = "CHEQUE DIFERIDO TERCERO" Or xtipo = "TARJETA" Then
                  Set xObjeto.referencia = elem.referencia
                ElseIf xtipo = "EFECTIVO" Then
                  Set xview = NewCompoundView(xegreso, "ITEMVALOR", xegreso.WorkSpace, elem.destino.valores, True)
                  Set xColumn1 = xview.ColumnFromPath("TIPOVALOR")
                  Set xFilter = NewFilterSpec(xColumn1, " = ", elem.referencia.tipovalor.id)
                  xview.addfilter (xFilter)
				  Set xVALOR = Nothing
                  For Each xitemvalor In xview.viewitems
                     Set xVALOR = xitemvalor.bo
                  Next
				  If Not xVALOR Is Nothing Then
                     Set xObjeto.referencia = xVALOR
				  End If
                  Set xObjeto.valorori_unidadvalorizacion = elem.valorori_unidadvalorizacion
                  xObjeto.valorori_importe = elem.valorori_importe
                  xObjeto.valor.importe = elem.valor.importe
               End If
               If classname(elem.origen) = "CHEQUERA" Or classname(elem.origen) = "CHEQUERAPAGODIFERIDO" Then
                  Set xObjeto.destino = elem.origen.CUENTABANCARIA
                Else
                  Set xObjeto.destino = elem.origen
               End If
               
            End If
            Set xObjeto.imputacioncontable = elem.imputacioncontable
         Next
        ' PEGA LOS IMPUESTOS
         Set listaimpuestos = xegreso.IMPUESTOSTRANSACCION
         For Each impegreso In listaimpuestos
            impuesto = impegreso.definicionimpuesto.impuesto.nombre
            For Each ximp In xTrAnulacion.IMPUESTOSTRANSACCION
               If ximp.definicionimpuesto.impuesto.nombre = impuesto Then
                  ximp.valor.importe = impegreso.valor.importe
                  Exit For
               End If
            Next
         Next
'         For Each xFacturas In xegreso.itemscancelacionfinanzas
'            Set xObjeto = CrearBo("ItemCancelacionFinanzas", xegreso)
'            xegreso.itemscancelacionfinanzas.Add (xObjeto)
'            Set xObjeto.referencia = xFacturas.referencia
'            xObjeto.valor.importe = xFacturas.valor.importe
'           ' set xFactura = xItemCancFinan.REFERENCIA.Booriginante
'         Next
         
         If RemuevePendientes Then
             'Libero los pendientes que puedan haber generado la transaccion
                For Each xitemtransaccion In xegreso.itemscancelacionfinanzas
                    If Not xitemtransaccion.itempendiente Is Nothing Then
                       For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                           If xcancelacion.itemorigen.id = xitemtransaccion.id Then
                                xitemtransaccion.itempendiente.CANCELACIONES.Remove (xcancelacion)
                                xcancelacion.Delete
                           End If
                       Next
                    End If
                Next
         End If
    
         Call ShowBO(xTrAnulacion)
         If RetornaTransaccion Then
            Set anulaEgresoValor = xTrAnulacion
         Else
            anulaEgresoValor = True
         End If
   Else
      If xegreso.estado = "N" Then
        If RetornaTransaccion Then
            Set anulaEgresoValor = Nothing
        Else
            anulaEgresoValor = True
        End If
      Else
        MsgBox "La Emisión " & xegreso.numerodocumento & " ya fue Anulada anteriormante - " & xegreso.EXTERNAL_ID
        If RetornaTransaccion Then
            Set anulaEgresoValor = Nothing
        Else
            anulaEgresoValor = False
        End If
      End If
   End If
 Else
      If xegreso.estado = "N" Then
        If RetornaTransaccion Then
            Set anulaEgresoValor = Nothing
        Else
            anulaEgresoValor = True
        End If
      Else
        MsgBox "Para poder generar este Proceso la Emisión de Pago debe estar en estado Cerrada"
        If RetornaTransaccion Then
            Set anulaEgresoValor = Nothing
        Else
            anulaEgresoValor = False
        End If
	  End If
  End If
End Function


Public Function AnulaTRInventario(atransaccion, xtipotrgenerar , ID_Flag_Anulado, RemuevePendientes,RevierteAsiento)
		set anulaTRInventario = Nothing
        set atransaccion_anulada = Nothing
        If classname(atransaccion) = "TRINGRESOINVENTARIO" Or classname(atransaccion) = "TREGRESOINVENTARIO" Or classname(atransaccion) = "TRTRANSFERENCIAINVENTARIO" _
          Or classname(atransaccion) = "TRREPORTEPRODUCIDO" Or classname(atransaccion) = "TRREPORTECONSUMIDO" Then
            If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
                xImporte = 0
                xabortar = False
                fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

                '     atransaccion.external_id = "Anulada - " & fecha
                set xVector = NewVector()
                set xBucket = NewBucket()
                set atransaccion_anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
                atransaccion_anulada.NUMERODOCUMENTO = atransaccion.NUMERODOCUMENTO
                set atransaccion_anulada.VINCULOTR = atransaccion
                set atransaccion_anulada.Destinatario = atransaccion.Originante
                set atransaccion_anulada.Originante = atransaccion.Destinatario
                atransaccion_anulada.Detalle = "Anulación de " & atransaccion.nombre
                set atransaccion_anulada.centrocostos = atransaccion.centrocostos
                atransaccion_anulada.FECHAACTUAL = atransaccion.FECHAACTUAL
                atransaccion_anulada.NOTA = atransaccion.NOTA
                set atransaccion_anulada.mediotransporte = atransaccion.mediotransporte
                set atransaccion_anulada.domicilio = atransaccion.domicilio
                set atransaccion_anulada.Total.unidadvalorizacion = atransaccion.Total.unidadvalorizacion
                atransaccion_anulada.cotizacion = atransaccion.cotizacion
                set atransaccion_anulada.tipopago = atransaccion.tipopago
                set atransaccion_anulada.domicilioentrega = atransaccion.domicilioentrega
                set atransaccion_anulada.Ubicaciondes = atransaccion.Ubicacionori
                set atransaccion_anulada.Ubicacionori = atransaccion.Ubicaciondes
                xNumero = "9" & Right(atransaccion.NUMERODOCUMENTO, 11)
                xtabla = classname(atransaccion)
                If classname(atransaccion) = "TRREPORTEPRODUCIDO" Or classname(atransaccion) = "TRREPORTECONSUMIDO" Then
                    set atransaccion_anulada.PROCESSREPORT.WORKORDER = atransaccion.PROCESSREPORT.WORKORDER
                    set atransaccion_anulada.PROCESSREPORT.REQUEST = atransaccion.PROCESSREPORT.REQUEST
                    set atransaccion_anulada.PROCESSREPORT.PROCESS = atransaccion.PROCESSREPORT.PROCESS

                End If
                xBucket.Value = " UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "', EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xNumero & "' WHERE ID = '" & atransaccion.ID & "' "
                xVector.Add(xBucket)
                xx = ExecutarSQL(xVector, "distrobj", "Anulada - " & fecha, atransaccion.workspace, -1)

                For Each Item In atransaccion.ItemsTransaccion
                    set xObjeto = CrearItemTransaccion(atransaccion_anulada)
                    set xObjeto.ReferenciaTipo = Item.ReferenciaTipo
                    set xObjeto.lote = Item.lote
                    set xObjeto.serie = Item.serie
                    set xObjeto.despacho = Item.despacho
                    set xObjeto.depositodes = Item.depositoori
                    set xObjeto.depositoori = Item.depositodes
                    set xObjeto.Ubicaciondes = Item.Ubicacionori
                    set xObjeto.Ubicacionori = Item.Ubicaciondes
                    set xObjeto.unidadmedida = Item.unidadmedida
                    xObjeto.cantidad.cantidad = Item.cantidad.cantidad
                    set xObjeto.UNIDADMEDIDANOLINEAL = Item.UNIDADMEDIDANOLINEAL
                    xObjeto.cantidadnolineal.cantidad = Item.cantidadnolineal.cantidad
                    xObjeto.valor.Importe = Item.valor.Importe
                    xObjeto.bultos = Item.bultos
                    xObjeto.PORCENTAJEBONIFICACION = Item.PORCENTAJEBONIFICACION
                    set xObjeto.centrocostos = Item.centrocostos
                    set xObjeto.imputacioncontable = Item.imputacioncontable
                    xObjeto.Detalle = Item.Detalle
                    xObjeto.observacion.Memo = Item.observacion.Memo
                    For Each ximpuestoori In Item.IMPUESTOSITEMTRANSACCION
                        For Each ximpuestodest In xObjeto.IMPUESTOSITEMTRANSACCION
                            If ximpuestoori.definicionimpuesto.impuesto.ID = ximpuestodest.definicionimpuesto.impuesto.ID Then
                                ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                            End If
                        Next
                    Next
                    If classname(atransaccion) = "TRREPORTEPRODUCIDO" Or classname(atransaccion) = "TRREPORTECONSUMIDO" Then
                        'xObjeto.PROCESSREPORTITEM.PLANNEDRESULT = Item.PROCESSREPORTITEM.PLANNEDREQUEST
                        set xObjeto.PROCESSREPORTITEM.BASICPROCESS = Item.PROCESSREPORTITEM.BASICPROCESS
                    End If
                Next
                For Each ximpuestoori In atransaccion.IMPUESTOSTRANSACCION
                    For Each ximpuestodest In atransaccion_anulada.IMPUESTOSTRANSACCION
                        If ximpuestoori.definicionimpuesto.impuesto.nombre = ximpuestodest.definicionimpuesto.impuesto.nombre Then
                            ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                        End If
                    Next
                Next
                'saldo los pendientes que queden de esta transaccion
                For Each xItemTR In atransaccion.ItemsTransaccion
                    set xview = NewCompoundView(atransaccion, "ITEMTR", atransaccion.workspace, Nothing, True)
                    'Set xColumnaRelacion = xview.ColumnFromPath("RELACIONTRORIG")
                    'Set xFilter = CalipsoFunctions.NewFilterSpec(xColumnaRelacion, " = ", "{89C23657-3F01-11D5-86AD-0080AD403F5F}")
                    'xview.AddFilter (xFilter)
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", xItemTR))
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("CANCELADO"), " = ", False))
                    set xpend = Nothing
                    If Not xview.ViewItems.IsEmpty Then
                        For Each xpendiente In xview.ViewItems
                            set xpend = xpendiente.bo
                        Next
                        SaldarPendiente(xpend)
                    End If
                Next
                If RemuevePendientes Then
                    'Libero los pendientes que puedan haber generado la transaccion
                    For Each xitemtransaccion In atransaccion.ItemsTransaccion
                        If Not xitemtransaccion.itempendiente Is Nothing Then
                            For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                                If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                    xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                    xcancelacion.Delete()
                                End If
                            Next
                        End If
                    Next
                End If
                If RevierteAsiento Then
                    set xAsiento = GetAsiento(atransaccion)
                    If Not xAsiento Is Nothing Then
                        set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, atransaccion_anulada.fechaactual)
                    End If
                End If
                ShowBO(atransaccion_anulada)
                Set ReturnValue = atransaccion_anulada
            Else
                MsgBox("Esta Transaccion ya fue anulada")
            End If

        End If
End Function


Public Function AnulaIngresoValor(atransaccion, xtipotragenerar, ID_Flag_Anulado, RemuevePendientes , GeneraCP ,xcodigoTRDesimputacion , Cambianumero ) 
        Dim Calcontrol , CalipsoFunctions, xview , xtrdesimputacion, xingreso, xTrAnulacion , itemsxingreso , xObjeto 
        Dim rta 
        Dim xws 
        xws = atransaccion.workspace
       
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)

        'Instancio un CalCtrl y el objeto CalipsoFunctions tiene las funciones de Calipso
        Set anulaIngresoValor = Nothing
        set xingreso = atransaccion
        If xingreso.estado = "C" Then
            If xingreso.EXTERNAL_ID = "" Then
                If GeneraCP Then
                    If xingreso.compromisopago Is Nothing Then
                        MsgBox("La transaccion todavia no genero el compromiso de pago. Todavia no se puede anular")
                        Exit Function
                    Else
                        If xingreso.compromisopago.SALDO2_IMPORTE <> xingreso.compromisopago.IMTOTAL2_IMPORTE Then
                            If atransaccion.compromisopago.IMTOTAL2_IMPORTE > atransaccion.Compania.PARAMETROS.TOLERANCIACP Then
                                rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Cliente.", 1, "Proceso de Anulación")
                                If rta = 2 Then
                                    Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                    Exit Function
                                End If
                                set xtrdesimputacion = DesimputarTransaccion( atransaccion,  xcodigoTRDesimputacion)
                                If xtrdesimputacion Is Nothing Then
                                    Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
                set xTrAnulacion = CrearTransaccion(xtipotragenerar, xingreso.unidadOperativa)
                set xTrAnulacion.VINCULOTR = xingreso
                Dim xId, fecha, xtabla , xNumero , XSQL , xTipo 
                Dim xVector , xBuc, xValor 
                xId = xingreso.ID
                fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
                xtabla = classname(xingreso)
                If Cambianumero Then
                    xNumero = Right("9" & Right(xingreso.NUMERODOCUMENTO, Len(xingreso.NUMERODOCUMENTO) - 1), Len(xingreso.NUMERODOCUMENTO))
                    XSQL = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xNumero & "' WHERE ID = '" & xingreso.ID & "'"
                Else
                    XSQL = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "' WHERE ID = '" & xingreso.ID & "'"
                End If
                set xVector 	= NewVector()
                set xBuc 	= NewBucket()
                xBuc.Value = XSQL
                xVector.Add(xBuc)
                x = ExecutarSQL(xVector, "DistrObj", False, xws, -1)
                '' ver que pasa con la tarjeta

                set xTrAnulacion.Destinatario = xingreso.Originante
                set xTrAnulacion.Originante = xingreso.Destinatario
                xTrAnulacion.NOTA = xingreso.NOTA
                'xingreso.nota = "Anul"
                xTrAnulacion.NUMERODOCUMENTO = xingreso.NUMERODOCUMENTO
                xTrAnulacion.Detalle = "Anulación de " & xingreso.nombre
                If xingreso.Attributes("centrocostos").isassigned Then
                    xTrAnulacion.centrocostos = xingreso.centrocostos
                End If
                If GeneraCP Then
                    xTrAnulacion.BOExtension.compromisopagotranulada = xingreso.compromisopago
                End If
                xTrAnulacion.FECHAACTUAL = xingreso.FECHAACTUAL
                set xTrAnulacion.Total.unidadvalorizacion = xingreso.Total.unidadvalorizacion
                xTrAnulacion.cotizacion = xingreso.cotizacion
                set itemsxingreso = xingreso.ItemsTransaccion
                For Each elem In itemsxingreso
                    set xObjeto = CrearItemTransaccion(xTrAnulacion)
                    xTipo = elem.Referencia.tipovalor.referenciatipovalor.Descripcion
                    If xTipo = "EFECTIVO" Or xTipo = "DOCUMENTO FINANCIERO" Or xTipo = "CHEQUE TERCERO" Or xTipo = "CHEQUE DIFERIDO TERCERO" Or xTipo = "TARJETA" Then
                        set xview = NewCompoundView(xingreso, "ITEMVALOR", xws, elem.destino.Valores, True)
                        xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", CStr(elem.ID)))
                        For Each xitemvalor In xview.ViewItems
                            If Not xitemvalor.bo.pasado Then
                                set xValor = xitemvalor.bo
                            Else
                                MsgBox("El valor todavia no esta conciliado. Debe Conciliarlo antes de realizar la anulación.")
                                set anulaIngresoValor = Nothing
                            End If
                        Next
                        SendDebug("Despues del for each")
                        set xObjeto.origen = elem.destino
                        set xValor = Nothing
                        If xTipo = "CHEQUE PROPIO" Or xTipo = "CHEQUE DIFERIDO PROPIO" Then
                            If elem.Referencia.pasado Then

                                set xview = NewCompoundView(xingreso, "ITEMVALOR", xws, elem.destino.Valores, True)
                                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("TIPOVALOR"), " = ", CStr(elem.Referencia.tipovalor.CONSOLIDATIPOVALORCA.ID)))
                                For Each xitemvalor In xview.ViewItems
                                    xValor = xitemvalor.bo
                                Next
                                set xObjeto.Referencia = xValor
                                set xObjeto.valorori_unidadvalorizacion = elem.valorori_unidadvalorizacion
                                xObjeto.valorori_importe = elem.valorori_importe
                                xObjeto.valor.Importe = elem.valor.Importe
                            Else
                                set xObjeto.Referencia = xValor
                            End If
                        ElseIf xTipo = "CHEQUE TERCERO" Or xTipo = "CHEQUE DIFERIDO TERCERO" Or xTipo = "TARJETA" Then
                            If elem.Referencia.pasado Then
                                SendDebug("Antes de la vista")
                                set xview = NewCompoundView(xingreso, "ITEMVALOR", xws, elem.destino.Valores, True)
                                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("TIPOVALOR"), " = ", CStr(elem.Referencia.tipovalor.CONSOLIDATIPOVALORCA.ID)))
                                For Each xitemvalor In xview.ViewItems
                                    set xValor = xitemvalor.bo
                                Next
                                set xObjeto.Referencia = xValor
                                set xObjeto.valorori_unidadvalorizacion = elem.valorori_unidadvalorizacion
                                xObjeto.valorori_importe = elem.valorori_importe
                                xObjeto.valor.Importe = elem.valor.Importe
                                SendDebug("Despues de la vista")
                            Else
                                set xObjeto.Referencia = elem.Referencia
                            End If

                        ElseIf xTipo = "EFECTIVO" Then
                            set xview = NewCompoundView(xingreso, "ITEMVALOR", xws, elem.destino.Valores, True)
                            xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("TIPOVALOR"), " = ", CStr(elem.Referencia.tipovalor.CONSOLIDATIPOVALORCA.ID)))
                            For Each xitemvalor In xview.ViewItems
                                set xValor = xitemvalor.bo
                            Next
                            set xObjeto.Referencia = xValor
                            set xObjeto.valorori_unidadvalorizacion = elem.valorori_unidadvalorizacion
                            xObjeto.valorori_importe = elem.valorori_importe
                            xObjeto.valor.Importe = elem.valor.Importe
                        End If
                        If Not elem.origen Is Nothing Then
                            set xObjeto.destino = elem.origen
                        End If
                    End If
                    set xObjeto.imputacioncontable = elem.imputacioncontable
                Next
                ' PEGA LOS IMPUESTOS
                SendDebug("Antes de impuestos")
                For Each impegreso In xingreso.IMPUESTOSTRANSACCION
                    For Each ximp In xTrAnulacion.IMPUESTOSTRANSACCION
                        If ximp.definicionimpuesto.impuesto.id = impegreso.definicionimpuesto.impuesto.id Then
                            ximp.valor.Importe = impegreso.valor.Importe
                            Exit For
                        End If
                    Next
                Next
                SendDebug("Antes de itemscancelacion")
                For Each xFacturas In xingreso.itemscancelacionfinanzas
                    set xObjeto = crearbo("ItemCancelacionFinanzas", xingreso)
                    xTrAnulacion.itemscancelacionfinanzas.Add(xObjeto)
                    set xObjeto.Referencia = xFacturas.Referencia
                    xObjeto.valor.Importe = xFacturas.valor.Importe
                    ' set xFactura = xItemCancFinan.REFERENCIA.Booriginante
                Next
                SendDebug("Antes Remueve")
                If RemuevePendientes Then
                    'Libero los pendientes que puedan haber generado la transaccion
                    For Each xitemtransaccion In atransaccion.itemscancelacionfinanzas
                        If Not xitemtransaccion.itempendiente Is Nothing Then
                            For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                                If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                    xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                    xcancelacion.Delete()
                                End If
                            Next
                        End If
                    Next
                End If
                SendDebug("Antes de showbo")
                Call ShowBO(xTrAnulacion)
                set anulaIngresoValor = xTrAnulacion

            Else
                MsgBox("La Emisión " & xingreso.NUMERODOCUMENTO & " ya fue Anulada anteriormante - " & xingreso.EXTERNAL_ID)
                set anulaIngresoValor = Nothing
            End If
        Else
            MsgBox("Para poder generar este Proceso la Emisión de Pago debe estar en estado Cerrada")
            set anulaIngresoValor = Nothing
        End If
End Function


Public Function anulaTRVentas(atransaccion  , xtipotrgenerar  ,   xcodigoTRDesimputacion  ,   ID_Flag_Anulado  ,   RemuevePendientes  , RevierteAsiento  )  
        Dim xAsiento  , xview , xview2 , _
        xtrdesimputacion  , xFactura  , xObjeto  
        Dim xImporte, xabortar  , rta, xmensaje  
        Dim xtabla  , fecha  
        Dim xVector  , xBucket  , xfactura_anulada  , xpend  , xNuevoAsiento  

        set xws = atransaccion.workspace

        Dim CalipsoFunctions  
     
        CalipsoFunctions = GetCalipsoFuncs(atransaccion)
        set anulaTRVentas = Nothing

    
        If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
            xImporte = 0
            xabortar = False
            If atransaccion.compromisopago.SALDO2_IMPORTE <> atransaccion.compromisopago.IMTOTAL2_IMPORTE Then
                If atransaccion.compromisopago.IMTOTAL2_IMPORTE > atransaccion.Compania.PARAMETROS.TOLERANCIACP Then
                    If classname(atransaccion.compromisopago) = "CPDEBITO" Or classname(atransaccion.compromisopago) = "CPFACTURA" Then
                        set xview2 = NewCompoundView(atransaccion, "ITEMORDENPAGO", xws, Nothing, True)
                        xview2.AddFilter(NewFilterSpec(xview2.ColumnFromPath("REFERENCIA"), "=", atransaccion.compromisopago))
                        xmensaje = ""
                        For Each xop In xview2.ViewItems
                            xmensaje = xmensaje & xop.bo.placeowner.nombre & Chr(13)
                            xabortar = True
                        Next
                        If xabortar Then
                            Call MsgBox("Esta Transaccion ya ha sido pagada total o parcialmente." & Chr(13) & "Debe anular o desimputar primero todos los pagos imputados a esta factura." & Chr(13) & "Transacciones Imputadas:" & Chr(13) & xmensaje, 0, "Proceso de Anulacion")
                            Exit Function
                        End If
                    End If
                    'si llego hasta aca significa que fue imputada a un credito.

                    rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Proveedor.", 1, "Proceso de Anulación")
                    If rta = 2 Then
                        Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                        Exit Function
                    End If
                    set xtrdesimputacion = DesimputarTransaccion(atransaccion, xcodigoTRDesimputacion)
					
                    If xtrdesimputacion Is Nothing Then
                        Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                        Exit Function
                    End If
                End If
            End If
            set xFactura = atransaccion

            fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
            '     xfactura.external_id = "Anulada - " & fecha
            set xVector = NewVector()
            set xBucket = NewBucket()
            set xfactura_anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
            xfactura_anulada.NUMERODOCUMENTO 	= xFactura.NUMERODOCUMENTO
            set xfactura_anulada.VINCULOTR 			= xFactura
            set xfactura_anulada.BOExtension.compromisopagotranulada = xFactura.compromisopago
            set xfactura_anulada.Destinatario 		= xFactura.Destinatario
            set xfactura_anulada.Originante 		= xFactura.Originante
            xfactura_anulada.Detalle 			= "Anulación de " & xFactura.nombre
            set xfactura_anulada.centrocostos 		= xFactura.centrocostos
            xfactura_anulada.FECHAACTUAL 		= xFactura.FECHAACTUAL
            xfactura_anulada.NOTA = xFactura.NOTA
            set xfactura_anulada.mediotransporte = xFactura.mediotransporte
            set xfactura_anulada.domicilio = xFactura.domicilio
            xfactura_anulada.Total.unidadvalorizacion = xFactura.Total.unidadvalorizacion
            xfactura_anulada.cotizacion = xFactura.cotizacion
            set xfactura_anulada.tipopago = xFactura.tipopago
            set xfactura_anulada.domicilioentrega = xFactura.domicilioentrega
            If classname(atransaccion) = "TRDEBITOCOMPRA" Or _
              classname(atransaccion) = "TRCREDITOCOMPRA" Or _
              classname(atransaccion) = "TRDEBITOVENTA" Or _
              classname(atransaccion) = "TRCREDITOVENTA" Then
                xfactura_anulada.GENERADAPORDIFCAMBIO = xFactura.GENERADAPORDIFCAMBIO
            End If
            'xnumero = "9" & Right(atransaccion.numerodocumento, 11)
            xtabla = classname(atransaccion)
            xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "' WHERE ID = '" & atransaccion.ID & "'"
            xVector.Add(xBucket)
            x = ExecutarSQL(xVector, "distrobj", "Anulada - " & fecha, xws, -1)

            For Each Item In xFactura.ItemsTransaccion
                set xObjeto = CrearItemTransaccion(xfactura_anulada)
                set xObjeto.Referencia = Item.Referencia
                set xObjeto.unidadmedida = Item.unidadmedida
                xObjeto.cantidad.cantidad = Item.cantidad.cantidad
                set xObjeto.UNIDADMEDIDANOLINEAL = Item.UNIDADMEDIDANOLINEAL
                xObjeto.cantidadnolineal.cantidad = Item.cantidadnolineal.cantidad
                xObjeto.valor.Importe = Item.valor.Importe
                xObjeto.bultos = Item.bultos
                xObjeto.PORCENTAJEBONIFICACION = Item.PORCENTAJEBONIFICACION
                set xObjeto.centrocostos = Item.centrocostos
                set xObjeto.imputacioncontable = Item.imputacioncontable
                xObjeto.Detalle = Item.Detalle
                xObjeto.observacion.Memo = Item.observacion.Memo
                For Each ximpuestoori In Item.IMPUESTOSITEMTRANSACCION
                    For Each ximpuestodest In xObjeto.IMPUESTOSITEMTRANSACCION
                        If ximpuestoori.definicionimpuesto.impuesto.ID = ximpuestodest.definicionimpuesto.impuesto.ID Then
                            ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                        End If
                    Next
                Next
             '   If EjecutaFuncionUsuarioItem Then
              '      AnulacionItemTRVenta(Item, xObjeto)
              '  End If
            Next
            For Each ximpuestoori In xFactura.IMPUESTOSTRANSACCION
                For Each ximpuestodest In xfactura_anulada.IMPUESTOSTRANSACCION
                    If ximpuestoori.definicionimpuesto.impuesto.nombre = ximpuestodest.definicionimpuesto.impuesto.nombre Then
                        ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                    End If
                Next
            Next
            xpend = Nothing
            'saldo los pendientes que queden de esta transaccion
            For Each xItemTR In atransaccion.ItemsTransaccion
                set xview = NewCompoundView(xFactura, "ITEMTR", xws, Nothing, True)
                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", xItemTR))
                xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("CANCELADO"), " = ", False))
                If Not xview.ViewItems.IsEmpty Then
                    For Each xpendiente In xview.ViewItems
                        set xpend = xpendiente.bo
                    Next
                    SaldarPendiente(xpend)
                End If
            Next

            If RemuevePendientes Then
                'Libero los pendientes que puedan haber generado la transaccion
                For Each xitemtransaccion In atransaccion.ItemsTransaccion
                    If Not xitemtransaccion.itempendiente Is Nothing Then
                        For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                            If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                xcancelacion.Delete()
                            End If
                        Next
                    End If
                Next
            End If
            If RevierteAsiento Then
                set xAsiento = GetAsiento(atransaccion)
                If Not xAsiento Is Nothing Then
                    set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, xfactura_anulada.FECHAACTUAL)

                End If
            End If
            ShowBO(xfactura_anulada)
            set anulaTRVentas = xfactura_anulada
        Else

            MsgBox("Esta Transaccion ya fue anulada")
        End If
        'End If

End Function


Public Function anulaParteServicios( atransaccion  ,   xtipotrgenerar  ,   ID_Flag_Anulado  ,   RemuevePendientes  ,   SinCambioNumero  ,     RevierteAsiento  )  
        Dim xAsiento  , xFactura  , xObjeto  
        Dim xImporte, xabortar  
        Dim xtabla  , fecha  , xNumero  
        Dim xVector  , xBucket  , xfactura_anulada  , xNuevoAsiento  
		Dim xws 
        set xws = atransaccion.workspace

        Dim CalipsoFunctions  
       
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)
        set anulaParteServicios = Nothing

        If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
            xImporte = 0
            xabortar = False
            set xFactura = atransaccion
            fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
            set xVector = NewVector()
            set xBucket = NewBucket()
            set xfactura_anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
            xfactura_anulada.NUMERODOCUMENTO = xFactura.NUMERODOCUMENTO
            set xfactura_anulada.VINCULOTR = xFactura
            set xfactura_anulada.Destinatario = xFactura.Destinatario
            set xfactura_anulada.Originante = xFactura.Originante
            xfactura_anulada.Detalle = "Anulación de " & xFactura.nombre
            set xfactura_anulada.centrocostos = xFactura.centrocostos
            xfactura_anulada.FECHAACTUAL = xFactura.FECHAACTUAL
            xfactura_anulada.NOTA = xFactura.NOTA
            set xfactura_anulada.mediotransporte = xFactura.mediotransporte
            set xfactura_anulada.Total.unidadvalorizacion = xFactura.Total.unidadvalorizacion
            xfactura_anulada.cotizacion = xFactura.cotizacion
            xNumero = "9" & Right(atransaccion.NUMERODOCUMENTO, 11)
            xtabla = classname(atransaccion)
            xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xNumero & "' WHERE ID = '" & atransaccion.ID & "'"
            If SinCambioNumero Then
                xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "' WHERE ID = '" & atransaccion.ID & "'"
            End If
            xVector.Add(xBucket)
            x = ExecutarSQL(xVector, "distrobj", "Anulada - " & fecha, xws, -1)

            For Each Item In xFactura.ItemsTransaccion
                set xObjeto = CrearItemTransaccion(xfactura_anulada)
                set xObjeto.Referencia = Item.Referencia
                set xObjeto.unidadmedida = Item.unidadmedida
                xObjeto.cantidad.cantidad = Item.cantidad.cantidad
                set xObjeto.UNIDADMEDIDANOLINEAL = Item.UNIDADMEDIDANOLINEAL
                xObjeto.cantidadnolineal.cantidad = Item.cantidadnolineal.cantidad
                xObjeto.valor.Importe = Item.valor.Importe
                xObjeto.PORCENTAJEBONIFICACION = Item.PORCENTAJEBONIFICACION
                set xObjeto.centrocostos = Item.centrocostos
                set xObjeto.imputacioncontable = Item.imputacioncontable
                xObjeto.Detalle = Item.Detalle
                xObjeto.observacion.Memo = Item.observacion.Memo
                set xObjeto.RecursoDeUso = Item.RecursoDeUso
                set xObjeto.ItemRecursoUso = Item.ItemRecursoUso
                For Each ximpuestoori In Item.IMPUESTOSITEMTRANSACCION
                    For Each ximpuestodest In xObjeto.IMPUESTOSITEMTRANSACCION
                        If ximpuestoori.definicionimpuesto.impuesto.ID = ximpuestodest.definicionimpuesto.impuesto.ID Then
                            ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                        End If
                    Next
                Next
                If Not Item.PROCESSREPORTITEM Is Nothing Then
                    set xObjeto.PROCESSREPORTITEM.PLANNEDREQUEST = Item.PROCESSREPORTITEM.PLANNEDREQUEST
                    set xObjeto.PROCESSREPORTITEM.BASICPROCESS = Item.PROCESSREPORTITEM.BASICPROCESS
                End If
            Next
            For Each ximpuestoori In xFactura.IMPUESTOSTRANSACCION
                For Each ximpuestodest In xfactura_anulada.IMPUESTOSTRANSACCION
                    If ximpuestoori.definicionimpuesto.impuesto.nombre = ximpuestodest.definicionimpuesto.impuesto.nombre Then
                        ximpuestodest.valor.Importe = ximpuestoori.valor.Importe
                    End If
                Next
            Next
            If RemuevePendientes Then
                'Libero los pendientes que puedan haber generado la transaccion
                For Each xitemtransaccion In atransaccion.ItemsTransaccion
                    If Not xitemtransaccion.itempendiente Is Nothing Then
                        For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                            If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                xcancelacion.Delete()
                            End If
                        Next
                    End If
                Next
            End If
            If RevierteAsiento Then
                set xAsiento = GetAsiento(atransaccion)
                If Not xAsiento Is Nothing Then
                    set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, xfactura_anulada.fecharegistro)
                End If
            End If
           
            set anulaParteServicios = xfactura_anulada
        Else

            MsgBox("Esta Transaccion ya fue anulada")
        End If
End Function


Public Function anulaTRDelgadaOrden( atransaccion  ,   xtipotrgenerar  ,   xcodigoTRDesimputacion  ,   ID_Flag_Anulado  ,   RemuevePendientes  ,     RevierteAsiento   ,     PrefijoNumero   ,     Trinventario  )  
        Dim xAsiento  , xview , xview2 , xtrdesimputacion  , xFactura  , xObjeto  
        Dim xImporte, xabortar  , rta , xmensaje  
        Dim xtabla  , fecha  , xNumero  
        Dim xVector  , xBucket  , xfactura_anulada  , xpend  , xNuevoAsiento  
        Dim xws 
        set xws = atransaccion.workspace

        Dim CalipsoFunctions  
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)
        set anulaTRDelgadaOrden = Nothing
        senddebug("Inicio Función Anula Tr")

        If atransaccion.tipotransaccion.OTP = "TDORDEN" Then
            If atransaccion.estado = "C" And atransaccion.EXTERNAL_ID = "" Then
                xImporte = 0
                xabortar = False
                If Not atransaccion.compromisopago Is Nothing Then
                    If atransaccion.compromisopago.SALDO2_IMPORTE <> atransaccion.compromisopago.IMTOTAL2_IMPORTE Then
                        If atransaccion.compromisopago.IMTOTAL2_IMPORTE > atransaccion.Compania.PARAMETROS.TOLERANCIACP Then
                            If classname(atransaccion.compromisopago) = "CPDEBITO" Or classname(atransaccion.compromisopago) = "CPFACTURA" Then
                                set xview2 = NewCompoundView(atransaccion, "ITEMORDENPAGO", xws, Nothing, True)
                                xview2.AddFilter(NewFilterSpec(xview2.ColumnFromPath("REFERENCIA"), "=", atransaccion.compromisopago))
                                xmensaje = ""
                                For Each xop In xview2.ViewItems
                                    xmensaje = xmensaje & xop.bo.placeowner.nombre & Chr(13)
                                    xabortar = True
                                Next
                                If xabortar Then
                                    Call MsgBox("Esta Transaccion ya ha sido pagada total o parcialmente." & Chr(13) & "Debe anular o desimputar primero todos los pagos imputados a esta factura." & Chr(13) & "Transacciones Imputadas:" & Chr(13) & xmensaje, 0, "Proceso de Anulacion")
                                    Exit Function
                                End If
                            End If
                            'si llego hasta aca significa que fue imputada a un credito.

                            rta = MsgBox("Esta transacción ya esta imputada, debe proceder a desimputarla." & Chr(13) & "Si tiene dudas consulte la cuenta corriente del Proveedor.", 1, "Proceso de Anulación")
                            If rta = vbCancel Then
                                Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                Exit Function
                            End If
                            set xtrdesimputacion = DesimputarTransaccion(atransaccion, xcodigoTRDesimputacion)
                            If xtrdesimputacion Is Nothing Then
                                Call MsgBox("Proceso Cancelado", 0, "Proceso de Anulacion")
                                Exit Function
                            End If
                        End If
                    End If
                End If
                set xFactura = atransaccion
                fecha = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
                '     xfactura.external_id = "Anulada - " & fecha
                set xVector = NewVector()
                set xBucket = NewBucket()
                set xfactura_anulada = CrearTransaccion(xtipotrgenerar, atransaccion.unidadOperativa)
                xfactura_anulada.NUMERODOCUMENTO = xFactura.NUMERODOCUMENTO
                set xfactura_anulada.VINCULOTR = xFactura
                If xFactura.tipotransaccion.CONFIGCP.GeneraCP Then
                    If Not xfactura_anulada.BOExtension Is Nothing Then
                        set xfactura_anulada.BOExtension.compromisopagotranulada = xFactura.compromisopago
                    End If
                End If
                If Trinventario Then
                    SendDebug("Seteo deposito de la cabecera")
                    set xfactura_anulada.Destinatario = xFactura.Originante
                    set xfactura_anulada.Originante = xFactura.Destinatario
                    set xfactura_anulada.Ubicaciondes = xFactura.Ubicacionori
                    set xfactura_anulada.Ubicacionori = xFactura.Ubicaciondes

                Else
                    If classname(xFactura.Destinatario) = "EMPLEADO" Then
                      On Error Resume Next
                            set xfactura_anulada.Destinatario = xFactura.Destinatario
                      Else
                        set xfactura_anulada.Destinatario = xFactura.Destinatario
                    End If
                    set xfactura_anulada.Originante = xFactura.Originante
                End If

                xfactura_anulada.Detalle = "Anulación de " & xFactura.nombre
                set xfactura_anulada.centrocostos = xFactura.centrocostos
                xfactura_anulada.FECHAACTUAL = xFactura.FECHAACTUAL
                xfactura_anulada.NOTA = xFactura.NOTA
                set xfactura_anulada.mediotransporte = xFactura.mediotransporte
                set xfactura_anulada.domicilio = xFactura.domicilio
                set xfactura_anulada.total_unidadvalorizacion = xFactura.total_unidadvalorizacion
                xfactura_anulada.cotizacion = xFactura.cotizacion
                set xfactura_anulada.tipopago = xFactura.tipopago
                set xfactura_anulada.domicilioentrega = xFactura.domicilioentrega
                xfactura_anulada.fecharegistro = xFactura.fecharegistro
                'If classname(atransaccion) = "TRDEBITOCOMPRA" Or _
                '   classname(atransaccion) = "TRCREDITOCOMPRA" Or _
                '   classname(atransaccion) = "TRDEBITOVENTA" Or _
                '   classname(atransaccion) = "TRCREDITOVENTA" Then
                '             xfactura_anulada.GENERADAPORDIFCAMBIO = xfactura.GENERADAPORDIFCAMBIO
                'End If
                If PrefijoNumero = "" Then
                    xNumero = "9" & Right(atransaccion.NUMERODOCUMENTO, 11)
                Else
                    xNumero = PrefijoNumero & Right(atransaccion.NUMERODOCUMENTO, Len(atransaccion.NUMERODOCUMENTO) - Len(PrefijoNumero))
                End If
                xtabla = classname(atransaccion)
                xBucket.Value = "UPDATE " & xtabla & " SET FLAG_ID = '" & ID_Flag_Anulado & "',EXTERNAL_ID = 'Anulada - " & Year(Now()) & Right("00" & Month(Now()), 2) & Right("00" & Day(Now()), 2) & "', NUMERODOCUMENTO = '" & xNumero & "' WHERE ID = '" & atransaccion.ID & "'"
                xVector.Add(xBucket)
                x = ExecutarSQL(xVector, "distrobj", "Anulada - " & fecha, xws, -1)

                For Each Item In xFactura.ItemsTransaccion
                    set xObjeto = CrearItemTransaccion(xfactura_anulada)
                    If classname(Item.Referencia) = "PRODUCTO" Or classname(Item.Referencia) = "CONCEPTOCONTABLE" Or classname(Item.Referencia) = "SERVICIO" Then
                        set xObjeto.Referencia = Item.Referencia
                    Else
                        If Not Item.ReferenciaTipo Is Nothing Then
                            set xObjeto.ReferenciaTipo = Item.ReferenciaTipo
                        End If
                    End If

                    xObjeto.PORCENTAJEBONIFICACION = Item.PORCENTAJEBONIFICACION
                    set xObjeto.centrocostos = Item.centrocostos
                    set xObjeto.imputacioncontable = Item.imputacioncontable
                    set xObjeto.CUENTACONTABLE = Item.CUENTACONTABLE
                    xObjeto.Detalle = Item.Detalle
                    xObjeto.observacion.Memo = Item.observacion.Memo
                    If Trinventario Then
                        set xObjeto.depositodes = Item.depositoori
                        set xObjeto.Ubicaciondes = Item.Ubicacionori
                        set xObjeto.depositoori = Item.depositodes
                        set xObjeto.Ubicacionori = Item.Ubicaciondes
                        set xObjeto.serie = Item.serie
                        set xObjeto.lote = Item.lote
                        set xObjeto.propietario = Item.propietario
                        SendDebug("Seteo serie y lote")
                    End If
                    set xObjeto.unidadmedida = Item.unidadmedida
                    xObjeto.Cantidad_Cantidad = Item.Cantidad_Cantidad
                    set xObjeto.UNIDADMEDIDANOLINEAL = Item.UNIDADMEDIDANOLINEAL
                    xObjeto.cantidadnl_cantidad = Item.cantidadnl_cantidad
                    xObjeto.Valor_Importe = Item.Valor_Importe
                    xObjeto.bultos = Item.bultos
                    For Each ximpuestoori In xFactura.IMPUESTOS
                        If ximpuestoori.boowner.ID = Item.ID Then
                            For Each ximpuestodest In xfactura_anulada.IMPUESTOS
                                If ximpuestodest.boowner.ID = xObjeto.ID Then
                                    If ximpuestoori.definicionimpuesto.impuesto.ID = ximpuestodest.definicionimpuesto.impuesto.ID Then
                                        ximpuestodest.Importe = ximpuestoori.Importe
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
                For Each ximpuestoori In xFactura.IMPUESTOS
                    If ximpuestoori.boowner.ID = xFactura.ID Then
                        For Each ximpuestodest In xfactura_anulada.IMPUESTOS
                            If ximpuestodest.boowner.ID = xfactura_anulada.ID Then
                                If ximpuestoori.definicionimpuesto.impuesto.nombre = ximpuestodest.definicionimpuesto.impuesto.nombre Then
                                    ximpuestodest.Importe = ximpuestoori.Importe
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                'saldo los pendientes que queden de esta transaccion
                xpend = Nothing
                For Each xItemTR In xFactura.ItemsTransaccion
                    set xview = NewCompoundView(xFactura, "ITEMTR", xws, Nothing, True)
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("ITEMTRANSACCION"), " = ", xItemTR))
                    xview.AddFilter(NewFilterSpec(xview.ColumnFromPath("CANCELADO"), " = ", False))
                    If Not xview.ViewItems.IsEmpty Then
                        For Each xpendiente In xview.ViewItems
                            set xpend = xpendiente.bo
                        Next
                        SaldarPendiente(xpend)
                    End If
                Next
                If RemuevePendientes Then
                    'Libero los pendientes que puedan haber generado la transaccion
                    For Each xitemtransaccion In atransaccion.ItemsTransaccion
                        If Not xitemtransaccion.itempendiente Is Nothing Then
                            For Each xcancelacion In xitemtransaccion.itempendiente.CANCELACIONES
                                If xcancelacion.itemorigen.ID = xitemtransaccion.ID Then
                                    xitemtransaccion.itempendiente.CANCELACIONES.Remove(xcancelacion)
                                    xcancelacion.Delete()
                                End If
                            Next
                        End If
                    Next
                End If
                If RevierteAsiento Then
                    set xAsiento = GetAsiento(atransaccion)
                    If Not xAsiento Is Nothing Then
                        set xNuevoAsiento = ReversarAsiento(xAsiento, xAsiento.Compania.CONFIGURADORCONTABLE.EJERCICIOCORRIENTE, xfactura_anulada.fecharegistro)
                    End If
                End If
                ShowBO(xfactura_anulada)
                set anulaTRDelgadaOrden = xfactura_anulada
            Else
                If atransaccion.estado = "A" Then
                    atransaccion.estado = "N"
                ElseIf atransaccion.estado = "P" Then
                    MsgBox("Esta Transaccion esta en estado a Procesar, debes anularla una vez que ya fue cerrada.")
                Else
                    MsgBox("Esta Transaccion ya fue anulada")
                End If
            End If

        End If

End Function

	
	
'' Funciones Varias

Public Function GetAsiento(atransaccion) 
		set GetAsiento = Nothing
        Dim xAsiento, xview, xTrProceso, xTrProceso_Proceso 
        Dim CalipsoFunctions, xws
        
        set CalipsoFunctions = GetCalipsoFuncs(atransaccion)
        set xws = atransaccion.workspace

        set xTrProceso = NewColumnSpec("TRPROCESOPORLOTE", "TRANSACCION", "")
        set xTrProceso_Proceso = NewColumnSpec("TRPROCESOPORLOTE", "PROCESOPORLOTE", "")
        set xview = NewCompoundView(atransaccion, "TRCONTABLE", xws, Nothing, True)
        xview.addjoin(NewJoinSpec(xview.ColumnFromPath("GENERADAPOR"), xTrProceso_Proceso, False))
        xview.addfilter(NewFilterSpec(xTrProceso, "=", atransaccion.ID))
        set xAsiento = Nothing
        For Each xitem In xview.viewitems
            set xAsiento = xitem.bo
        Next
        set GetAsiento = xAsiento
End Function

Public Function ReversarAsiento(aAsiento , aEjercicio , FechaAplicacion )
        Dim xNuevoAsiento, xitem , xPeriodo , xViewPeriodos, xejercicio 
        Dim xNuevoitem
        Dim CalipsoFunctions 
        set CalipsoFunctions = GetCalipsoFuncs(aAsiento)
        Dim xws 
        set xws = aAsiento.workspace
		set ReversarAsiento = Nothing
   
        set xPeriodo = Nothing
        set xViewPeriodos = NewCompoundView(aAsiento, "PERIODO", xws, Nothing, False)

        Call xViewPeriodos.addfilter(NewFilterSpec(xViewPeriodos.ColumnFromPath("DESDEFECHA"), "<=", FechaAplicacion))
        Call xViewPeriodos.addfilter(NewFilterSpec(xViewPeriodos.ColumnFromPath("HASTAFECHA"), ">=", FechaAplicacion))

        For Each xItemViewPeriodos In xViewPeriodos.viewitems
            set xPeriodo = xItemViewPeriodos.bo
            Exit For
        Next

        set xejercicio = xPeriodo.Place.Owner

        set xNuevoAsiento = CrearTRContable(aAsiento.tipotransaccion.codigo, xejercicio.codigo, FechaAplicacion, aAsiento.unidadOperativa)
        Call nomensaje(xNuevoAsiento, True)
        xNuevoAsiento.Detalle = "Reverso " & aAsiento.nombre
        xNuevoAsiento.cotizacion = aAsiento.cotizacion
        xNuevoAsiento.Ajuste = aAsiento.Ajuste
        xNuevoAsiento.Subdiario = aAsiento.Subdiario
        For Each xitem In aAsiento.ItemsTransaccion
            xNuevoitem = CrearItemTransaccion(xNuevoAsiento)
            xNuevoitem.Referencia = xitem.Referencia
            xNuevoitem.debeoriginal.unidadvalorizacion = xitem.haberoriginal.unidadvalorizacion
            xNuevoitem.debeoriginal.Importe = xitem.haberoriginal.Importe
            xNuevoitem.haberoriginal.unidadvalorizacion = xitem.debeoriginal.unidadvalorizacion
            xNuevoitem.haberoriginal.Importe = xitem.debeoriginal.Importe
            xNuevoitem.centrocostos = xitem.centrocostos
            xNuevoitem.Descripcion = xitem.Descripcion
            xNuevoitem.Detalle = xitem.Detalle
        Next
        set ReversarAsiento = xNuevoAsiento
End Function


Function DesimputarTransaccion(  atransaccion  ,  xcodigoTRDesimputacion)  
        set DesimputarTransaccion = Nothing
		set xws = atransaccion.workspace
        Dim xCp  , xOrigCp  , xview , xView2 , imputaciones
        Dim d  , xitemdes  , xtrdesimputacion  , xitem  
        set xCp = NewColumnSpec("COMPROMISOPAGO", "ID", "CP")
        set xOrigCp = NewColumnSpec("COMPROMISOPAGO", "TRORIGINANTE", "CP")
        set xview = NewCompoundView(atransaccion, "ITEMTRIMPUTACION", xws, Nothing, True)
        xview.AddJoin(NewJoinSpec(xview.ColumnFromPath("ORIGINANTE"), xCp, False))
        xview.AddFilter(NewFilterSpec(xOrigCp, "=", atransaccion))

        set xView2 = NewCompoundView(atransaccion, "ITEMTRIMPUTACION", xws, Nothing, True)
        xView2.AddJoin(NewJoinSpec(xView2.ColumnFromPath("DESTINATARIO"), xCp, False))
        xView2.AddFilter(NewFilterSpec(xOrigCp, "=", atransaccion))
        xview.Union = xView2
        xview.UnionAll = True
        imputaciones = newcontainer()
        set d = CreateObject("Scripting.Dictionary")
        For Each xitemimputacion In xview.ViewItems
            set xitemdes = ExisteBO(atransaccion, "ITEMTRDESIMPUTACION", "REFERENCIA", CStr(xitemimputacion.bo.placeowner.ID), Nothing, True, False, "=")
            If xitemdes Is Nothing Then
                If Not d.Exists(CStr(xitemimputacion.bo.placeowner.ID)) Then
                    imputaciones.Add(xitemimputacion.bo.placeowner)
                    d.Add CStr(xitemimputacion.bo.placeowner.ID), "" 
                End If
            End If
        Next
        If imputaciones.Size > 0 Then
            set xtrdesimputacion = CrearTransaccion(xcodigoTRDesimputacion, atransaccion.unidadOperativa)
            set xtrdesimputacion.Destinatario = atransaccion.Destinatario
            For Each xtrimputacion In imputaciones
                set xitem = CrearItemTransaccion(xtrdesimputacion)
                set xitem.Referencia = xtrimputacion
            Next
            ShowBO(xtrdesimputacion)
            set DesimputarTransaccion = xtrdesimputacion
        Else
            MsgBox("No se encontro la imputacion correspondiente, verifique las imputaciones antes de continuar")
            Exit Function
        End If
End Function
