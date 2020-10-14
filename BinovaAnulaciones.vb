''BinovaAnulaciones 1.0.0'

Sub include(sFileName)
Stop
	dim oFileSystem
	dim oFile
	dim sCodigo
	
	set oFileSystem = createobject("scripting.filesystemobject")
	set oFile = oFileSystem.OpenTextFile(sFileName, 1)
	sCodigo = oFile.ReadAll()
	
	executeglobal sCodigo
End Sub


'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' Librerías incluidas.
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
include "\\Repositorio\calipso\Util\binovafuncs\binovafuncs.vbs"

Sub Main 

Stop
 
 set Transaccion 	 		= Objeto.Value
 strFuncion  	 		= Funcion.Value
 vTipotrgenerar			= Tipotrgenerar.Value
 vTipoTrEgresoAGenerar 	= TipoTrEgresoAGenerar.Value
 vCodigoTRDesimputacion 	= codigoTRDesimputacion.Value
 vIdFlagAnulado			= ID_Flag_Anulado.Value
 vRemuevePendientes 		= RemuevePendientes.Value
 vRevierteAsiento			= RevierteAsiento.Value
 vReemite				= Reemite.Value
 vGeneraCP				= GeneraCP.Value
 vRetornaTransaccion		= RetornaTransaccion.Value
 vCambianumero					= Cambianumero.Value
 VCodigoConceptoContableAnulacion	= aCodigoConcContableAnulacion.Value
 vEjercicio						= Ejercicio.Value
 vFechaAplicacion					= FechaAplicacion.Value
 vPrefijoNumero   				= PrefijoNumero.Value    
 vTrinventario					= EsTrInventario.Value
 vMensaje						= Mensaje.Value
 
 
 
    Select case strFuncion
		
		'Anulación
		
		Case "AnulaTransaccionSencilla"  
			set returnvalue = AnulaTransaccionSencilla(Transaccion,vIdFlagAnulado,vRevierteAsiento, vMensaje)
		
		Case "AnulaTRCompra"
			returnvalue = AnulaTRCompra(Transaccion, vTipotrgenerar,vCodigoTRDesimputacion,vIdFlagAnulado,vRemuevePendientes,vRevierteAsiento)
			
		Case "AnulaOrdenPago"
			returnvalue = AnulaOrdenPago(Transaccion,vTipotrgenerar, vCodigoTRDesimputacion ,vIdFlagAnulado,VCodigoConceptoContableAnulacion,vTipoTrEgresoAGenerar,vReemite)	
		
		Case "AnulaEgreso"	
			'Si devuelvo la Transaccion o no
			If vRetornaTransaccion Then
				set returnvalue = AnulaEgresoValor(Transaccion ,  vTipotrgenerar ,  vReemite ,  vIdFlagAnulado , vGeneraCP , vCodigoTRDesimputacion , vRemuevePendientes ,  vRevierteAsiento , vCambianumero , vRetornaTransaccion ) 
			Else
				returnvalue = AnulaEgresoValor(Transaccion ,  vTipotrgenerar ,  vReemite ,  vIdFlagAnulado , vGeneraCP , vCodigoTRDesimputacion , vRemuevePendientes ,  vRevierteAsiento , vCambianumero , vRetornaTransaccion )
			End if
		
		Case "AnulaIngresoValor"	
			set returnvalue = AnulaIngresoValor(Transaccion ,  vTipotrgenerar ,  vIdFlagAnulado , vRemuevePendientes,vGeneraCP , vCodigoTRDesimputacion  , vCambianumero) 
			
		Case "AnulaTRInventario"
			set returnvalue = AnulaTRInventario(Transaccion, vTipotrgenerar , vIdFlagAnulado, vRemuevePendientes,vRevierteAsiento)
			
		Case "anulaTRVentas"
			set Returnvalue	= anulaTRVentas(Transaccion  , vTipotrgenerar  ,   vCodigoTRDesimputacion  ,   vIdFlagAnulado  ,   vRemuevePendientes  , vRevierteAsiento  )  
		
		Case "anulaTRDelgadaOrden"		
			set Returnvalue = anulaTRDelgadaOrden( Transaccion  ,   vTipotrgenerar  ,   vCodigoTRDesimputacion  ,   vIdFlagAnulado  ,   vRemuevePendientes  ,     vRevierteAsiento   ,     vPrefijoNumero   ,     vTrinventario  ) 
			
		Case "anulaParteServicios"
			set Returnvalue = anulaParteServicios( Transaccion  ,   vTipotrgenerar  ,   vIdFlagAnulado  ,   vRemuevePendientes  ,  vCambianumero  ,     vRevierteAsiento  )  
		
		'Varias
		
		Case "GetAsiento"
			set returnvalue = GetAsiento(Transaccion)	
		
		Case "ReversarAsiento"
			set returnvalue = ReversarAsiento(Transaccion , vEjercicio , vFechaAplicacion )
		
		Case "DesimputarTransaccion"
			set returnvalue = DesimputarTransaccion(Transaccion  , vCodigoTRDesimputacion )	
			
	End Select
 
End sub