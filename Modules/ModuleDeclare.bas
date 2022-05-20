Attribute VB_Name = "ModuleDeclare"
'variable de conexion a base de datos
Global Base As New ADODB.Connection

'variables de recorset
Global RsUsuarios As New ADODB.Recordset
Global RsProductos As New ADODB.Recordset 'meti aqui

Global RsDespacho As New ADODB.Recordset
Global RsDetallesDespacho As New ADODB.Recordset
Global RsTemporalDespacho As New ADODB.Recordset

Global RsDevoluciones As New ADODB.Recordset
Global RsDevolucionesDetalles As New ADODB.Recordset
Global RsTemporalDevoluciones As New ADODB.Recordset

Global RsConfig As New ADODB.Recordset
Global RsShipping As New ADODB.Recordset

'variables para los logs
Global RsLogs As New ADODB.Recordset
Global RsLogsProductos As New ADODB.Recordset
Global RsLogsUsuarios As New ADODB.Recordset
Global RsLogsMantenimiento As New ADODB.Recordset
Global RsLogsReportes As New ADODB.Recordset
Global RsLogsDespachos As New ADODB.Recordset

'variables de productos
Global vIDProductos As Integer
Global vMostrarProducto As Integer 'añadi este el 5-oct
Global vDProductos As Integer
Global vMostrarDetallesDespacho As Integer

'variables de las devoluciones
Global vMostrarDetallesDevoluciones As String
Global vDProductosDevolucion As String
Global vDProductoDevolucion As String

'variables para mostrar nombre y tarea
Global vUsername As String
Global vTarea As String

'variables para la imagen
Global RutaOrigen As String
Global RutaDestino As String
Global ArchivoNombre As String

'variables para el respaldo DB
Global RutaOrigenDB As String
Global RutaDestinoDB As String
Global ArchivoNombreDB As String

'variables para el reporte a la hora de crear e imprimir
Global vNFactura As String
Global vCliente As String
Global vVendedor As String
Global vDespachador As String
Global vFecha As String
Global vZona As String

'variables para el reporte de devoluciones
Global vNFacturaDevo As String
Global vClienteDevo As String
Global vMotivo As String
Global vZonaDevo As String
Global vVendedorDevo As String
Global vObservaciones As String
Global vEntregado As String
Global vDevuelto As String
Global vDespachadorDevo As String
Global vDevueltoPor As String

'para el filtro de ubicaciones
Global vSeccion As String

'variables para mis Etiquetas
Global TagSeccion As String
Global TagSeccion1 As String
Global TagMin As String
Global TagMax As String
Global TagCodigo As String
Global TagCodigo1 As String
Global TagDescripcion As String
Global TagDescripcion1 As String
Global TagAplicable As String
Global TagKit As String
Global TagPzas As String

'variables del log y login forzoso
Global Guilty As String 'quien restablece la bd NO LA ESTOY USANDO
Global vNProductos As String ' cuantos productos tenia el despacho
Global vNProductosDevo As String 'cuantos productos devolvio NO LA ESTOY USANDO AUN
Global vRestablecer As Integer ' cuando se restablece esta variable tiende a 1 y asi se que se restablecio y uso el guilty

'variables para los filtros de LOGS
Global vMesLogs As String
Global vAñoLogs As String

'variables para el cambio de imagenes del login
Global Error As String
Global Normal As String

'variable para descontar en la devolucion
Global vDescuento As String
Global vDescuentoOn As String

'variables para las etiquetas del shipping
Global vCodigoShipping As String
Global vDireccionShipping As String
Global vBultoShipping As String
Global vClienteShipping As String
Global vZonaShipping As String
Global vFechaShipping As String



