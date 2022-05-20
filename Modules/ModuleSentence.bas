Attribute VB_Name = "ModuleSentence"
Sub main()
    With Base
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BDProyecto.MDB;Persist Security Info=False"
        LoginForm.Show
    End With
End Sub
'manipular usuarios
Sub Usuarios()
With RsUsuarios
    If .State = 1 Then .Close
    .Open "select * from Usuarios order by Nombre asc", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub
'manipular productos
Sub Productos()
With RsProductos
    If .State = 1 Then .Close
    .Open "select * from Productos order by Descripcion asc", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub
'ordeno por ID para resolver el problema del desorde si ordeno por fecha
'manipular despachos
Sub Despacho()
With RsDespacho
    If .State = 1 Then .Close
    .Open "select * from Despacho order by Fecha DESC", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

'manipular detalles de los despachos
Sub DetallesDespacho()
With RsDetallesDespacho
    If .State = 1 Then .Close
    .Open "select * from DetallesDespacho", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

'manipular el temporal
Sub TemporalDespacho()
With RsTemporalDespacho
    If .State = 1 Then .Close
    .Open "SELECT * FROM TemporalDespacho", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

'manipular las devoluciones
Sub Devoluciones()
With RsDevoluciones
    If .State = 1 Then .Close
    .Open "select * from Devoluciones order by FechaDevo asc", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

'manipular los detalles de las devoluciones
Sub DevolucionesDetalles()
With RsDevolucionesDetalles
    If .State = 1 Then .Close
    .Open "select * from DevolucionesDetalles", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

'manipular el temporal de las devoluciones
Sub TemporalDevoluciones()
With RsTemporalDevoluciones
    If .State = 1 Then .Close
    .Open "select * from TemporalDevoluciones", Base, adOpenDynamic, adLockOptimistic 'no me corria por esta linea dado que tenia que poner adLockOptimistic y tenia otra cosa
End With
End Sub

Sub MostrarSeccion()
With RsProductos
          If .State = 1 Then .Close
          .Open "SELECT * FROM Productos WHERE Ubicacion LIKE '" & vSeccion & "' ORDER BY Descripcion ASC"
          .Requery
          Set UbicacionesSeccionesForm.GrillaSecciones.DataSource = RsProductos
          UbicacionesSeccionesForm.EstilosSecciones
End With
End Sub

Public Function Comprobar_Mail(Direccion As String) As Boolean
On Error GoTo ErrFunction
    Dim oReg As RegExp
    ' Crea un Nuevo objeto RegExp
    Set oReg = New RegExp
    ' Expresión regular
    oReg.Pattern = "^[\w-\.]+@\w+\.\w+$"
    ' Comprueba y Retorna TRue o false
    Comprobar_Mail = oReg.Test(Direccion)
    Set oReg = Nothing
Exit Function
'Error
ErrFunction:
    MsgBox Err.Description, vbCritical
    If Not oReg Is Nothing Then
        Set oReg = Nothing
    End If
End Function

Sub Logs()
With RsLogs
          If .State = 1 Then .Close
          .Open "SELECT * FROM Logs ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub LogsProductos()
With RsLogsProductos
          If .State = 1 Then .Close
          .Open "SELECT * FROM LogsProductos ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub LogsUsuarios()
With RsLogsUsuarios
          If .State = 1 Then .Close
          .Open "SELECT * FROM LogsUsuarios ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub LogsMantenimiento()
With RsLogsMantenimiento
          If .State = 1 Then .Close
          .Open "SELECT * FROM LogsMantenimiento ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub LogsReportes()
With RsLogsReportes
          If .State = 1 Then .Close
          .Open "SELECT * FROM LogsReportes ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub LogsDespachos()
With RsLogsDespachos
          If .State = 1 Then .Close
          .Open "SELECT * FROM LogsDespachos ORDER BY Fecha DESC", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub Config()
With RsConfig
          If .State = 1 Then .Close
          .Open "SELECT * FROM Config", Base, adOpenDynamic, adLockOptimistic
End With
End Sub

Sub Shipping()
With RsShipping
          If .State = 1 Then .Close
          .Open "SELECT * FROM Shipping", Base, adOpenDynamic, adLockOptimistic
End With
End Sub
