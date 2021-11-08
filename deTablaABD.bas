' (GSC, 29/04/19) Módulo que recorre una tabla de Excel y actualiza e inserta datos en Oracle.

' Procedimiento principal. Se llamará desde los botones en cada hoja de la Excel.
Sub TrasladarTablaaBD()
    ' 0. Variables
    Dim hojaActiva As String
    Dim correspondencia As String: correspondencia = "CORRESPONDENCIA"
    Dim NombreTablaBD As String
    Dim NombreTabla As String
    Dim Mascara As String
    Dim MascaraTipos As String
    
    ' 1. Obtengo el nombre interno de la hoja activa. Así los nombres de las hojas pueden cambiar
    '    en caso de ser necesario.
    hojaActiva = ActiveSheet.CodeName

    ' 2. Busco la correspondencia HojaActiva - Tabla BD
    '    Estas correspondencias están predefinidas en la hoja correspondencia
    Call ObtenerCorrespondenciaBD(correspondencia, hojaActiva, NombreTabla, NombreTablaBD, Mascara, MascaraTipos)
     
    ' 3. Informo si tenemos un "" devuelto.
    If (NombreTablaBD = "") Then
        MsgBox ("Correspondencia Hoja - Tabla BD no establecida")
    Else
        Call TraspasarInfoABD(hojaActiva, NombreTabla, NombreTablaBD, Mascara, MascaraTipos)
    End If
End Sub

' Función que obtiene la correspondencia del CodeName de la hoja con el nombre de la tabla de BD
Private Function ObtenerCorrespondenciaBD(HojaCorrespondencia As String, NombreHoja As String, NombreTabla As String, NombreTablaBD As String, Mascara As String, MascaraTipos As String)
    ' 0. Variables
    Dim columna As Integer: columna = 1
    Dim fila As Integer: fila = 2
    Dim encontrado As Boolean: encontrado = False
    Dim CC_NombreHoja As String
    
    ' 1. Obtengo la primera celda de la hoja de correspondencias y voy comprobando hasta que la encuentre o sea vacía
    Set celda = Sheets(HojaCorrespondencia).Cells(fila, columna)
    
    Do Until celda.Value = "" Or encontrado
        
        ' 1.a Obtengo el valor de la celda. Corresponde al NameCode de la hoja.
        CC_NombreHoja = celda.Value
        
        ' 1.b Compruebo si el nombre de la hoja corresponde con el parámetro
        If (CC_NombreHoja = NombreHoja) Then
           encontrado = True
        Else
            ' 1.c No encontrado, avanzamos en fila.
            fila = fila + 1
            Set celda = Sheets(HojaCorrespondencia).Cells(fila, columna)
        End If
    Loop
    
    ' 2. Si hemos encontrado la correspondencia
    If (encontrado) Then
        ' 2.a Obtenemos el nobre de la tabla de bd y de la hoja
        NombreTablaBD = Sheets(HojaCorrespondencia).Cells(fila, 3).Value
        NombreTabla = Sheets(HojaCorrespondencia).Cells(fila, 4).Value
        Mascara = Sheets(HojaCorrespondencia).Cells(fila, 5).Value
        MascaraTipos = Sheets(HojaCorrespondencia).Cells(fila, 6).Value
    Else
        ' 2.b Devolvemos un "" que revisaremos desde fuera.
        NombreTabla = ""
        NombreTablaBD = ""
        Mascara = ""
        MascaraTipos = ""
    End If
End Function

' Función que inserta/actualiza los registros de la tabla de la hoja activa en la BD.
Private Function TraspasarInfoABD(hojaActiva As String, NombreTablaXLS As String, NombreTablaBD As String, Mascara As String, MascaraTipos As String)
    ' 0. Variables
    Dim objectos As ListObject
    Dim nObjetos As Integer
    Dim cont As Integer: cont = 1
    Dim encontrado As Boolean: encontrado = False
    Dim ConexionBD As ADODB.Connection
    
    '
    Dim headers() As String
        
    ' 1. Debemos recorrer los ListObjects para encontrar la Tabla que se ha puesto.
    nObjetos = ActiveSheet.ListObjects.Count
    
    Do While ((cont <= nObjetos) And (Not encontrado))
        objetos = ActiveSheet.ListObjects(cont)
        If (objetos = NombreTablaXLS) Then
            encontrado = True
        Else
            cont = cont + 1
        End If
    Loop
    
    If (Not encontrado) Then
        MsgBox ("Correspondencia Hoja - Tabla Datos no establecida")
    Else
        
        'Construccion de los SQL's
        Dim consulta As String: consulta = "SELECT * FROM " & NombreTablaBD
        Dim consultaWhere As String: consultaWhere = " WHERE "
        Dim update As String: update = "UPDATE " & NombreTablaBD & " SET "
        Dim updateSetters As String
        Dim updateWheres As String: updateWheres = " WHERE "
        
        'Inicializo variable
        Dim esRegistroActualizable As Boolean: esRegistroActualizable = False
        Dim actualizaciones As Integer: actualizaciones = 0
        
        'Obtengo el número de columnas de la tabla.
        Dim numHeaders As Integer: numHeaders = ActiveSheet.ListObjects(cont).HeaderRowRange.Count
        
        'Redimensiono el array de los headers.
        ReDim headers(numHeaders - 1) As String
        
        ' 2. Abro conexión con ORACLE
        Set ConexionBD = ObtenerConexionOracle
        
        Dim valorCelda As String
        Dim valorCabecera As String
        
        Dim filaAct As Integer: filaAct = 0
        Dim columnaCelda As Integer: columnaCelda = 0
        'Recorro las celdas visibles de la tabla, es decir las filtradas. Monto las consultas y ejecuto.
        For Each cell In ActiveSheet.ListObjects(cont).Range.SpecialCells(xlCellTypeVisible).Cells
            'Si estamos en la fila 0, guardamos los headers para poder acceder luego a ellos
            ' para construir los SQL's
            ' (GSC, 21/10/21) Particularidad del desarrollo, la primera columna de la tabla no está en la 0, está en la 3.
            '     Recupero el valor y le resto 0, para que el primer campo de la tabla lo deje en headers(0), no en headers(2)
            columnaCelda = cell.Column - 2
            If (filaAct = 0) Then
                headers(columnaCelda - 1) = cell.Value
            Else
            
                'Obtenemos el valor de la celda.
                valorCelda = cell.Value
                'Obtenemos el valor del header de la celda
                valorCabecera = headers(columnaCelda - 1)
            
                Dim valorParseado As String
                
                'Comprobamos la máscara para ver qué hacemos con el registro/celda.
                '    0. El campo en esa posición no se tiene en cuenta para el acceso a la BD.
                '    1. El campo en esa posición se tiene en cuenta en el SELECT. Campos fijos, no actualizables
                '    2. El campo en esa posición se tiene en cuenta en el UPDATE. Campos actualizables, el WHERE se puede usar el del SELECT.
                '    3. Campo de actualización VERDADERO/FALSO
                
                ' (GSC, 22/10/21) Añado la opción de que las cabeceras que aparecen en la Excel sean distintas a las que tenemos en BD.
                '   Por esto, es necesario tener una correlación para construir bien el SQL.
                
                valorCabecera = CambiaHeaderPorBD(valorCabecera)
                
                Select Case ComprobarMascara(columnaCelda, Mascara)
                    Case 0
                       ' Paso del campo
                    Case 1
                       ' SELECT
                        valorParseado = TrataValor(valorCelda, columnaCelda, MascaraTipos)
                        consultaWhere = consultaWhere & valorCabecera & "=" & valorParseado & " AND "
                    Case 2
                        ' UPDATE
                        valorParseado = TrataValor(valorCelda, columnaCelda, MascaraTipos)
                        'valorCabecera = Left(valorCabecera, Len(valorCabecera) - 4)
                        updateSetters = updateSetters & valorCabecera & "=" & valorParseado & ", "
                    Case 3
                        ' FLAG control Actualizable VERDADERO/FALSO
                        esRegistroActualizable = (valorCelda = "VERDADERO")
                End Select
                              
                                
            End If
            'Comprobamos si estamos en la última celda para sumar 1 a las filas recorridas.
            If ((columnaCelda Mod numHeaders) = 0) Then
                ' En el caso de que estemos en la última columna, será donde tenemos que ejecutar las consultas.
                ' Siempre y cuando no estemos en la primera fila.
                
                If (filaAct > 0) Then
                
                    ' SELECT: Quito los últimos 4 caracteres al WHERE puesto que tiene un AND de mas
                    consultaWhere = Left(consultaWhere, Len(consultaWhere) - 4)
                    consulta = consulta & consultaWhere
                    
                    ' UPDATE: Quito el último caracter de los Setters
                    '    Me aprovecho del WHERE de la consulta puesto que son fijos.
                    updateSetters = Left(updateSetters, Len(updateSetters) - 2)
                    update = update & updateSetters & consultaWhere
                    MsgBox update
                                     
                    '    Debemos ejecutar el update siempre y cuando el registro esté marcado para actualizar.
                    If (esRegistroActualizable) Then
                        Select Case EjecutarSQLRecordset(consulta, ConexionBD)
                            Case -1
                                ' Se ha producido un error al ejecutar el SELECT sobre BD
                                MsgBox "Error de actualización"
                            Case 0
                                ' Ejecutamos el UPDATE sobre BD
                                If (EjecutarSQLExecute(update, ConexionBD)) Then
                                    actualizaciones = actualizaciones + 1
                                End If
                            Case 1
                                ' El registro no existe en BD. pasamos de él puesto que no controlamos inserts.
                                'TODO: qué hacer en este caso?
                        End Select
                    End If
                       
                    ' Inicializamos las variables
                    consulta = "SELECT * FROM " & NombreTablaBD
                    consultaWhere = " WHERE "
                    update = "UPDATE " & NombreTablaBD & " SET "
                    updateSetters = ""
                    updateWheres = " WHERE "
                
                    esRegistroActualizable = False
                        
                End If
                
                
                filaAct = filaAct + 1
            End If
            
        Next
       
        MsgBox ("Updates: " & actualizaciones)
        
        Call DesconexionOracle(ConexionBD)
        
    End If
    
End Function

' Funcion que comprueba el valor de la máscara en ese "bit"
Private Function ComprobarMascara(indiceCampo As Integer, Mascara As String) As Integer
    ComprobarMascara = Mid(Mascara, indiceCampo, 1)
End Function

' Funcion que comprueba el valor de la máscara en ese "bit" pero que devuelve el contenido como String
Private Function ComprobarMascaraTipos(indiceCampo As Integer, Mascara As String) As String
    ComprobarMascaraTipos = Mid(Mascara, indiceCampo, 1)
End Function

Public Function ObtenerConexionOracle() As ADODB.Connection

    Dim pCadenaConexionOracle As String
   

    Dim conexionOracle As ADODB.Connection
    Set conexionOracle = New ADODB.Connection
    
    ' Es necesario indicar los parámetros de conexión a la BD. 
    pCadenaConexionOracle = "Provider=OraOLEDB.Oracle;" & _
                            "Data Source=******;" & _
                            "User Id=******;" & _
                            "Password=*******"
    
    With conexionOracle
        .ConnectionString = pCadenaConexionOracle
        .Open
    End With
    
    Set ObtenerConexionOracle = conexionOracle
End Function

Public Function DesconexionOracle(conexion As ADODB.Connection)
    conexion.Close
    
End Function

Private Function EjecutarSQLExecute(consulta As String, conexion As ADODB.Connection) As Boolean
    On Error GoTo fail
    'MsgBox consulta
    conexion.Execute consulta
    
    GoTo ok
fail:
    'MsgBox consulta, "EjecutarSQLRecordset"
    EjecutarSQLExecute = False
ok:
    EjecutarSQLExecute = True
End Function

Private Function EjecutarSQLRecordset(consulta As String, conexion As ADODB.Connection) As Integer
    On Error GoTo fail
    
    ' Creo recordset
    Dim RegistroSql As New ADODB.RecordSet
    ' Ejecuto consulta
    RegistroSql.Open consulta, conexion
    ' Devuelvo si me ha dado resultados
    ' Si devuelve EOF, no existe
    If (RegistroSql.EOF) Then
        EjecutarSQLRecordset = 1
        GoTo ok
    Else
        EjecutarSQLRecordset = 0
        GoTo ok
    End If
    
fail:
    EjecutarSQLRecordset = -1
ok:

End Function

Private Function ComaPorPunto(valorInicial As String) As String
    ComaPorPunto = Replace(valorInicial, ",", ".")
End Function
' Función mediante la cual tratamos el valor
Private Function TrataValor(valorInicial As String, columnaCelda As Integer, MascaraTipos As String) As String
    Dim valor As String
    Dim tipoValor As Integer
    If (Len(valorInicial) = 0) Then
        ' Si el valor es = '', tengo que meter un espacio para evitar que casque.
        TrataValor = "' '"
    Else
        ' (GSC, 21/10/21) No está funcionando correctamente. Ni utilizando el código
        ' anterior ni utilizando VarType puesto que siempre cree que es un String.
        ' (GSC, 22/10/21) Utilizo el método de comprobación de máscara pero con la máscara
        ' de los tipos.
        Select Case ComprobarMascaraTipos(columnaCelda, MascaraTipos)
        Case "F"
            ' Si es un número, tengo que cambiar la , por .
            TrataValor = ComaPorPunto(valorInicial)
        Case "S"
            ' Si es un string, tengo que meterle las '
            TrataValor = "'" & valorInicial & "'"
        Case "D"
            ' Si es una fecha, tengo que ponerle el to_date
            TrataValor = "to_date('" & valorInicial & "','dd/mm/yyyy')"
        End Select
        
    End If
    
End Function
' Función en la que establecemos la correlación entre las cabeceras de la tabla y los atributos de la BD.
Private Function CambiaHeaderPorBD(header_actual As String) As String
    Select Case header_actual
    Case "CODIGO_CLIENTE"
        CambiaHeaderPorBD = "CUSTOMER_NUMBER"
    Case "SALES_PRICE_NEW"
        CambiaHeaderPorBD = "SALES_PRICE"
    Case "PRICE_CHANGE_DATE_NEW"
        CambiaHeaderPorBD = "PRICE_CHANGE_DATE"
    Case "SALES_PRICE_1_NEW"
        CambiaHeaderPorBD = "SALES_PRICE_1"
    Case Else
        CambiaHeaderPorBD = header_actual
    End Select
End Function
