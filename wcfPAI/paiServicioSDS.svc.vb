Option Explicit On
Option Strict On

'Imports paiServicios.Sds.SaludPublica.paiServicioSDS
Imports System.Data.SqlClient
Imports Sds.PAI.wsPaiEntity
Imports System.IO

Public Class paiServicioSDS
        Implements IpaiServicioSDS

    Private Const conexionBD As String = "Data Source = SDSPRUEBD1\PRUEBASSQL2014,27849;Initial Catalog= Pai;User Id=usrApp_Pai; password=Pai*$20$15_/;Persist Security Info=False"
    Private Const MAYOR_EDAD As Decimal = 18
    Private Const ESTADO_ACTIVO As Integer = 2
    Private Const ESTADO_FALLECIDO As Integer = 4

    Private fechaBase As Date = CDate("01/01/1862")
        Private Const rutaLog As String = "C:\temp\"
        Public MsgValidacion As String = ""


    Public Sub New()
        End Sub

        Private Sub insertarLog(metodo As String, detalleError As String, arParms() As SqlParameter)
            Try
                Const fic As String = rutaLog & "logServicioWebPAI4.txt"

                If Not File.Exists(fic) Then
                    Dim oArchivo As System.IO.FileStream = System.IO.File.Create(fic)
                    oArchivo.Close()
                    oArchivo.Dispose()
                End If

                Dim sw As System.IO.StreamWriter = New StreamWriter(fic, True)
                sw.WriteLine("Metodo: " & metodo & ";PosibleError:" & detalleError & ";Fecha: " & Date.Now.ToString())
                If Not arParms Is Nothing Then
                    Dim lineaParametro As String = String.Empty
                    Dim lineaDatos As String = String.Empty
                    For Each param As SqlParameter In arParms
                        lineaParametro &= param.ParameterName.ToString & ";"
                        lineaDatos &= param.Value.ToString & ";"
                    Next
                    sw.WriteLine(lineaParametro)
                    sw.WriteLine(lineaDatos)
                    sw.WriteLine("")
                End If

                sw.Close()
                sw.Dispose()
            Catch ex As Exception

            End Try
        End Sub


    Private Sub insertarLogBD(metodo As String, detalleError As String, arParms() As SqlParameter)
        Try
            Dim lineaParametro As String = String.Empty
            Dim lineaDatos As String = String.Empty
            If Not arParms Is Nothing Then
                For i As Integer = 0 To 3
                    lineaParametro &= arParms(i).ParameterName.ToString & ";"
                    lineaDatos &= arParms(i).Value.ToString & ";"
                Next
                'For Each param As SqlParameter In arParms
                '    lineaParametro &= param.ParameterName.ToString & ";"
                '    lineaDatos &= param.Value.ToString & ";"
                'Next
            End If

            Dim cadena As String = conexionBD

            Dim arParmsInterno() As SqlParameter = New SqlParameter(25) {}

            arParmsInterno(0) = New SqlParameter("@metodo", 12)
            arParmsInterno(0).Value = metodo

            arParmsInterno(1) = New SqlParameter("@detalle", 12)
            arParmsInterno(1).Value = detalleError

            arParmsInterno(2) = New SqlParameter("@lineaParametros", 12)
            arParmsInterno(2).Value = lineaParametro

            arParmsInterno(3) = New SqlParameter("@lineaDatos", 12)
            arParmsInterno(3).Value = lineaDatos

            SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarLogServicioWeb", arParmsInterno)
        Catch ex As Exception

        End Try
    End Sub

    Public Function calcularEdad(ByVal FechaNac As Date) As Decimal
        Dim Edad As Decimal
        Dim AnioEdad As Integer = DatePart(DateInterval.Year, FechaNac)
        Dim mesesEdad As Integer = DatePart(DateInterval.Month, FechaNac)
        Dim anioactual As Integer = DatePart(DateInterval.Year, Now)
        Dim mesactual As Integer = DatePart(DateInterval.Month, Now)
        Dim N As Single, a As Integer, m As Integer
        N = CSng(DateDiff("m", FechaNac, Now) / 12)
        a = CInt(Int(N))
        N = (N - a) * 12
        m = CInt(Int(N))
        Edad = CDec(a & "," & m)
        Return Edad
    End Function

    Public Function validarTablasDominio(id_Tabla As Short, codigoTabla As String, ByRef MensajeValidacion As String) As Boolean
            Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
            Dim nroRegistros As Integer = 0
            Dim fValidacion As Boolean = True

            Dim arParms() As SqlParameter = New SqlParameter(1) {}
            arParms(0) = New SqlParameter("@id_Tabla", 8)
            arParms(0).Value = id_Tabla
            arParms(1) = New SqlParameter("@codigoTabla", 22)
            arParms(1).Value = codigoTabla

            nroRegistros = CInt(SqlHelper.ExecuteScalar(cadena, "pa_ValidarTablasDominio", arParms))
            If nroRegistros <= 0 Then
                fValidacion = False
                MensajeValidacion = "El código: " + codigoTabla + " no existe en la tabla: " + nombreTabla(id_Tabla) + " "
            End If
            Return fValidacion
        End Function

        Public Function nombreTabla(idTabla As Short) As String
            Dim tabla As String = String.Empty

            If idTabla = 1 Then
                tabla = "Aseguradora"
            End If

            If idTabla = 2 Then
                tabla = "TIPO DE AFILIACIÓN"
            End If

            If idTabla = 2 Then
                tabla = "CAUSA DE NO VACUNACIÓN"
            End If

            If idTabla = 3 Then
                tabla = "CAUSA DE NO VACUNACIÓN"
            End If

            If idTabla = 4 Then
                tabla = "VACUNA"
            End If

            If idTabla = 5 Then
                tabla = "DOSIS"
            End If

            If idTabla = 6 Then
                tabla = "PREGUNTA CUESTIONARIO"
            End If

            If idTabla = 7 Then
                tabla = "BARRIO"
            End If

            If idTabla = 8 Then
                tabla = "UPZ"
            End If

            If idTabla = 9 Then
                tabla = "LOCALIDAD"
            End If

            If idTabla = 10 Then
                tabla = "MUNICIPIO"
            End If

            If idTabla = 11 Then
                tabla = "DEPARTAMENTO"
            End If

            If idTabla = 12 Then
                tabla = "PAÍS"
            End If

            If idTabla = 13 Then
                tabla = "ZONA"
            End If

            If idTabla = 14 Then
                tabla = "TIPO DE DOCUMENTO"
            End If

            If idTabla = 15 Then
                tabla = "INSTITUCIÓN"
            End If

            If idTabla = 16 Then
                tabla = "ESTADO PERSONA"
            End If

            If idTabla = 17 Then
                tabla = "CAUSA DE NO PRESENTAR EL CERTIFICADO DE NACIDO VIVO"
            End If

            If idTabla = 18 Then
                tabla = "ETNIA"
            End If

            If idTabla = 19 Then
                tabla = "GRUPO POBLACIONAL"
            End If

            If idTabla = 20 Then
                tabla = "GÉNERO"
            End If

            If idTabla = 21 Then
                tabla = "GRUPO SANGUINEO"
            End If

            If idTabla = 22 Then
                tabla = "RH"
            End If

            If idTabla = 23 Then
                tabla = "CONDICIÓN MUJER"
            End If

            If idTabla = 24 Then
                tabla = "PRESENTACIÓN"
            End If

            If idTabla = 25 Then
                tabla = "PRESENTACIÓN COMERCIAL"
            End If

            If idTabla = 26 Then
                tabla = "CAMPAÑA"
            End If

            If idTabla = 27 Then
                tabla = "PERTENENCIA POS"
            End If

            If idTabla = 28 Then
                tabla = "CATEGORÍA DE CONTÁCTENOS"
            End If

            If idTabla = 29 Then
                tabla = "ESTADO CONTACTENOS"
            End If

            If idTabla = 30 Then
                tabla = "MOTIVO DE NO COMUNICACIÓN SEGUIMIENTO A COHORTE"
            End If

            If idTabla = 31 Then
                tabla = "GRUPO POBLACIONAL"
            End If

            If idTabla = 32 Then
                tabla = "FUNCIONARIO"
            End If

            If idTabla = 33 Then
                tabla = "TIPO DE SEGUIMIENTO"
            End If

            If idTabla = 34 Then
                tabla = "MOTIVO DE NO VACUNACIÓN"
            End If
            If idTabla = 35 Then
                tabla = "PESO"
            End If
            Return tabla
        End Function

    Public Function actualizarPersona(per_consecutivo As Long,
                                          per_TipoId As String,
                                          per_Id As String,
                                          per_CertNacVivo As Long,
                                          per_CertDefuncion As String,
                                          per_TipoIdM As String,
                                          per_IdM As String,
                                          per_NumeroHijoM As Integer,
                                          per_Nombre1M As String,
                                          per_Nombre2M As String,
                                          per_Apellido1M As String,
                                          per_Apellido2M As String,
                                          per_Nombre1 As String,
                                          per_Nombre2 As String,
                                          per_Apellido1 As String,
                                          per_Apellido2 As String,
                                          per_FechaNac As Date,
                                          per_Func As String,
                                          per_Institucion As String,
                                          per_Estado As Integer,
                                          per_cni_id As Integer,
                                          per_idEtnia As Integer,
                                          per_IdGrupoPoblacional As String,
                                          per_IdGenero As String,
                                          per_IdGrupoSanguineo As String,
                                          per_IdRh As String) As resultadoConsultaEntity Implements IpaiServicioSDS.actualizarPersona

        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        Dim fValidacion As Boolean = True
        Dim MensajeValidacion As String = String.Empty

        '---------------Aqui se validan los valores en 0 o vacios

        If per_consecutivo = 0 Then
            oResultado.resultado = False
            oResultado.errores = "la variable per_consecutivo no puede ser 0"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_TipoId) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_tipoId no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Id) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Id no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_TipoIdM) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_TipoIdM no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_IdM) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_IdM no puede estar vacía"
            fValidacion = False
        End If
        If per_NumeroHijoM = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_NumeroHijoM no puede ser 0"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Nombre1M) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Nombre1M no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Apellido1M) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Apellido1M no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Nombre1) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Nombre1 no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Apellido1) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Apellido1 no puede estar vacía"
            fValidacion = False
        End If
        If per_FechaNac = Date.MinValue Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_FechaNac no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_Func) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_FechaNac no puede estar vacía"
            fValidacion = False
        End If

        If String.IsNullOrEmpty(per_Institucion) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Institucion no puede estar vacía"
            fValidacion = False
        End If

        If per_idEtnia = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_idEtnia no puede ser 0"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(per_IdGenero) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_IdGenero no puede estar vacía"
            fValidacion = False
        End If

        '----------------- Aqui se valida contra las tablas de dominio
        'TODO: contra que tabla valido per_consecutivo? debo colocar persona en el procedimiento almacenado
        If Not validarTablasDominio(36, per_consecutivo.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(14, per_TipoId, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores = MensajeValidacion
        End If
        If Not validarTablasDominio(14, per_TipoIdM, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(32, per_Func, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(15, per_Institucion, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(16, per_Estado.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(17, per_cni_id.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(18, per_idEtnia.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(20, per_IdGenero, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If

        If fValidacion Then
            Dim arParms() As SqlParameter = New SqlParameter(25) {}

            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_consecutivo

            arParms(1) = New SqlParameter("@per_TipoId", 3)
            If String.IsNullOrEmpty(per_TipoId) Then
                arParms(1).Value = System.DBNull.Value
            Else
                arParms(1).Value = per_TipoId
            End If

            arParms(2) = New SqlParameter("@per_Id", 22)
            If String.IsNullOrEmpty(per_Id) Then
                arParms(2).Value = System.DBNull.Value
            Else
                arParms(2).Value = per_Id
            End If

            arParms(3) = New SqlParameter("@per_CertNacVivo", 0)
            If per_CertNacVivo = 0 Then
                arParms(3).Value = System.DBNull.Value
            Else
                arParms(3).Value = per_CertNacVivo
            End If

            arParms(4) = New SqlParameter("@per_CertDefuncion", 22)
            If String.IsNullOrEmpty(per_CertDefuncion) Then
                arParms(4).Value = System.DBNull.Value
            Else
                arParms(4).Value = per_CertDefuncion
            End If

            arParms(5) = New SqlParameter("@per_TipoIdM", 3)
            If String.IsNullOrEmpty(per_TipoIdM) Then
                arParms(5).Value = System.DBNull.Value
            Else
                arParms(5).Value = per_TipoIdM
            End If

            arParms(6) = New SqlParameter("@per_IdM", 22)
            If String.IsNullOrEmpty(per_IdM) Then
                arParms(6).Value = System.DBNull.Value
            Else
                arParms(6).Value = per_IdM
            End If

            arParms(7) = New SqlParameter("@per_NumeroHijoM", 16)
            If per_NumeroHijoM = 0 Then
                arParms(7).Value = System.DBNull.Value
            Else
                arParms(7).Value = per_NumeroHijoM
            End If

            arParms(8) = New SqlParameter("@per_Nombre1M", 22)
            If String.IsNullOrEmpty(per_Nombre1M) Then
                arParms(8).Value = System.DBNull.Value
            Else
                arParms(8).Value = per_Nombre1M
            End If

            arParms(9) = New SqlParameter("@per_Nombre2M", 22)
            If String.IsNullOrEmpty(per_Nombre2M) Then
                arParms(9).Value = System.DBNull.Value
            Else
                arParms(9).Value = per_Nombre2M
            End If

            arParms(10) = New SqlParameter("@per_Apellido1M", 22)
            If String.IsNullOrEmpty(per_Apellido1M) Then
                arParms(10).Value = System.DBNull.Value
            Else
                arParms(10).Value = per_Apellido1M
            End If

            arParms(11) = New SqlParameter("@per_Apellido2M", 22)
            If String.IsNullOrEmpty(per_Apellido2M) Then
                arParms(11).Value = System.DBNull.Value
            Else
                arParms(11).Value = per_Apellido2M
            End If

            arParms(12) = New SqlParameter("@per_Nombre1", 22)
            If String.IsNullOrEmpty(per_Nombre1) Then
                arParms(12).Value = System.DBNull.Value
            Else
                arParms(12).Value = per_Nombre1
            End If

            arParms(13) = New SqlParameter("@per_Nombre2", 22)
            If String.IsNullOrEmpty(per_Nombre2) Then
                arParms(13).Value = System.DBNull.Value
            Else
                arParms(13).Value = per_Nombre2
            End If

            arParms(14) = New SqlParameter("@per_Apellido1", 22)
            If String.IsNullOrEmpty(per_Apellido1) Then
                arParms(14).Value = System.DBNull.Value
            Else
                arParms(14).Value = per_Apellido1
            End If

            arParms(15) = New SqlParameter("@per_Apellido2", 22)
            If String.IsNullOrEmpty(per_Apellido2) Then
                arParms(15).Value = System.DBNull.Value
            Else
                arParms(15).Value = per_Apellido2
            End If

            arParms(16) = New SqlParameter("@per_FechaNac", 31)
            If per_FechaNac = Date.MinValue Then
                arParms(16).Value = System.DBNull.Value
            Else
                arParms(16).Value = per_FechaNac
            End If

            arParms(17) = New SqlParameter("@per_Func", 12)
            If String.IsNullOrEmpty(per_Func) Then
                arParms(17).Value = System.DBNull.Value
            Else
                arParms(17).Value = per_Func
            End If

            arParms(18) = New SqlParameter("@per_Institucion", 3)
            If String.IsNullOrEmpty(per_Institucion) Then
                arParms(18).Value = System.DBNull.Value
            Else
                arParms(18).Value = per_Institucion
            End If

            arParms(19) = New SqlParameter("@per_Estado", 8)
            If per_Estado = 0 Then
                arParms(19).Value = System.DBNull.Value
            Else
                arParms(19).Value = per_Estado
            End If

            arParms(20) = New SqlParameter("@per_cni_id", 8)
            If per_cni_id = 0 Then
                arParms(20).Value = System.DBNull.Value
            Else
                arParms(20).Value = per_cni_id
            End If

            arParms(21) = New SqlParameter("@per_idEtnia", 8)
            If per_idEtnia = 0 Then
                arParms(21).Value = System.DBNull.Value
            Else
                arParms(21).Value = per_idEtnia
            End If

            arParms(22) = New SqlParameter("@per_IdGrupoPoblacional", 3)
            If per_IdGrupoPoblacional = "0" Then
                arParms(22).Value = System.DBNull.Value
            Else
                arParms(22).Value = per_IdGrupoPoblacional
            End If

            'Alejandro Muñoz 26/10/2017 - Actualización cambio. Para que no haga homologación, ya que compensar envía ya los datos validados'
            arParms(23) = New SqlParameter("@per_IdGenero", 3)
            If String.IsNullOrEmpty(per_IdGenero) Then
                arParms(23).Value = System.DBNull.Value
            Else
                arParms(23).Value = per_IdGenero
            End If


            'arParms(23) = New SqlParameter("@per_IdGenero", 3)
            'If String.IsNullOrEmpty(per_IdGenero) Then
            '    arParms(23).Value = System.DBNull.Value
            'Else
            '    If per_IdGenero = "M" Then
            '        arParms(23).Value = "H"
            '    ElseIf per_IdGenero = "F" Then
            '        arParms(23).Value = "M"
            '    End If
            'End If

            arParms(24) = New SqlParameter("@per_IdGrupoSanguineo", 3)
            If String.IsNullOrEmpty(per_IdGrupoSanguineo) Then
                arParms(24).Value = System.DBNull.Value
            Else
                arParms(24).Value = per_IdGrupoSanguineo
            End If

            arParms(25) = New SqlParameter("@per_IdRh", 3)
            If String.IsNullOrEmpty(per_IdRh) Then
                arParms(25).Value = System.DBNull.Value
            Else
                arParms(25).Value = per_IdRh
            End If

            Try
                If arParms(1).Value.ToString.Equals("RC") Or arParms(1).Value.ToString.Equals("TI") Then
                    insertarLog("actualizarPersona()", "VALIDACION(RC-TI)", arParms)
                    insertarLogBD("actualizarPersona()", "VALIDACION(RC-TI)", arParms)
                End If

                If SqlHelper.ExecuteNonQuery(cadena, "pa_ActualizarPersonaWS", arParms) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty
                    oResultado.consecutivo = per_consecutivo.ToString()

                    insertarLog("actualizarPersona()", "CORRECTO", arParms)
                    insertarLogBD("actualizarPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no identificado."

                    insertarLog("actualizarPersona()", "No actualizo ningun registro", arParms)
                    insertarLogBD("actualizarPersona()", "No actualizo ningun registro", arParms)
                End If

            Catch ex As SqlException
                Dim errorMessage As String = ex.Message
                Dim errorCode As Integer = ex.ErrorCode

                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("actualizarPersona()", ex.Message, arParms)
                insertarLogBD("actualizarPersona()", ex.Message, arParms)
            End Try
            'Catch ex As Exception
            '    oResultado.resultado = False
            '    oResultado.errores = ex.Message.ToString()
            'End Try

        Else
            oResultado.resultado = False
            oResultado.errores = MsgValidacion
        End If

        Return oResultado
    End Function

    Public Function insertarContactenos(con_cat_id As Integer, con_mensaje As String, con_fun_idFunc As String) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarContactenos
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity



        Dim arParms() As SqlParameter = New SqlParameter(2) {}

        arParms(0) = New SqlParameter("@con_cat_id", 16)
        arParms(0).Value = con_cat_id

        arParms(1) = New SqlParameter("@con_mensaje", 22)
        arParms(1).Value = con_mensaje

        arParms(2) = New SqlParameter("@con_fun_idFunc", 12)
        arParms(2).Value = con_fun_idFunc


        Try
            'If CLng(SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarContactenos_getId", arParms)) > 0 Then 'm 26/09/2013 pa_InsertarContactenos
            Dim idResult = (SqlHelper.ExecuteScalar(cadena, "pa_InsertarContactenos_getId", arParms))
            idResult = If(IsDBNull(idResult), -1, CLng(idResult))
            If CLng(idResult) > 0 Then 'm 26/09/2013 pa_InsertarContactenos
                oResultado.resultado = True
                oResultado.errores = String.Empty
                oResultado.consecutivo = CStr(idResult)

                insertarLog("insertarContactenos()", "CORRECTO", arParms)
            Else
                oResultado.resultado = False
                oResultado.errores = "Registro no ingresado."

                insertarLog("insertarContactenos()", "No inserto ningun registro", arParms)
            End If
        Catch ex As Exception
            oResultado.resultado = False
            oResultado.errores = ex.Message.ToString()
            insertarLog("insertarContactenos()", ex.Message, arParms)
        End Try
        Return oResultado

    End Function

    'Public Function insertarVacunaPersona(per_Consecutivo As Long, vac_Id As Integer, dos_Id As Integer, pse_Id As Integer, cam_id As Integer, vac_FechaVacuna As Date, vac_actualizacion As Boolean, ins_Id As String, fun_idFunc As String, com_Id As Integer, pos_Id As Integer, vac_Lote As String, vac_EdadVacunaAnios As Integer, vac_EdadVacunaMeses As Integer, vac_EdadVacunaDias As Integer, vac_EdadVacunaTotalDias As Integer) As PAIEntity.resultadoConsultaEntity 'Implements IpaiServicioSDS.insertarVacunaPersona
    '    Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
    '    Dim oResultado As New resultadoConsultaEntity

    '    If per_Consecutivo <> 0 Then
    '        If ValidacionDatosBasicosObligatoriosVacuna(vac_Id, dos_Id, pse_Id, ins_Id, fun_idFunc, fun_idFunc) Then

    '            Dim arParms() As SqlParameter = New SqlParameter(15) {}

    '            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
    '            arParms(0).Value = per_Consecutivo

    '            arParms(1) = New SqlParameter("@vac_Id", 8)
    '            arParms(1).Value = vac_Id

    '            arParms(2) = New SqlParameter("@dos_Id", 8)
    '            arParms(2).Value = dos_Id

    '            arParms(3) = New SqlParameter("@pse_Id", 8)
    '            arParms(3).Value = pse_Id

    '            arParms(4) = New SqlParameter("@cam_id", 8)
    '            If cam_id = 0 Then
    '                arParms(4).Value = System.DBNull.Value
    '            Else
    '                arParms(4).Value = cam_id
    '            End If

    '            arParms(5) = New SqlParameter("@vac_FechaVacuna", 31)
    '            arParms(5).Value = vac_FechaVacuna

    '            arParms(6) = New SqlParameter("@vac_actualizacion", 2)
    '            arParms(6).Value = vac_actualizacion

    '            arParms(7) = New SqlParameter("@ins_Id", 3)
    '            arParms(7).Value = ins_Id

    '            arParms(8) = New SqlParameter("@fun_idFunc", 12)
    '            arParms(8).Value = fun_idFunc

    '            arParms(9) = New SqlParameter("@com_Id", 8)
    '            arParms(9).Value = com_Id

    '            arParms(10) = New SqlParameter("@pos_Id", 20)
    '            arParms(10).Value = pos_Id

    '            arParms(11) = New SqlParameter("@vac_Lote", 22)
    '            arParms(11).Value = vac_Lote

    '            arParms(12) = New SqlParameter("@vac_EdadVacunaAnios", 8)
    '            arParms(12).Value = vac_EdadVacunaAnios

    '            arParms(13) = New SqlParameter("@vac_EdadVacunaMeses", 8)
    '            arParms(13).Value = vac_EdadVacunaMeses

    '            arParms(14) = New SqlParameter("@vac_EdadVacunaDias", 8)
    '            arParms(14).Value = vac_EdadVacunaDias

    '            arParms(15) = New SqlParameter("@vac_EdadVacunaTotalDias", 8)
    '            arParms(15).Value = vac_EdadVacunaTotalDias


    '            Try
    '                If CLng(SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarVacunaPersonaWS", arParms)) > 0 Then
    '                    oResultado.resultado = True
    '                    oResultado.errores = String.Empty

    '                    insertarLog("insertarVacunaPersona()", "CORRECTO", arParms)
    '                    insertarLogBD("insertarVacunaPersona()", "CORRECTO", arParms)
    '                Else
    '                    oResultado.resultado = False
    '                    oResultado.errores = "Registro no ingresado."

    '                    insertarLog("insertarVacunaPersona()", "No realizo la insercion del regitro", arParms)
    '                    insertarLogBD("insertarVacunaPersona()", "No realizo la insercion del regitro", arParms)
    '                End If

    '            Catch ex As SqlException
    '                Dim errorMessage As String = ex.Message
    '                Dim errorCode As Integer = ex.ErrorCode
    '                oResultado.resultado = False
    '                oResultado.errores = errorMessage

    '                insertarLog("insertarVacunaPersona()", ex.Message, arParms)
    '                insertarLogBD("insertarVacunaPersona()", ex.Message, arParms)
    '            End Try
    '            'Catch ex As Exception
    '            '    oResultado.resultado = False
    '            '    oResultado.errores = ex.Message.ToString()
    '            'End Try
    '        Else
    '            oResultado.resultado = False
    '            oResultado.errores = MsgValidacion
    '        End If
    '    Else
    '        oResultado.resultado = False
    '        oResultado.errores = "Per_consecutivo fue enviado como 0"
    '    End If
    '    Return oResultado
    'End Function


    Public Function seleccionarEsquemavacunasPAIPendiente(ByVal per_consecutivo As Long) As VacunaCollection Implements IpaiServicioSDS.seleccionarEsquemavacunasPAIPendiente
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        Dim arParms() As SqlParameter = New SqlParameter(0) {}

        arParms(0) = New SqlParameter("@per_consecutivo", 0)
        arParms(0).Value = per_consecutivo

        Dim dsEsquema As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "Pa_SeleccionarVacunasFaltantes", arParms)
        Dim EsquemaVacunacion As New VacunaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsEsquema.Tables(0).Rows
            Dim Esquema As New VacunaEntity()
            Esquema.vac_id = CInt(dr("vac_id"))
            Esquema.vac_Nombre = CStr(dr("vac_nombre"))
            Esquema.Dos_Nombre = CStr(dr("dos_nombre"))
            Esquema.Dos_id = CInt(dr("dos_Id"))
            Esquema.pos_nombre = CStr(dr("pos_Descripcion"))
            Esquema.pos_Id = CInt(dr("vac_pos"))
            'Esquema.grup_nombre = CStr(dr("grup_Nombre"))
            'Esquema.grup_Id = CInt(dr("grup_id"))

            EsquemaVacunacion.Add(Esquema)
        Next

        'insertarLog("seleccionarVacunasPersona()", "CORRECTO", "No tiene Parametros")
        dsEsquema.Dispose()

        Return EsquemaVacunacion
    End Function


    Public Function seleccionarAfiliacionPersona(per_Consecutivo As Long) As PersonaAfiliacionEntity Implements IpaiServicioSDS.seleccionarAfiliacionPersona
        Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
        Dim oPersonaAfiliacion As New PersonaAfiliacionEntity
        If per_Consecutivo <> 0 Then

            Dim arParms() As SqlParameter = New SqlParameter(0) {}
            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_Consecutivo

            Dim dsPersonaAfiliacion As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarAfiliacionPersona", arParms)
            If Not dsPersonaAfiliacion Is Nothing AndAlso dsPersonaAfiliacion.Tables(0).Rows.Count > 0 Then
                With dsPersonaAfiliacion.Tables(0)
                    oPersonaAfiliacion.per_Consecutivo = CLng(.Rows(0)("per_Consecutivo").ToString())
                    oPersonaAfiliacion.ase_id = .Rows(0)("ase_id").ToString()
                    oPersonaAfiliacion.ase_nombre = .Rows(0)("ase_nombre").ToString()
                    oPersonaAfiliacion.reg_id = CInt(.Rows(0)("reg_id").ToString())
                    oPersonaAfiliacion.reg_Nombre = .Rows(0)("reg_Nombre").ToString()
                    If Not String.IsNullOrEmpty(.Rows(0)("tia_id").ToString()) Then
                        oPersonaAfiliacion.tia_id = CInt(.Rows(0)("tia_id").ToString())
                    End If
                End With
                insertarLog("seleccionarAfiliacionPersona()", "CORRECTO", arParms)
                insertarLogBD("seleccionarAfiliacionPersona()", "CORRECTO", arParms)
            End If

        End If
        Return oPersonaAfiliacion
    End Function


    Public Function seleccionarPersonaBusqueda(TipoIdVacunado As String,
                                                   NumeroIdVacunado As String,
                                                   PrimerNombreVacunado As String,
                                                   SegundoNombreVacunado As String,
                                                   PrimerApellidoVacunado As String,
                                                   SegundoApellidoVacunado As String,
                                                   per_parInstitucion As String,
                                                   per_FechaNac As Date,
                                                   TipoIdentificacionMadre As String,
                                                   NumeroIdentificacionMadre As String,
                                                   PrimerNombreMadre As String,
                                                   SegundoNombreMadre As String,
                                                   PrimerApellidoMadre As String,
                                                   SegundoApellidoMadre As String,
                                                   grupoEtareo As Integer) As PersonaCollection Implements IpaiServicioSDS.seleccionarPersonaBusqueda

        Dim arParms() As SqlParameter = New SqlParameter(0) {}
        arParms(0) = New SqlParameter("@NumeroIdVacunado", 3)
        arParms(0).Value = NumeroIdVacunado

        insertarLog("seleccionarPersonaBusqueda()", "CORRECTO", arParms)
        insertarLogBD("seleccionarPersonaBusqueda()", "CORRECTO", arParms)

        Return seleccionarPersonaBusquedaAttr(TipoIdVacunado, NumeroIdVacunado,
                                                PrimerNombreVacunado, SegundoNombreVacunado, PrimerApellidoVacunado, SegundoApellidoVacunado,
                                                per_parInstitucion, per_FechaNac, TipoIdentificacionMadre, NumeroIdentificacionMadre,
                                                PrimerNombreMadre, SegundoNombreMadre, PrimerApellidoMadre, SegundoApellidoMadre,
                                                grupoEtareo, True, False)
    End Function


#Region "SeleccionaPersonaBusquedaCon Mas attributos"
    '--25/09/2013--
    Public Function seleccionarPersonaBusquedaAttr(TipoIdVacunado As String,
                                                       NumeroIdVacunado As String,
                                                       PrimerNombreVacunado As String,
                                                       SegundoNombreVacunado As String,
                                                       PrimerApellidoVacunado As String,
                                                       SegundoApellidoVacunado As String,
                                                       per_parInstitucion As String,
                                                       per_FechaNac As Date,
                                                       TipoIdentificacionMadre As String,
                                                       NumeroIdentificacionMadre As String,
                                                       PrimerNombreMadre As String,
                                                       SegundoNombreMadre As String,
                                                       PrimerApellidoMadre As String,
                                                       SegundoApellidoMadre As String,
                                                       grupoEtareo As Integer,
                                                       bGetMadre As Boolean,
                                                       bGetHdocum As Boolean) As PersonaCollection Implements IpaiServicioSDS.seleccionarPersonaBusquedaAttr

        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity

        Dim arParms() As SqlParameter = New SqlParameter(14) {}


        arParms(0) = New SqlParameter("@TipoIdVacunado", 3)
        If String.IsNullOrEmpty(TipoIdVacunado) Then
            arParms(0).Value = System.DBNull.Value
        Else
            arParms(0).Value = TipoIdVacunado
        End If

        arParms(1) = New SqlParameter("@NumeroIdVacunado", 22)
        If String.IsNullOrEmpty(NumeroIdVacunado) Then
            arParms(1).Value = System.DBNull.Value
        Else
            arParms(1).Value = NumeroIdVacunado
        End If

        arParms(2) = New SqlParameter("@PrimerNombreVacunado", 22)
        If String.IsNullOrEmpty(PrimerNombreVacunado) Then
            arParms(2).Value = System.DBNull.Value
        Else
            arParms(2).Value = PrimerNombreVacunado
        End If

        arParms(3) = New SqlParameter("@SegundoNombreVacunado", 22)
        If String.IsNullOrEmpty(SegundoNombreVacunado) Then
            arParms(3).Value = System.DBNull.Value
        Else
            arParms(3).Value = SegundoNombreVacunado
        End If

        arParms(4) = New SqlParameter("@PrimerApellidoVacunado", 22)
        If String.IsNullOrEmpty(PrimerApellidoVacunado) Then
            arParms(4).Value = System.DBNull.Value
        Else
            arParms(4).Value = PrimerApellidoVacunado
        End If

        arParms(5) = New SqlParameter("@SegundoApellidoVacunado", 22)
        If String.IsNullOrEmpty(SegundoApellidoVacunado) Then
            arParms(5).Value = System.DBNull.Value
        Else
            arParms(5).Value = SegundoApellidoVacunado
        End If

        arParms(6) = New SqlParameter("@per_parInstitucion", 3)
        If String.IsNullOrEmpty(per_parInstitucion) Then
            arParms(6).Value = System.DBNull.Value
        Else
            arParms(6).Value = per_parInstitucion
        End If

        arParms(7) = New SqlParameter("@per_FechaNac", 31)
        arParms(7).Value = per_FechaNac

        arParms(8) = New SqlParameter("@TipoIdentificacionMadre", 3)
        If String.IsNullOrEmpty(TipoIdentificacionMadre) Then
            arParms(8).Value = System.DBNull.Value
        Else
            arParms(8).Value = TipoIdentificacionMadre
        End If

        arParms(9) = New SqlParameter("@NumeroIdentificacionMadre", 22)
        If String.IsNullOrEmpty(NumeroIdentificacionMadre) Then
            arParms(9).Value = System.DBNull.Value
        Else
            arParms(9).Value = NumeroIdentificacionMadre
        End If

        arParms(10) = New SqlParameter("@PrimerNombreMadre", 22)
        If String.IsNullOrEmpty(PrimerNombreMadre) Then
            arParms(10).Value = System.DBNull.Value
        Else
            arParms(10).Value = PrimerNombreMadre
        End If

        arParms(11) = New SqlParameter("@SegundoNombreMadre", 22)
        If String.IsNullOrEmpty(SegundoNombreMadre) Then
            arParms(11).Value = System.DBNull.Value
        Else
            arParms(11).Value = SegundoNombreMadre
        End If

        arParms(12) = New SqlParameter("@PrimerApellidoMadre", 22)
        If String.IsNullOrEmpty(PrimerApellidoMadre) Then
            arParms(12).Value = System.DBNull.Value
        Else
            arParms(12).Value = PrimerApellidoMadre
        End If

        arParms(13) = New SqlParameter("@SegundoApellidoMadre", 22)
        If String.IsNullOrEmpty(PrimerApellidoMadre) Then
            arParms(13).Value = System.DBNull.Value
        Else
            arParms(13).Value = SegundoApellidoMadre
        End If

        arParms(14) = New SqlParameter("@grupoEtareo", 8)
        arParms(14).Value = grupoEtareo

        Dim dsPersona As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPersonaBusqueda", arParms)
        Dim personas As New PersonaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsPersona.Tables(0).Rows
            Dim persona As New PersonaEntity()

            persona.per_Consecutivo = CLng(dr("per_Consecutivo"))

            If Not dr("per_TipoId").GetType.ToString.Equals("System.DBNull") Then
                persona.per_TipoId = CStr(dr("per_TipoId"))
            End If

            If Not dr("per_Id").GetType.ToString.Equals("System.DBNull") Then
                persona.per_Id = CStr(dr("per_Id"))
            End If

            If Not dr("per_CertNacVivo").GetType.ToString.Equals("System.DBNull") Then
                persona.per_CertNacVivo = CStr(dr("per_CertNacVivo"))
            End If

            If Not dr("per_CertDefuncion").GetType.ToString.Equals("System.DBNull") Then
                persona.per_CertDefuncion = CStr(dr("per_CertDefuncion"))
            End If

            If Not dr("per_TipoIdM").GetType.ToString.Equals("System.DBNull") Then
                persona.per_TipoIdM = CStr(dr("per_TipoIdM"))
            End If

            If Not dr("per_IdM").GetType.ToString.Equals("System.DBNull") Then
                persona.per_IdM = CStr(dr("per_IdM"))
            End If

            If Not dr("per_NumeroHijoM").GetType.ToString.Equals("System.DBNull") Then
                persona.per_NumeroHijoM = CShort(dr("per_NumeroHijoM"))
            End If

            If Not dr("per_nombre1").GetType.ToString.Equals("System.DBNull") Then
                persona.primerNombre = CStr(dr("per_nombre1"))
            End If

            If Not dr("per_nombre2").GetType.ToString.Equals("System.DBNull") Then
                persona.segundoNombre = CStr(dr("per_nombre2"))
            End If


            If Not dr("per_apellido1").GetType.ToString.Equals("System.DBNull") Then
                persona.primerApellido = CStr(dr("per_apellido1"))
            End If

            If Not dr("per_apellido2").GetType.ToString.Equals("System.DBNull") Then
                persona.segundoApellido = CStr(dr("per_apellido2"))
            End If

            If Not dr("per_FechaNac").GetType.ToString.Equals("System.DBNull") Then
                persona.perFechaNac = CDate(dr("per_FechaNac"))
            End If

            If Not dr("per_parInstitucion").GetType.ToString.Equals("System.DBNull") Then
                persona.per_parInstitucion = CStr(dr("per_parInstitucion"))
            End If

            If Not dr("per_FechaAlm").GetType.ToString.Equals("System.DBNull") Then
                persona.per_FechaAlm = CDate(dr("per_FechaAlm"))
            End If

            If Not dr("per_func").GetType.ToString.Equals("System.DBNull") Then
                persona.per_Func = CStr(dr("per_Func"))
            End If

            If Not dr("per_Institucion").GetType.ToString.Equals("System.DBNull") Then
                persona.per_Institucion = CStr(dr("per_Institucion"))
            End If

            If Not dr("per_Estado").GetType.ToString.Equals("System.DBNull") Then
                persona.per_Estado = CInt(dr("per_Estado"))
            End If

            If Not dr("per_Cni_id").GetType.ToString.Equals("System.DBNull") Then
                persona.cni_id = CInt(dr("per_Cni_id")) 'm 9/20/2013 cni_id por per_causaNoVacuna
            End If

            If Not dr("per_idEtnia").GetType.ToString.Equals("System.DBNull") Then
                persona.etn_idEtnia = CShort(CInt(dr("per_idEtnia")))
            End If

            If Not dr("per_IdGrupoPoblacional").GetType.ToString.Equals("System.DBNull") Then
                persona.gru_IdGrupo = CStr(dr("per_IdGrupoPoblacional"))
            End If
            'Alejandro Muñoz 26/10/2017 - Actualización cambio. Para que no haga homologación, ya que compensar envía ya los datos validados'
            If Not dr("per_IdGenero").GetType.ToString.Equals("System.DBNull") Then 'm 9/20/2013 per_Genero por per_IdGenero
                persona.per_Genero = CStr(dr("per_IdGenero")) 'm 9/20/2013 per_Genero por per_IdGenero

                'If Not dr("per_IdGenero").GetType.ToString.Equals("System.DBNull") Then 'm 9/20/2013 per_Genero por per_IdGenero
                '    If CStr(dr("per_IdGenero")) = "M" Then
                '        persona.per_Genero = "H"
                '    ElseIf CStr(dr("per_IdGenero")) = "F" Then
                '        persona.per_Genero = "M"
                '    End If
                'm 9/20/2013 per_Genero por per_IdGenero
            End If

            If Not dr("per_IdGrupoSanguineo").GetType.ToString.Equals("System.DBNull") Then
                persona.perGrupoSanguineo = CStr(dr("per_IdGrupoSanguineo"))
            End If

            If Not dr("per_IdRh").GetType.ToString.Equals("System.DBNull") Then
                persona.perRh = CStr(dr("per_IdRh"))
            End If

            'If Not dr("EstadoSeguimiento").GetType.ToString.Equals("System.DBNull") Then
            '    persona.ese_Descripcion = CStr(dr("EstadoSeguimiento"))
            'End If
            'i 23/9/2013
            'como no existe los campos de en la vista de la madre a excepción 
            'de tipo[per_TipoIdM] y numero de documento[per_IdM], los busco
            'uno a uno con base a este indice(per_TipoIdM, per_IdM)
            If bGetMadre Then
                If persona.per_TipoIdM IsNot Nothing And persona.per_IdM IsNot Nothing Then
                    Dim PersonasM As PersonaCollection
                    PersonasM = seleccionarPersonaBusquedaAttr(persona.per_TipoIdM, persona.per_IdM, Nothing, Nothing, Nothing, Nothing, Nothing, New Date(),
                                                   Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, 3, False, False)
                    If PersonasM.Count > 0 Then
                        persona.primerNombreM = PersonasM(0).primerNombre
                        persona.segundoNombreM = PersonasM(0).segundoNombre
                        persona.primerApellidoM = PersonasM(0).primerApellido
                        persona.segundoApellidoM = PersonasM(0).segundoApellido
                        persona.per_ConsecutivoM = PersonasM(0).per_Consecutivo
                    End If
                End If
            End If

            If bGetHdocum Then
                Dim strJson As String = "{}"
                persona.hisDocumentJson = seleccionarPersonaIdentificacion(persona.per_Consecutivo)
            End If
            'fin i 23/09/2013


            personas.Add(persona)
        Next
        dsPersona.Dispose()

        insertarLog("seleccionarPersonaBusquedaAttr()", "CORRECTO", arParms)
        insertarLogBD("seleccionarPersonaBusquedaAttr()", "CORRECTO", arParms)

        Return personas
    End Function

#End Region


    Public Function seleccionarTablaDominio(id_Tabla As Short) As TablaDominioCollection Implements IpaiServicioSDS.seleccionarTablaDominio
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim cTablaDominio As New TablaDominioCollection

        Dim arParms() As SqlParameter = New SqlParameter(0) {}
        arParms(0) = New SqlParameter("@id_Tabla", 16)
        arParms(0).Value = id_Tabla

        Dim dsTablaDominio As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarTablasDominio", arParms)

        If Not dsTablaDominio Is Nothing AndAlso dsTablaDominio.Tables(0).Rows.Count > 0 Then
            For Each dr As Data.DataRow In dsTablaDominio.Tables(0).Rows
                Dim eTablaDominio As New TablaDominioEntity
                eTablaDominio.td_id = dr(0).ToString
                eTablaDominio.td_descripcion = dr(1).ToString
                cTablaDominio.Add(eTablaDominio)
            Next
        End If

        Return cTablaDominio
    End Function

    Public Function seleccionarUbicacionPersona(per_Consecutivo As Long) As UbicacionPersonaEntity Implements IpaiServicioSDS.seleccionarUbicacionPersona
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oPersonaUbicacion As New UbicacionPersonaEntity

        Dim arParms() As SqlParameter = New SqlParameter(0) {}
        arParms(0) = New SqlParameter("@per_Consecutivo", 0)
        arParms(0).Value = per_Consecutivo

        Dim dsPersonaUbicacion As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarUbicacionPersona", arParms)
        If Not dsPersonaUbicacion Is Nothing AndAlso dsPersonaUbicacion.Tables(0).Rows.Count > 0 Then
            With dsPersonaUbicacion.Tables(0)
                oPersonaUbicacion.per_Consecutivo = CLng(.Rows(0)("per_Consecutivo").ToString())
                oPersonaUbicacion.dir_Direccion = .Rows(0)("dir_Direccion").ToString()
                oPersonaUbicacion.bar_Id = .Rows(0)("bar_Id").ToString()
                oPersonaUbicacion.bar_Nombre = .Rows(0)("bar_Nombre").ToString()
                If Not .Rows(0)("upz_Id").ToString().Equals(String.Empty) Then
                    oPersonaUbicacion.upz_Id = CInt(.Rows(0)("upz_Id").ToString())
                End If
                oPersonaUbicacion.upz_Nombre = .Rows(0)("upz_Nombre").ToString()
                If Not .Rows(0)("loc_Id").ToString().Equals(String.Empty) Then
                    oPersonaUbicacion.loc_id = CInt(.Rows(0)("loc_Id").ToString())
                End If
                oPersonaUbicacion.loc_Nombre = .Rows(0)("loc_Nombre").ToString()
                oPersonaUbicacion.dir_Codigo_direccion = .Rows(0)("dir_Codigo_direccion").ToString()
                oPersonaUbicacion.dir_CoordenadaX = .Rows(0)("dir_CoordenadaX").ToString()
                oPersonaUbicacion.dir_CoordenadaY = .Rows(0)("dir_CoordenadaY").ToString()
                oPersonaUbicacion.tel_Contacto = .Rows(0)("tel_Contacto").ToString()
                oPersonaUbicacion.tel_Telefono = .Rows(0)("tel_Telefono").ToString()
                oPersonaUbicacion.cor_correo = .Rows(0)("cor_correo").ToString()
                If Not String.IsNullOrEmpty(.Rows(0)("dir_mun_id").ToString()) Then
                    oPersonaUbicacion.dir_mun_id = CInt(.Rows(0)("dir_mun_id").ToString())
                End If
                If Not String.IsNullOrEmpty(.Rows(0)("dir_dep_Id").ToString()) Then
                    oPersonaUbicacion.dir_dep_Id = CInt(.Rows(0)("dir_dep_Id").ToString())
                End If
                oPersonaUbicacion.dir_pais_Id = .Rows(0)("dir_pais_Id").ToString()
                If Not String.IsNullOrEmpty(.Rows(0)("dir_zon_Id").ToString()) Then
                    oPersonaUbicacion.dir_zon_Id = CInt(.Rows(0)("dir_zon_Id").ToString())
                End If
            End With

            insertarLog("seleccionarUbicacionPersona()", "CORRECTO", arParms)
            insertarLogBD("seleccionarUbicacionPersona()", "CORRECTO", arParms)
        End If
        Return oPersonaUbicacion
    End Function

    Public Function seleccionarVacunasPersona(per_Consecutivo As Long) As VacunaCollection Implements IpaiServicioSDS.seleccionarVacunasPersona
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity

        Dim arParms() As SqlParameter = New SqlParameter(0) {}

        arParms(0) = New SqlParameter("@per_Consecutivo", 0)
        arParms(0).Value = per_Consecutivo

        Dim dsVacuna As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarVacunasPersona", arParms)
        Dim vacunas As New VacunaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsVacuna.Tables(0).Rows
            Dim vacuna As New VacunaEntity()
            vacuna.vac_id = CInt(dr("vac_Id"))
            vacuna.Dos_id = CInt(dr("dos_Id"))
            vacuna.vac_FechaVacuna = CDate(dr("vac_FechaVacuna"))

            If Not dr("vac_EdadVacunaAnios").GetType.ToString.Equals("System.DBNull") Then
                vacuna.vac_EdadVacunaAnios = CInt(dr("vac_EdadVacunaAnios"))
            End If
            If Not dr("vac_EdadVacunaMeses").GetType.ToString.Equals("System.DBNull") Then
                vacuna.vac_EdadVacunaMeses = CInt(dr("vac_EdadVacunaMeses"))
            End If
            If Not dr("vac_EdadVacunaDias").GetType.ToString.Equals("System.DBNull") Then
                vacuna.vac_EdadVacunaDias = CInt(dr("vac_EdadVacunaDias"))
            End If

            'i 20/9/2013
            vacuna.per_Consecutivo = per_Consecutivo
            vacuna.pse_Id = If(IsDBNull(dr("pse_Id")), vacuna.pse_Id, CInt(dr("pse_Id")))
            vacuna.pos_Id = If(IsDBNull(dr("pos_Id")), vacuna.pos_Id, CInt(dr("pos_Id")))
            vacuna.com_Id = If(IsDBNull(dr("com_Id")), vacuna.com_Id, CInt(dr("com_Id")))


            vacunas.Add(vacuna)
        Next

        insertarLog("seleccionarVacunasPersona()", "CORRECTO", arParms)
        insertarLogBD("seleccionarVacunasPersona()", "CORRECTO", arParms)
        dsVacuna.Dispose()

        Return vacunas
    End Function

    Public Function seleccionarEstadoContactenos(idCaso As Long) As String Implements IpaiServicioSDS.seleccionarEstadoContactenos
            Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
            Dim cTablaDominio As New TablaDominioCollection

            Dim arParms() As SqlParameter = New SqlParameter(0) {}
            arParms(0) = New SqlParameter("@idCaso", 0)
            arParms(0).Value = idCaso

            Dim estado As String = CStr(SqlHelper.ExecuteScalar(cadena, "pa_SeleccionarContactenosEstadoPorId", arParms))

            insertarLog("seleccionarEstadoContactenos()", "CORRECTO", arParms)

            Return estado
        End Function

        Public Function insertarSeguimientoCohorte(tsc_id As Integer, seg_Numero_Llamada As Integer, seg_Fecha_Llamada As Date, seg_Comunica As Boolean, seg_mnc As Integer, seg_Mensaje As Boolean, res_Id As Integer, mnv_Id As Integer, seg_Observaciones As String, seg_per_consecutivo As Long, seg_Activo As Boolean, seg_FechaSeguimientoPersonal As Date, seg_UserName As String) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarSeguimientoCohorte
            Dim oResultado As New resultadoConsultaEntity
            Dim fValidacion As Boolean = True
            Dim MensajeValidacion As String = String.Empty


            '---------------Aqui se validan los valores en 0 o vacios
            If tsc_id = 0 Then
                oResultado.resultado = False
                oResultado.errores += " La variable tsc_id no puede estar vacía"
                fValidacion = False
            End If
            If seg_Numero_Llamada = 0 Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_Numero_Llamada no puede estar vacía"
                fValidacion = False
            End If

            If seg_Fecha_Llamada = Date.MinValue Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_Fecha_Llamada no puede estar vacía"
                fValidacion = False
            End If

            If String.IsNullOrEmpty(CType(seg_Comunica, String)) Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_Comunica no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(CType(seg_Mensaje, String)) Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_Mensaje no puede estar vacía"
                fValidacion = False
            End If

            If seg_per_consecutivo = 0 Then
                oResultado.resultado = False
                oResultado.errores = "la variable seg_per_consecutivo no puede ser 0"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(CType(seg_Activo, String)) Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_Activo no puede estar vacía"
                fValidacion = False
            End If

            If seg_FechaSeguimientoPersonal = Date.MinValue Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_FechaSeguimientoPersonal no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(seg_UserName) Then
                oResultado.resultado = False
                oResultado.errores += " La variable seg_UserName no puede estar vacía"
                fValidacion = False
            End If

            '----------------- Aqui se valida contra las tablas de dominio
            If Not validarTablasDominio(33, tsc_id.ToString, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(30, seg_mnc.ToString, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(34, mnv_Id.ToString, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(32, seg_UserName, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If

            '---------------------------

            If fValidacion Then
                Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
                Dim arParms() As SqlParameter = New SqlParameter(12) {}

                arParms(0) = New SqlParameter("@tsc_id", 8)
                arParms(0).Value = tsc_id

                arParms(1) = New SqlParameter("@seg_Numero_Llamada", 22)
                If seg_Numero_Llamada = 0 Then
                    arParms(1).Value = System.DBNull.Value
                Else
                    arParms(1).Value = seg_Numero_Llamada
                End If


                arParms(2) = New SqlParameter("@seg_Fecha_Llamada", 31)
                If seg_Fecha_Llamada = Date.MinValue Then
                    arParms(2).Value = System.DBNull.Value
                Else
                    arParms(2).Value = seg_Fecha_Llamada
                End If

                arParms(3) = New SqlParameter("@seg_Comunica", 2)
                arParms(3).Value = seg_Comunica

                arParms(4) = New SqlParameter("@seg_mnc", 8)
                If seg_mnc = 0 Then
                    arParms(4).Value = System.DBNull.Value
                Else
                    arParms(4).Value = seg_mnc
                End If

                arParms(5) = New SqlParameter("@seg_Mensaje", 2)
                arParms(5).Value = seg_Mensaje

                arParms(6) = New SqlParameter("@res_Id", 8)
                If res_Id = 0 Then
                    arParms(6).Value = System.DBNull.Value
                Else
                    arParms(6).Value = res_Id
                End If

                arParms(7) = New SqlParameter("@mnv_Id", 8)
                If mnv_Id = 0 Then
                    arParms(7).Value = System.DBNull.Value
                Else
                    arParms(7).Value = mnv_Id
                End If

                arParms(8) = New SqlParameter("@seg_Observaciones", 22)
                If String.IsNullOrEmpty(seg_Observaciones) Then
                    arParms(8).Value = System.DBNull.Value
                Else
                    arParms(8).Value = seg_Observaciones
                End If

                arParms(9) = New SqlParameter("@seg_per_consecutivo", 0)
                arParms(9).Value = seg_per_consecutivo

                arParms(10) = New SqlParameter("@seg_Activo", 2)
                arParms(10).Value = seg_Activo

                arParms(11) = New SqlParameter("@seg_FechaSeguimientoPersonal", 31)
                If seg_FechaSeguimientoPersonal = Date.MinValue Then
                    arParms(11).Value = System.DBNull.Value
                Else
                    arParms(11).Value = seg_FechaSeguimientoPersonal
                End If

                arParms(12) = New SqlParameter("@seg_UserName", 12)
                arParms(12).Value = seg_UserName

                Try
                    'If SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarSeguimiento_getId", arParms) > 0 Then
                    Dim idResult = SqlHelper.ExecuteScalar(cadena, "pa_InsertarSeguimiento_getIdWS", arParms)
                    idResult = If(IsDBNull(idResult), -1, CLng(idResult))

                    If CLng(idResult) > 0 Then
                        oResultado.resultado = True
                        oResultado.errores = String.Empty
                        oResultado.consecutivo = CStr(idResult)

                        insertarLog("insertarSeguimientoCohorte()", "CORRECTO", arParms)
                    Else
                        oResultado.resultado = False
                        oResultado.errores = "El registro no fué insertado."

                        insertarLog("insertarSeguimientoCohorte()", "No realizo la insercion", arParms)
                    End If

                Catch ex As SqlException
                    Dim errorMessage As String = ex.Message
                    Dim errorCode As Integer = ex.ErrorCode
                    oResultado.resultado = False
                    oResultado.errores = errorMessage

                    insertarLog("insertarSeguimientoCohorte()", ex.Message, arParms)
                End Try
            End If
            Return oResultado
        End Function

        Public Function seleccionarTablasConNovedades() As TablaDominioCollection Implements IpaiServicioSDS.seleccionarTablasConNovedades
            Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
            Dim cTablaDominio As New TablaDominioCollection

            Dim dsTablasDominio As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarTablasDominioModificadas")
            If Not dsTablasDominio Is Nothing Then
                For Each dr As Data.DataRow In dsTablasDominio.Tables(0).Rows
                    Dim eTablaDominio As New TablaDominioEntity
                    eTablaDominio.td_id = dr(0).ToString
                    'eTablaDominio.td_descripcion = dr(1).ToString 'e 9/20/2013
                    cTablaDominio.Add(eTablaDominio)
                Next
            End If

            Return cTablaDominio
        End Function

        Public Function insertarAfiliacionPersona(per_Consecutivo As Long,
                                                  ase_id As String,
                                                  reg_Id As Integer,
                                                  tia_id As Integer) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarAfiliacionPersona

            Dim oResultado As New resultadoConsultaEntity
            Dim fValidacion As Boolean = True
            Dim MensajeValidacion As String = String.Empty

            '---------------Aqui se validan los valores en 0
            If per_Consecutivo = 0 Then
                oResultado.resultado = False
                oResultado.errores = " La variable per_Consecutivo no puede ser 0"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(ase_id) Then
                oResultado.resultado = False
                oResultado.errores += " La variable ase_id no puede estar vacía"
                fValidacion = False
            End If
            If tia_id = 0 Then
                oResultado.resultado = False
                oResultado.errores += " La variable tia_id no puede ser 0"
                fValidacion = False
            End If
            '----------------- Aqui se valida contra las tablas de dominio
            If Not validarTablasDominio(1, ase_id, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(2, tia_id.ToString, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If

            If fValidacion Then
                Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
                Dim arParms() As SqlParameter = New SqlParameter(3) {}

                arParms(0) = New SqlParameter("@per_Consecutivo", 0)
                arParms(0).Value = per_Consecutivo

                arParms(1) = New SqlParameter("@ase_id", 22)
                arParms(1).Value = ase_id

                arParms(2) = New SqlParameter("@reg_Id", 0)
                arParms(2).Value = reg_Id

                arParms(3) = New SqlParameter("@tia_id", 0)
                arParms(3).Value = tia_id

                Try
                    If SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarAseguradoraPersonaWS", arParms) > 0 Then
                        oResultado.resultado = True
                        oResultado.errores = String.Empty
                        oResultado.consecutivo = CStr(per_Consecutivo)

                        insertarLog("insertarAfiliacionPersona()", "CORRECTO", arParms)
                        insertarLogBD("insertarAfiliacionPersona()", "CORRECTO", arParms)
                    Else
                        oResultado.resultado = False
                        oResultado.errores = "Registro no identificado."

                        insertarLog("insertarAfiliacionPersona()", "No realizo la insercion", arParms)
                        insertarLogBD("insertarAfiliacionPersona()", "No realizo la insercion", arParms)
                    End If

                Catch ex As SqlException
                    Dim errorMessage As String = ex.Message
                    Dim errorCode As Integer = ex.ErrorCode
                    oResultado.resultado = False
                    oResultado.errores = errorMessage

                    insertarLog("insertarAfiliacionPersona()", ex.Message, arParms)
                    insertarLogBD("insertarAfiliacionPersona()", ex.Message, arParms)
                End Try
            End If

            Return oResultado
        End Function

        Public Function insertarUbicacionPersona(per_Consecutivo As Long, dir_Direccion As String, dir_Barrio As String, dir_Codigo_direccion As String, dir_Upz As String, dir_CoordenadaX As String, dir_CoordenadaY As String, dir_Localidad As Integer, dir_Estrato As String, tel_Telefono As String, tel_Contacto As String, cor_correo As String, dir_mun_id As Integer, dir_dep_Id As Integer, dir_pais_Id As String, dir_zon_Id As Integer) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarUbicacionPersona

            Dim oResultado As New resultadoConsultaEntity
            Dim fValidacion As Boolean = True

            '---------------Aqui se validan los valores en 0 o vacios
            If per_Consecutivo = 0 Then
                oResultado.resultado = False
                oResultado.errores = "la variable per_Consecutivo no puede ser 0"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(dir_Direccion) Then
                oResultado.resultado = False
                oResultado.errores += " La variable dir_Direccion no puede estar vacía"
                fValidacion = False
            End If

            If fValidacion Then
                Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
                Dim arParms() As SqlParameter = New SqlParameter(15) {}

                arParms(0) = New SqlParameter("@per_Consecutivo", 0)
                arParms(0).Value = per_Consecutivo

                arParms(1) = New SqlParameter("@dir_Direccion", 22)
                If String.IsNullOrEmpty(dir_Direccion) Then
                    arParms(1).Value = System.DBNull.Value
                Else
                    arParms(1).Value = dir_Direccion
                End If

                arParms(2) = New SqlParameter("@dir_Barrio", 22)
                If String.IsNullOrEmpty(dir_Barrio) Then
                    arParms(2).Value = System.DBNull.Value
                Else
                    arParms(2).Value = dir_Barrio
                End If

                arParms(3) = New SqlParameter("@dir_codigo_direccion", 22)
                If String.IsNullOrEmpty(dir_Codigo_direccion) Then
                    arParms(3).Value = System.DBNull.Value
                Else
                    arParms(3).Value = dir_Codigo_direccion
                End If

                arParms(4) = New SqlParameter("@dir_Upz", 22)
                If String.IsNullOrEmpty(dir_Upz) OrElse dir_Upz.Equals("0") Then
                    arParms(4).Value = System.DBNull.Value
                Else
                    arParms(4).Value = dir_Upz
                End If

                arParms(5) = New SqlParameter("@dir_CoordenadaX", 22)
                If String.IsNullOrEmpty(dir_CoordenadaX) Then
                    arParms(5).Value = System.DBNull.Value
                Else
                    arParms(5).Value = dir_CoordenadaX
                End If

                arParms(6) = New SqlParameter("@dir_CoordenadaY", 22)
                If String.IsNullOrEmpty(dir_CoordenadaY) Then
                    arParms(6).Value = System.DBNull.Value
                Else
                    arParms(6).Value = dir_CoordenadaY
                End If

                arParms(7) = New SqlParameter("@dir_Localidad", 8)
                If dir_Localidad = 0 Then
                    arParms(7).Value = System.DBNull.Value
                Else
                    arParms(7).Value = dir_Localidad
                End If

                arParms(8) = New SqlParameter("@dir_Estrato", 8)
                If String.IsNullOrEmpty(dir_Estrato) OrElse dir_Upz.Equals("0") Then
                    arParms(8).Value = System.DBNull.Value
                Else
                    arParms(8).Value = dir_Estrato
                End If


                arParms(9) = New SqlParameter("@tel_Telefono", 22)
                If String.IsNullOrEmpty(tel_Telefono) Then
                    arParms(9).Value = System.DBNull.Value
                Else
                    arParms(9).Value = tel_Telefono
                End If

                arParms(10) = New SqlParameter("@tel_Contacto", 22)
                If String.IsNullOrEmpty(tel_Contacto) Then
                    arParms(10).Value = System.DBNull.Value
                Else
                    arParms(10).Value = tel_Contacto
                End If

                arParms(11) = New SqlParameter("@cor_correo", 22)
                If String.IsNullOrEmpty(cor_correo) Then
                    arParms(11).Value = System.DBNull.Value
                Else
                    arParms(11).Value = cor_correo
                End If

                arParms(12) = New SqlParameter("@dir_mun_id", 8)
                If String.IsNullOrEmpty(CStr(dir_mun_id)) OrElse dir_mun_id = 0 Then
                    arParms(12).Value = System.DBNull.Value
                Else
                    arParms(12).Value = dir_mun_id
                End If

                arParms(13) = New SqlParameter("@dir_dep_Id", 8)
                If String.IsNullOrEmpty(CStr(dir_dep_Id)) OrElse dir_dep_Id = 0 Then
                    arParms(13).Value = System.DBNull.Value
                Else
                    arParms(13).Value = dir_dep_Id
                End If

                arParms(14) = New SqlParameter("@dir_pais_Id", 3)
                If String.IsNullOrEmpty(dir_pais_Id) Then
                    arParms(14).Value = System.DBNull.Value
                Else
                    arParms(14).Value = dir_pais_Id
                End If

                arParms(15) = New SqlParameter("@dir_zon_Id", 20)
                If String.IsNullOrEmpty(CStr(dir_zon_Id)) OrElse dir_zon_Id = 0 Then
                    arParms(15).Value = System.DBNull.Value
                Else
                    arParms(15).Value = dir_zon_Id
                End If

                Try
                    If SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarUbicacionPersonaWS", arParms) > 0 Then
                        oResultado.resultado = True
                        oResultado.errores = String.Empty
                        oResultado.consecutivo = CStr(per_Consecutivo)

                        insertarLog("insertarUbicacionPersona()", "CORRECTO", arParms)
                        insertarLogBD("insertarUbicacionPersona()", "CORRECTO", arParms)
                    Else
                        oResultado.resultado = False
                        oResultado.errores = "Registro no identificado."

                        insertarLog("insertarUbicacionPersona()", "No se inserto ninguna ubicacion", arParms)
                        insertarLogBD("insertarUbicacionPersona()", "No se inserto ninguna ubicacion", arParms)
                    End If
                Catch ex As SqlException
                    Dim errorMessage As String = ex.Message
                    Dim errorCode As Integer = ex.ErrorCode
                    oResultado.resultado = False
                    oResultado.errores = errorMessage

                    insertarLog("insertarUbicacionPersona()", ex.Message, arParms)
                    insertarLogBD("insertarUbicacionPersona()", ex.Message, arParms)
                End Try
            End If

            Return oResultado
        End Function

        Public Function seleccionarSeguimientoCohorte(mesNacimiento As Integer, anioNacimiento As Integer) As PersonaCohorteCollection Implements IpaiServicioSDS.seleccionarSeguimientoCohorte
            Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
            Dim cPersonaCohorte As New PersonaCohorteCollection

            Dim arParms() As SqlParameter = New SqlParameter(1) {}
            arParms(0) = New SqlParameter("@mesNacimiento", 16)
            arParms(0).Value = mesNacimiento

            arParms(1) = New SqlParameter("@anioNacimiento", 16)
            arParms(1).Value = anioNacimiento

            Dim dsPersonaCohorte As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPersonaCohorteCompensar", arParms)

            If Not dsPersonaCohorte Is Nothing AndAlso dsPersonaCohorte.Tables(0).Rows.Count > 0 Then
                For Each dr As Data.DataRow In dsPersonaCohorte.Tables(0).Rows
                    Dim ePersonaCohorte As New PersonaCohorteEntity
                    ePersonaCohorte.per_Consecutivo = CLng(dr("per_Consecutivo")) 'notnull
                    ePersonaCohorte.per_TipoId = If(IsDBNull(dr("per_TipoId")), ePersonaCohorte.per_TipoId, CStr(dr("per_TipoId"))) 'm 9/20/2013
                    ePersonaCohorte.per_Id = If(IsDBNull(dr("per_Id")), ePersonaCohorte.per_Id, CStr(dr("per_Id"))) 'm 9/20/2013
                    ePersonaCohorte.nombresApellidos = If(IsDBNull(dr("nombresApellidos")), ePersonaCohorte.nombresApellidos, CStr(dr("nombresApellidos"))) 'm 9/20/2013
                    ePersonaCohorte.per_FechaNac = If(IsDBNull(dr("per_FechaNac")), ePersonaCohorte.per_FechaNac, CDate(dr("per_FechaNac"))) 'm 9/20/2013
                    ePersonaCohorte.gen_Nombre = If(IsDBNull(dr("gen_Nombre")), ePersonaCohorte.gen_Nombre, CStr(dr("gen_Nombre"))) 'm 9/20/2013 notnull
                    ePersonaCohorte.reg_Nombre = If(IsDBNull(dr("reg_Nombre")), ePersonaCohorte.reg_Nombre, CStr(dr("reg_Nombre"))) 'm 9/20/2013
                    ePersonaCohorte.ase_Nombre = If(IsDBNull(dr("ase_Nombre")), ePersonaCohorte.ase_Nombre, CStr(dr("ase_Nombre"))) 'm 9/20/2013
                    ePersonaCohorte.cor_correo = If(IsDBNull(dr("cor_correo")), ePersonaCohorte.cor_correo, CStr(dr("cor_correo"))) 'm 9/20/2013
                    ePersonaCohorte.dir_Direccion = If(IsDBNull(dr("dir_Direccion")), ePersonaCohorte.dir_Direccion, CStr(dr("dir_Direccion"))) 'm 9/20/2013
                    ePersonaCohorte.dir_Localidad = If(IsDBNull(dr("dir_Localidad")), ePersonaCohorte.dir_Localidad, CStr(dr("dir_Localidad"))) 'm 9/20/2013
                    ePersonaCohorte.est_Descripcion = If(IsDBNull(dr("est_Descripcion")), ePersonaCohorte.est_Descripcion, CStr(dr("est_Descripcion"))) 'm 9/20/2013 notnull
                    ePersonaCohorte.ese_Descripcion = If(IsDBNull(dr("ese_Descripcion")), ePersonaCohorte.ese_Descripcion, CStr(dr("ese_Descripcion")))  'm 9/20/2013
                    ePersonaCohorte.ope_Descripcion = If(IsDBNull(dr("ope_Descripcion")), ePersonaCohorte.ope_Descripcion, CStr(dr("ope_Descripcion")))  'm 9/20/2013
                    ePersonaCohorte.vacunasPendientes = If(IsDBNull(dr("vacunasPendientes")), ePersonaCohorte.vacunasPendientes, CStr(dr("vacunasPendientes"))) 'm 9/20/2013


                    cPersonaCohorte.Add(ePersonaCohorte)
                Next
            End If

            insertarLog("seleccionarSeguimientoCohorte()", "CORRECTO", arParms)

            Return cPersonaCohorte
        End Function

        Public Function insertarPersonaVacuna(per_Id As String,
                                        per_TipoId As String,
                                        per_TipoIdM As String,
                                        per_IdM As String,
                                        per_NumeroHijoM As Short,
                                        primerNombre As String,
                                        segundoNombre As String,
                                        primerApellido As String,
                                        segundoApellido As String,
                                        primerNombreM As String,
                                        segundoNombreM As String,
                                        primerApellidoM As String,
                                        segundoApellidoM As String,
                                        per_ParInstitucion As String,
                                        cni_id As Integer,
                                        etn_idEtnia As Short,
                                        gru_IdGrupo As String,
                                        per_Genero As String,
                                        perGrupoSanguineo As String,
                                        perRh As String,
                                        cdm_idCondicion As Integer,
                                        perFechaNac As Date,
                                        tel_Telefono As String,
                                        tel_Contacto As String,
                                        dir_Direccion As String,
                                        bar_Id As String,
                                        upz_Id As Integer,
                                        loc_id As Integer,
                                        dir_mun_id As Integer,
                                        dir_dep_Id As Integer,
                                        dir_pais_Id As String,
                                        dir_zon_Id As Integer,
                                        cor_correo As String,
                                        ase_id As String,
                                        tia_id As Integer,
                                        cnv_Id As Integer,
                                        vac_IdNoAplicada As Integer,
                                        dos_idNoAplicada As Integer,
                                        pes_Peso As String,
                                        vac_Id As Integer,
                                        dos_Id As Integer,
                                        pse_Id As Integer,
                                        com_Id As Integer,
                                        cam_id As Integer,
                                        vac_FechaVacuna As Date,
                                        vac_actualizacion As Boolean,
                                        ins_IdVacuna As String,
                                        vac_Lote As String,
                                        pos_Id As Integer,
                                        per_Func As String,
                                        per_Institucion As String) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarPersonaVacuna

            Dim eResultadoConsulta As New resultadoConsultaEntity
            Dim ePersona As New PersonaEntity

            Dim fValidacion As Boolean = True
            Dim MensajeValidacion As String = String.Empty

            '---------------Aqui se validan los valores en 0 o vacios

            If String.IsNullOrEmpty(per_Id) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_Id no puede estar vacía"
                fValidacion = False
            End If

            If String.IsNullOrEmpty(per_TipoId) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_TipoId no puede estar vacía"
                fValidacion = False
            End If

            If String.IsNullOrEmpty(per_TipoIdM) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_TipoIdM no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(per_IdM) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_IdM no puede estar vacía"
                fValidacion = False
            End If
            If per_NumeroHijoM = 0 Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores = "la variable per_NumeroHijoM no puede ser 0"
                fValidacion = False
            End If

            If String.IsNullOrEmpty(primerNombre) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable primerNombre  no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(primerApellido) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable primerApellido  no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(primerNombreM) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable primerNombreM  no puede estar vacía"
                fValidacion = False
            End If

            If String.IsNullOrEmpty(primerApellidoM) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable primerApellidoM  no puede estar vacía"
                fValidacion = False
            End If
            If perFechaNac = Date.MinValue Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable perFechaNac no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(dir_Direccion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable dir_Direccion no puede estar vacía"
                fValidacion = False
            End If
            If tia_id = 0 Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable tia_id no puede estar vacía"
                fValidacion = False
            End If
            If vac_FechaVacuna = Date.MinValue Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable vac_FechaVacuna no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(ins_IdVacuna) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable ins_IdVacuna no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(vac_Lote) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable vac_Lote no puede estar vacía"
                fValidacion = False
            End If
            If pos_Id = 0 Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable pos_Id no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(per_Func) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_Func no puede estar vacía"
                fValidacion = False
            End If
            If String.IsNullOrEmpty(per_Institucion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += " La variable per_Institucion no puede estar vacía"
                fValidacion = False
            End If


            '----------------- Aqui se valida contra las tablas de dominio
            If Not validarTablasDominio(14, per_TipoId, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores = MensajeValidacion
            End If
            If Not validarTablasDominio(14, per_TipoIdM, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(15, per_Institucion, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(17, cni_id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(18, etn_idEtnia.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(19, gru_IdGrupo, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(20, per_Genero, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If

            If Not validarTablasDominio(21, perGrupoSanguineo, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(22, perRh, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(23, cdm_idCondicion.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(1, ase_id, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(2, tia_id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(4, vac_IdNoAplicada.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(5, dos_idNoAplicada.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(35, pes_Peso.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(4, vac_Id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(5, dos_Id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(24, pse_Id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(25, com_Id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(26, cam_id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(27, pos_Id.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(32, per_Func.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If
            If Not validarTablasDominio(15, per_Institucion.ToString, MensajeValidacion) Then
                eResultadoConsulta.resultado = False
                eResultadoConsulta.errores += MensajeValidacion
            End If

            If fValidacion Then
                '----------------------------------------------------------------------------------------
                ' Verificar si el usuario existe en la base de datos
                '----------------------------------------------------------------------------------------
                If seleccionarPersonaIdContar(per_Id, per_TipoId) > 0 AndAlso Not String.IsNullOrEmpty(per_Id) Then
                    eResultadoConsulta.consecutivo = String.Empty
                    eResultadoConsulta.errores = "La persona ingresada ya existe en el aplicativo, debe seleccionarla en la opción de búsqueda."
                    eResultadoConsulta.resultado = False

                    insertarLog("insertarPersonaVacuna", "La persona ingresada ya existe en el aplicativo, debe seleccionarla en la opción de búsqueda.", Nothing)
                    insertarLogBD("insertarPersonaVacuna", "La persona ingresada ya existe en el aplicativo, debe seleccionarla en la opción de búsqueda.", Nothing)

                    Return eResultadoConsulta
                    Exit Function
                Else

                    '----------------------------------------------------------------------------------------
                    ' Cargar los datos de la persona que fueron digitados en la pantalla al objeto oPersonaE
                    '----------------------------------------------------------------------------------------            
                    ePersona.per_NumeroHijoM = 0

                    '----------------------------------------------------------
                    ' Datos de la Madre        
                    '----------------------------------------------------------
                    ePersona.per_TipoIdM = per_TipoIdM
                    ePersona.per_IdM = per_IdM
                    ePersona.per_NumeroHijoM = per_NumeroHijoM
                    ePersona.primerApellidoM = primerApellidoM.ToUpper
                    ePersona.segundoApellidoM = segundoApellidoM.ToUpper
                    ePersona.primerNombreM = primerNombreM.ToUpper
                    ePersona.segundoNombreM = segundoNombreM.ToUpper

                    '----------------------------------------------------------
                    ' Datos del vacunado
                    '----------------------------------------------------------
                    ePersona.cni_id = cni_id
                    ePersona.per_TipoId = per_TipoId
                    ePersona.per_Id = per_Id
                    If ePersona.per_TipoId.Equals("CN") Then
                        If String.IsNullOrEmpty(ePersona.per_Id) Then
                            ePersona.per_CertNacVivo = "0"
                        Else
                            ePersona.per_CertNacVivo = ePersona.per_Id
                        End If
                    Else
                        ePersona.per_CertNacVivo = "0"
                    End If
                    ePersona.primerApellido = primerApellido.ToUpper
                    ePersona.segundoApellido = segundoApellido.ToUpper
                    ePersona.primerNombre = primerNombre.ToUpper
                    ePersona.segundoNombre = segundoNombre.ToUpper
                    ePersona.perFechaNac = perFechaNac
                    ePersona.per_Genero = per_Genero.ToUpper
                    ePersona.perGrupoSanguineo = perGrupoSanguineo
                    ePersona.perRh = perRh
                    ePersona.etn_idEtnia = etn_idEtnia
                    ePersona.gru_IdGrupo = gru_IdGrupo

                    '----------------------------------------------------------
                    ' Otros datos del objeto persona
                    '----------------------------------------------------------                                
                    ePersona.per_Institucion = per_Institucion
                    ePersona.per_Func = per_Func
                    ePersona.per_Estado = 2 ' Usuario en estado activo 
                    ePersona.per_FechaAlm = Date.Now

                    '----------------------------------------------------------------------------------------
                    ' Realizar la validacion de la fecha de nacimiento
                    '----------------------------------------------------------------------------------------
                    If (ePersona.perFechaNac >= fechaBase) AndAlso
                       CDate(ePersona.perFechaNac) <= Date.Now Then
                        If per_Id <> per_IdM Then

                            '----------------------------------------------------------------------------------------
                            ' Se pregunta si la institucion del usuario esta vacia, de ser asi no se puede actualizar
                            '----------------------------------------------------------------------------------------
                            If Not String.IsNullOrEmpty(ePersona.per_Institucion) Then
                                '----------------------------------------------------------------------------------------
                                ' Insertar los datos básicos de una persona
                                '----------------------------------------------------------------------------------------
                                Dim oResultado As resultadoConsultaEntity
                                oResultado = InsertarPersona(ePersona.per_TipoId,
                                                         ePersona.per_Id,
                                                         CLng(ePersona.per_CertNacVivo),
                                                         ePersona.per_CertDefuncion,
                                                         ePersona.per_TipoIdM,
                                                         ePersona.per_IdM,
                                                         ePersona.per_NumeroHijoM,
                                                         ePersona.primerNombreM,
                                                         ePersona.segundoNombreM,
                                                         ePersona.primerApellidoM,
                                                         ePersona.segundoApellidoM,
                                                         ePersona.primerNombre,
                                                         ePersona.segundoNombre,
                                                         ePersona.primerApellido,
                                                         ePersona.segundoApellido,
                                                         ePersona.perFechaNac,
                                                         ePersona.per_Func,
                                                         ePersona.per_Institucion,
                                                         ePersona.per_Estado,
                                                         ePersona.cni_id,
                                                         ePersona.etn_idEtnia,
                                                         ePersona.gru_IdGrupo,
                                                         ePersona.per_Genero,
                                                         ePersona.perGrupoSanguineo,
                                                         ePersona.perRh)

                                '--------------------------------------------------------------------------
                                ' Verificar si los datos de la persona fueron guardados para continuar..
                                '--------------------------------------------------------------------------
                                If oResultado.resultado And CLng(oResultado.consecutivo) <> 0 Then
                                    ePersona.per_Consecutivo = CLng(oResultado.consecutivo)
                                    '--------------------------------------------------------------------
                                    ' Insertar datos de peso de la persona (Solo recien nacidos)
                                    '--------------------------------------------------------------------
                                    If Not pes_Peso.Equals(String.Empty) Then
                                        ePersona.pes_Peso = pes_Peso
                                        InsertarPesoPersona(ePersona.per_Consecutivo,
                                                            CInt(ePersona.pes_Peso))
                                    End If
                                    '--------------------------------------------------------------------
                                    ' Insertar datos de afiliacion
                                    '--------------------------------------------------------------------
                                    Dim ePersonaAfiliacion As New PersonaAfiliacionEntity
                                    ePersonaAfiliacion.per_Consecutivo = ePersona.per_Consecutivo
                                    ePersonaAfiliacion.ase_id = ase_id
                                    ePersonaAfiliacion.tia_id = tia_id

                                    'TODO: Validar el origen del regimen al ingreso de una persona
                                    insertarAfiliacionPersona(ePersonaAfiliacion.per_Consecutivo,
                                                              ePersonaAfiliacion.ase_id,
                                                              1,
                                                              ePersonaAfiliacion.tia_id)

                                    '--------------------------------------------------------------------
                                    ' Insertar datos Ubicacion               
                                    '--------------------------------------------------------------------
                                    Dim eUbicacion As New UbicacionPersonaEntity
                                    eUbicacion.per_Consecutivo = ePersona.per_Consecutivo
                                    eUbicacion.dir_Direccion = dir_Direccion
                                    eUbicacion.bar_Id = bar_Id
                                    eUbicacion.Fecha = Date.Now
                                    eUbicacion.Activo = CBool(1)
                                    eUbicacion.loc_id = loc_id
                                    eUbicacion.tel_Telefono = tel_Telefono
                                    eUbicacion.tel_Contacto = tel_Contacto
                                    eUbicacion.cor_correo = cor_correo
                                    eUbicacion.dir_mun_id = dir_mun_id
                                    eUbicacion.dir_dep_Id = dir_dep_Id
                                    eUbicacion.dir_pais_Id = dir_pais_Id
                                    eUbicacion.dir_zon_Id = dir_zon_Id
                                    eUbicacion.upz_Id = upz_Id
                                    'eUbicacion.dir_Codigo_direccion = hdfCodDireccion.Value                            
                                    'eUbicacion.dir_CoordenadaX = hdfCoordX.Value
                                    'eUbicacion.dir_CoordenadaY = hdfCoorY.Value                            
                                    'eUbicacion.dir_Estrato = hdfEstrato.Value                            

                                    insertarUbicacionPersona(eUbicacion.per_Consecutivo,
                                                            eUbicacion.dir_Direccion,
                                                            eUbicacion.bar_Id,
                                                            eUbicacion.dir_Codigo_direccion,
                                                            CStr(eUbicacion.upz_Id),
                                                            eUbicacion.dir_CoordenadaX,
                                                            eUbicacion.dir_CoordenadaY,
                                                            eUbicacion.loc_id,
                                                            eUbicacion.dir_Estrato,
                                                            eUbicacion.tel_Telefono,
                                                            eUbicacion.tel_Contacto,
                                                            eUbicacion.cor_correo,
                                                            eUbicacion.dir_mun_id,
                                                            eUbicacion.dir_dep_Id,
                                                            eUbicacion.dir_pais_Id,
                                                            eUbicacion.dir_zon_Id
                                                            )

                                    '--------------------------------------------------------------------
                                    ' InsertarCuestionarioPersona
                                    '--------------------------------------------------------------------
                                    'InsertarCuestionario(opersonaE.per_Consecutivo, opersonaE.per_Institucion)

                                    '--------------------------------------------------------------------
                                    ' Actualizar la institucion de parto si es necesario
                                    '--------------------------------------------------------------------
                                    ActualizarInstitucionParto(ePersona.per_Consecutivo, per_ParInstitucion)

                                    '--------------------------------------------------------------------
                                    ' Actualizar la institucion de parto si es necesario
                                    '--------------------------------------------------------------------
                                    InsertarCausasNoVacunaPersona(cnv_Id, vac_IdNoAplicada, ePersona.per_Consecutivo, dos_idNoAplicada)

                                    '--------------------------------------------------------------------
                                    ' Insertar vacuna persona
                                    '--------------------------------------------------------------------
                                    Dim eVacunaPersona As New VacunaPersonaEntity
                                    eVacunaPersona.vac_FechaVacuna = vac_FechaVacuna
                                    '-------------------------------------------------------------------
                                    ' Verificar la fecha de vacunacion
                                    '-------------------------------------------------------------------
                                    If eVacunaPersona.vac_FechaVacuna >= fechaBase AndAlso
                                       eVacunaPersona.vac_FechaVacuna <= Date.Now AndAlso
                                       eVacunaPersona.vac_FechaVacuna >= ePersona.perFechaNac Then

                                        '------------------------------------------------------------------------------------------------------
                                        ' Se valida que todo lo indispensable se haya seleccionado. Vacuna, Presentación y Presentación Vacuna.
                                        '------------------------------------------------------------------------------------------------------                                                                
                                        eVacunaPersona.per_Consecutivo = ePersona.per_Consecutivo
                                        eVacunaPersona.fun_idFunc = per_Func
                                        eVacunaPersona.cam_id = cam_id
                                        eVacunaPersona.vac_actualizacion = vac_actualizacion
                                        eVacunaPersona.cdm_idCondicion = cdm_idCondicion
                                        eVacunaPersona.vac_Id = vac_Id
                                        eVacunaPersona.dos_Id = dos_Id
                                        eVacunaPersona.pse_Id = pse_Id
                                        eVacunaPersona.ins_Id = ins_IdVacuna
                                        eVacunaPersona.com_Id = com_Id
                                        eVacunaPersona.pos_Id = pos_Id
                                        eVacunaPersona.vac_Lote = vac_Lote
                                        '-------------------------------------------------------------------
                                        ' Consultar la institucion a la cual pertenece el usuario del sistema.
                                        '-------------------------------------------------------------------                                          
                                        eVacunaPersona.ins_Id = per_Institucion

                                        '----------------------------------------------------------------
                                        ' Calculo de la edad al momento de la vacuna
                                        '----------------------------------------------------------------     

                                        calculoEdad(ePersona.perFechaNac, eVacunaPersona.vac_FechaVacuna, eVacunaPersona.vac_EdadVacunaAnios, eVacunaPersona.vac_EdadVacunaMeses, eVacunaPersona.vac_EdadVacunaDias)
                                        eVacunaPersona.vac_EdadVacunaTotalDias = (eVacunaPersona.vac_EdadVacunaAnios * 12 * 30) + (eVacunaPersona.vac_EdadVacunaMeses * 30) + eVacunaPersona.vac_EdadVacunaDias
                                        '----------------------------------------------------------------
                                        ' Guardar vacuna
                                        '----------------------------------------------------------------                                  
                                        InsertarVacunaPersona(eVacunaPersona.per_Consecutivo,
                                                                eVacunaPersona.vac_Id,
                                                                eVacunaPersona.dos_Id,
                                                                eVacunaPersona.pse_Id,
                                                                eVacunaPersona.cam_id,
                                                                eVacunaPersona.vac_FechaVacuna,
                                                                eVacunaPersona.vac_actualizacion,
                                                                eVacunaPersona.ins_Id,
                                                                eVacunaPersona.fun_idFunc,
                                                                eVacunaPersona.com_Id,
                                                                eVacunaPersona.pos_Id,
                                                                eVacunaPersona.vac_Lote,
                                                                eVacunaPersona.vac_EdadVacunaAnios,
                                                                eVacunaPersona.vac_EdadVacunaMeses,
                                                                eVacunaPersona.vac_EdadVacunaDias,
                                                                eVacunaPersona.vac_EdadVacunaTotalDias,
                                                                eVacunaPersona.cdm_idCondicion)

                                        '----------------------------------------------------------------
                                        'Actualizar el esquema de vacunacion (Seguimiento)
                                        '----------------------------------------------------------------                                
                                        ActualizarEsquemaSeguimientoPersona(ePersona.per_Consecutivo)

                                        '--------------------------------------------------------------------
                                        ' Mostrar mensajes de guardado correcto               
                                        '--------------------------------------------------------------------                            
                                        eResultadoConsulta.consecutivo = String.Empty
                                        eResultadoConsulta.errores = "El registro se ha almacenado correctamente."
                                        eResultadoConsulta.resultado = True
                                        Return eResultadoConsulta
                                        Exit Function
                                    Else
                                        eResultadoConsulta.consecutivo = String.Empty
                                        eResultadoConsulta.errores = "La fecha de vacunación es inconsistente, verifiquela."
                                        eResultadoConsulta.resultado = False

                                        insertarLog("insertarPersonaVacuna", "La fecha de vacunación es inconsistente, verifiquela.", Nothing)
                                        insertarLogBD("insertarPersonaVacuna", "La fecha de vacunación es inconsistente, verifiquela.", Nothing)

                                        Return eResultadoConsulta
                                        Exit Function
                                    End If
                                Else
                                    eResultadoConsulta.consecutivo = String.Empty
                                    eResultadoConsulta.errores = "El registro NO se ha almacenado. Intente nuevamente más tarde."
                                    eResultadoConsulta.resultado = False

                                    insertarLog("insertarPersonaVacuna", "El registro NO se ha almacenado. Intente nuevamente más tarde", Nothing)
                                    insertarLogBD("insertarPersonaVacuna", "El registro NO se ha almacenado. Intente nuevamente más tarde", Nothing)

                                    Return eResultadoConsulta
                                    Exit Function
                                End If
                            Else
                                eResultadoConsulta.consecutivo = String.Empty
                                eResultadoConsulta.errores = "No hay institución seleccionada, deberá volver a conectarse."
                                eResultadoConsulta.resultado = False

                                insertarLog("insertarPersonaVacuna", "No hay institución seleccionada, deberá volver a conectarse", Nothing)
                                insertarLogBD("insertarPersonaVacuna", "No hay institución seleccionada, deberá volver a conectarse", Nothing)

                                Return eResultadoConsulta
                                Exit Function
                            End If
                        Else
                            eResultadoConsulta.consecutivo = String.Empty
                            eResultadoConsulta.errores = "El Numero de Identificacion del Menor es igual al de la Madre."
                            eResultadoConsulta.resultado = False

                            insertarLog("insertarPersonaVacuna", "El Numero de Identificacion del Menor es igual al de la Madre.", Nothing)
                            insertarLogBD("insertarPersonaVacuna", "El Numero de Identificacion del Menor es igual al de la Madre.", Nothing)

                            Return eResultadoConsulta
                            Exit Function
                        End If

                    Else
                        eResultadoConsulta.consecutivo = String.Empty
                        eResultadoConsulta.errores = "La fecha de nacimiento es inconsistente, verifiquela."
                        eResultadoConsulta.resultado = False

                        insertarLog("insertarPersonaVacuna", "La fecha de nacimiento es inconsistente, verifiquela", Nothing)
                        insertarLogBD("insertarPersonaVacuna", "La fecha de nacimiento es inconsistente, verifiquela", Nothing)

                        Return eResultadoConsulta
                        Exit Function
                    End If
                End If
            End If
            Return eResultadoConsulta
        End Function

    Public Function InsertarPersona(ByVal per_TipoId As String,
                                        ByVal per_Id As String,
                                        ByVal per_CertNacVivo As Long,
                                        ByVal per_CertDefuncion As String,
                                        ByVal per_TipoIdM As String,
                                        ByVal per_IdM As String,
                                        ByVal per_NumeroHijoM As Integer,
                                        ByVal per_Nombre1M As String,
                                        ByVal per_Nombre2M As String,
                                        ByVal per_Apellido1M As String,
                                        ByVal per_Apellido2M As String,
                                        ByVal per_Nombre1 As String,
                                        ByVal per_Nombre2 As String,
                                        ByVal per_Apellido1 As String,
                                        ByVal per_Apellido2 As String,
                                        ByVal per_FechaNac As Date,
                                        ByVal per_Func As String,
                                        ByVal per_Institucion As String,
                                        ByVal per_Estado As Integer,
                                        ByVal per_cni_id As Integer,
                                        ByVal per_idEtnia As Integer,
                                        ByVal per_IdGrupoPoblacional As String,
                                        ByVal per_IdGenero As String,
                                        ByVal per_IdGrupoSanguineo As String,
                                        ByVal per_IdRh As String) As resultadoConsultaEntity Implements IpaiServicioSDS.InsertarPersona


        Dim oResultado As New resultadoConsultaEntity

        Dim fValidacion As Boolean = True
        Dim MensajeValidacion As String = String.Empty

        '---------------Aqui se validan los valores en 0 o vacios
        If String.IsNullOrEmpty(per_TipoId) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_tipoId no puede estar vacía"
        End If
        If String.IsNullOrEmpty(per_Id) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Id no puede estar vacía"
        End If
        If String.IsNullOrEmpty(per_Nombre1) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Nombre1 no puede estar vacía"
        End If
        If String.IsNullOrEmpty(per_Apellido1) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Apellido1 no puede estar vacía"
        End If
        If per_FechaNac = Date.MinValue Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_FechaNac no puede estar vacía"
        Else
            Dim edad As Decimal = calcularEdad(per_FechaNac)
            If edad < MAYOR_EDAD Then 'Si es menor de edad (18 años) se valida que tenga la información de la mama.
                If String.IsNullOrEmpty(per_TipoIdM) Then
                    oResultado.resultado = False
                    oResultado.errores += " La variable per_TipoIdM no puede estar vacía"
                End If
                If String.IsNullOrEmpty(per_IdM) Then
                    oResultado.resultado = False
                    oResultado.errores += " La variable per_IdM no puede estar vacía"
                End If
                If String.IsNullOrEmpty(per_Nombre1M) Then
                    oResultado.resultado = False
                    oResultado.errores += " La variable per_Nombre1M no puede estar vacía"
                End If
                If String.IsNullOrEmpty(per_Apellido1M) Then
                    oResultado.resultado = False
                    oResultado.errores += " La variable per_Apellido1M no puede estar vacía"
                End If
                If Not validarTablasDominio(14, per_TipoIdM, MensajeValidacion) Then
                    oResultado.resultado = False
                    oResultado.errores += MensajeValidacion
                End If
            End If
        End If
        If String.IsNullOrEmpty(per_Func) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Func no puede estar vacía"
        End If
        If String.IsNullOrEmpty(per_Institucion) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Institucion no puede estar vacía"
        End If
        If per_Estado <> ESTADO_ACTIVO OrElse per_Estado <> ESTADO_FALLECIDO Then
            oResultado.resultado = False
            oResultado.errores = "La variable per_Estado debe ser 2 (Registro activo) o 4 (Fallecido)"
            fValidacion = False
        End If
        If per_idEtnia = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_idEtnia no puede estar vacía"
        End If
        If String.IsNullOrEmpty(per_IdGenero) Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_IdGenero no puede estar vacía"
        End If

        '----------------- Aqui se valida contra las tablas de dominio
        If Not validarTablasDominio(14, per_TipoId, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores = MensajeValidacion
        End If
        If Not validarTablasDominio(32, per_Func.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(15, per_Institucion.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not validarTablasDominio(16, per_Estado.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If per_cni_id <> 0 Then
            If Not validarTablasDominio(17, per_cni_id.ToString, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
        End If
        If Not validarTablasDominio(18, per_idEtnia.ToString, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not String.IsNullOrEmpty(per_IdGrupoPoblacional) Then
            If Not validarTablasDominio(19, per_IdGrupoPoblacional, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
        End If
        If Not validarTablasDominio(20, per_IdGenero, MensajeValidacion) Then
            oResultado.resultado = False
            oResultado.errores += MensajeValidacion
        End If
        If Not String.IsNullOrEmpty(per_IdGrupoSanguineo) Then
            If Not validarTablasDominio(20, per_IdGrupoSanguineo, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
        End If
        If Not String.IsNullOrEmpty(per_IdRh) Then
            If Not validarTablasDominio(22, per_IdRh, MensajeValidacion) Then
                oResultado.resultado = False
                oResultado.errores += MensajeValidacion
            End If
        End If

        '-------------------

        If fValidacion Then
            Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
            Dim arParms() As SqlParameter = New SqlParameter(25) {}

            arParms(0) = New SqlParameter("@per_TipoId", 3)
            If String.IsNullOrEmpty(per_TipoId) Then
                arParms(0).Value = System.DBNull.Value
            Else
                arParms(0).Value = per_TipoId
            End If

            arParms(1) = New SqlParameter("@per_Id", 22)
            If String.IsNullOrEmpty(per_Id) Then
                arParms(1).Value = System.DBNull.Value
            Else
                arParms(1).Value = per_Id
            End If

            arParms(2) = New SqlParameter("@per_CertNacVivo", 0)
            If per_CertNacVivo = 0 Then
                arParms(2).Value = System.DBNull.Value
            Else
                arParms(2).Value = per_CertNacVivo
            End If

            arParms(3) = New SqlParameter("@per_CertDefuncion", 22)
            If String.IsNullOrEmpty(per_CertDefuncion) Then
                arParms(3).Value = System.DBNull.Value
            Else
                arParms(3).Value = per_CertDefuncion
            End If

            arParms(4) = New SqlParameter("@per_TipoIdM", 3)
            If String.IsNullOrEmpty(per_TipoIdM) Then
                arParms(4).Value = System.DBNull.Value
            Else
                arParms(4).Value = per_TipoIdM
            End If

            arParms(5) = New SqlParameter("@per_IdM", 22)
            If String.IsNullOrEmpty(per_IdM) Then
                arParms(5).Value = System.DBNull.Value
            Else
                arParms(5).Value = per_IdM
            End If

            arParms(6) = New SqlParameter("@per_NumeroHijoM", 16)
            If per_NumeroHijoM = 0 Then
                arParms(6).Value = System.DBNull.Value
            Else
                arParms(6).Value = per_NumeroHijoM
            End If

            arParms(7) = New SqlParameter("@per_Nombre1M", 22)
            If String.IsNullOrEmpty(per_Nombre1M) Then
                arParms(7).Value = System.DBNull.Value
            Else
                arParms(7).Value = per_Nombre1M
            End If

            arParms(8) = New SqlParameter("@per_Nombre2M", 22)
            If String.IsNullOrEmpty(per_Nombre2M) Then
                arParms(8).Value = System.DBNull.Value
            Else
                arParms(8).Value = per_Nombre2M
            End If

            arParms(9) = New SqlParameter("@per_Apellido1M", 22)
            If String.IsNullOrEmpty(per_Apellido1M) Then
                arParms(9).Value = System.DBNull.Value
            Else
                arParms(9).Value = per_Apellido1M
            End If

            arParms(10) = New SqlParameter("@per_Apellido2M", 22)
            If String.IsNullOrEmpty(per_Apellido2M) Then
                arParms(10).Value = System.DBNull.Value
            Else
                arParms(10).Value = per_Apellido2M
            End If

            arParms(11) = New SqlParameter("@per_Nombre1", 22)
            If String.IsNullOrEmpty(per_Nombre1) Then
                arParms(11).Value = System.DBNull.Value
            Else
                arParms(11).Value = per_Nombre1
            End If

            arParms(12) = New SqlParameter("@per_Nombre2", 22)
            If String.IsNullOrEmpty(per_Nombre2) Then
                arParms(12).Value = System.DBNull.Value
            Else
                arParms(12).Value = per_Nombre2
            End If

            arParms(13) = New SqlParameter("@per_Apellido1", 22)
            If String.IsNullOrEmpty(per_Apellido1) Then
                arParms(13).Value = System.DBNull.Value
            Else
                arParms(13).Value = per_Apellido1
            End If

            arParms(14) = New SqlParameter("@per_Apellido2", 22)
            If String.IsNullOrEmpty(per_Apellido2) Then
                arParms(14).Value = System.DBNull.Value
            Else
                arParms(14).Value = per_Apellido2
            End If

            arParms(15) = New SqlParameter("@per_FechaNac", 31)
            If per_FechaNac = Date.MinValue Then
                arParms(15).Value = System.DBNull.Value
            Else
                arParms(15).Value = per_FechaNac
            End If

            arParms(16) = New SqlParameter("@per_Func", 12)
            If String.IsNullOrEmpty(per_Func) Then
                arParms(16).Value = System.DBNull.Value
            Else
                arParms(16).Value = per_Func
            End If

            arParms(17) = New SqlParameter("@per_Institucion", 3)
            If String.IsNullOrEmpty(per_Institucion) Then
                arParms(17).Value = System.DBNull.Value
            Else
                arParms(17).Value = per_Institucion
            End If

            arParms(18) = New SqlParameter("@per_Estado", 8)
            If per_Estado = 0 Then
                arParms(18).Value = System.DBNull.Value
            Else
                arParms(18).Value = per_Estado
            End If

            arParms(19) = New SqlParameter("@per_cni_id", 8)
            If per_cni_id = 0 Then
                arParms(19).Value = System.DBNull.Value
            Else
                arParms(19).Value = per_cni_id
            End If

            arParms(20) = New SqlParameter("@per_idEtnia", 22)
            If per_idEtnia = 0 Then
                arParms(20).Value = System.DBNull.Value
            Else
                arParms(20).Value = per_idEtnia
            End If

            arParms(21) = New SqlParameter("@per_IdGrupoPoblacional", 3)
            If per_IdGrupoPoblacional = "0" Then
                arParms(21).Value = System.DBNull.Value
            Else
                arParms(21).Value = per_IdGrupoPoblacional
            End If

            'Alejandro Muñoz 26/10/2017 - Actualización cambio. Para que no haga homologación, ya que compensar envía ya los datos validados'
            arParms(22) = New SqlParameter("@per_IdGenero", 3)
            If String.IsNullOrEmpty(per_IdGenero) Then
                arParms(22).Value = System.DBNull.Value
            Else
                arParms(22).Value = per_IdGenero
            End If

            'arParms(22) = New SqlParameter("@per_IdGenero", 3)
            'If String.IsNullOrEmpty(per_IdGenero) Then

            '    arParms(22).Value = System.DBNull.Value
            '    'Alejandro Muñoz 24/10/2017' Solución error de envío del campo 'Género' x compensar
            'Else
            '    If per_IdGenero = "M" Then
            '        arParms(22).Value = "H"
            '    ElseIf per_IdGenero = "F" Then
            '        arParms(22).Value = "M"
            '    End If
            'End If

            arParms(23) = New SqlParameter("@per_IdGrupoSanguineo", 3)
            If String.IsNullOrEmpty(per_IdGrupoSanguineo) Then
                arParms(23).Value = System.DBNull.Value
            Else
                arParms(23).Value = per_IdGrupoSanguineo
            End If

            arParms(24) = New SqlParameter("@per_IdRh", 3)
            If String.IsNullOrEmpty(per_IdRh) Then
                arParms(24).Value = System.DBNull.Value
            Else
                arParms(24).Value = per_IdRh
            End If

            arParms(25) = New SqlParameter("@per_Consecutivo", 0)
            arParms(25).Value = System.DBNull.Value

            Try

                'If arParms(0).Value.ToString.Equals("RC") Then
                '    insertarLog("insertarPersona()", "VALIDACION (RC)", arParms)
                '    insertarLogBD("insertarPersona()", "VALIDACION (RC)", arParms)

                '    oResultado.resultado = False
                '    oResultado.errores = "No se puede ingresar una persona nueva con RC. Realice la búsqueda por los datos de la mama. (Posible duplicado)."
                'Else

                'End If

                Dim consec As Long = CLng(SqlHelper.ExecuteScalar(cadena, "pa_InsertarPersonaWS", arParms))
                If consec > 0 Then
                    oResultado.consecutivo = CStr(consec)
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("insertarPersona()", "CORRECTO", arParms)
                    insertarLogBD("insertarPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no identificado."

                    insertarLog("insertarPersona()", "El registro no fue ingresado", arParms)
                    insertarLogBD("insertarPersona()", "El registro no fue ingresado", arParms)
                End If
            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("insertarPersona()", e.Message, arParms)
                insertarLogBD("insertarPersona()", e.Message, arParms)
            End Try
        End If

        Return oResultado
    End Function

    Public Function seleccionarPersonaIdContar(ByVal per_Id As String, ByVal per_TipoId As String) As Integer
            Dim cadena As String = conexionBD
            Dim oPersona As New PersonaEntity()

            Dim arParms() As SqlParameter = New SqlParameter(1) {}
            arParms(0) = New SqlParameter("@per_TipoId", 0)
            arParms(0).Value = per_TipoId
            arParms(1) = New SqlParameter("@per_Id", 0)
            arParms(1).Value = per_Id
            Dim nroPersonas As Integer = CInt(SqlHelper.ExecuteScalar(cadena, "pa_SeleccionarPersonaIdContar", arParms))

            insertarLog("seleccionarPersonaIdContar()", "CORRECTO", arParms)
            insertarLogBD("seleccionarPersonaIdContar()", "CORRECTO", arParms)

            Return nroPersonas
        End Function

    Public Function InsertarPesoPersona(ByVal per_Consecutivo As Long,
                                            ByVal pes_Peso As Integer) As resultadoConsultaEntity

        Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        If per_Consecutivo <> 0 And pes_Peso <> 0 Then


            Dim arParms() As SqlParameter = New SqlParameter(1) {}

            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_Consecutivo

            arParms(1) = New SqlParameter("@pes_Peso", 8)
            arParms(1).Value = pes_Peso

            Try
                If SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarPesoWS", arParms) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("insertarPesoPersona()", "CORRECTO", arParms)
                    insertarLogBD("insertarPesoPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no identificado."

                    insertarLog("insertarPesoPersona()", "Registro no ingresado", arParms)
                    insertarLogBD("insertarPesoPersona()", "Registro no ingresado", arParms)
                End If
            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("insertarPesoPersona()", e.Message, arParms)
                insertarLogBD("insertarPesoPersona()", e.Message, arParms)
            End Try
            'Catch ex As Exception
            '    oResultado.resultado = False
            '    oResultado.errores = ex.Message.ToString()
            'End Try
        Else
            oResultado.resultado = False
            oResultado.errores = "El per_consecutivo fue envido como 0"
        End If
        Return oResultado
    End Function

    Public Function calculoEdad(edadInicial As Date, edadFinal As Date, Optional ByRef anios As Integer = 0, Optional ByRef meses As Integer = 0, Optional ByRef dias As Integer = 0) As String
            Dim mesNac, mesActual As Integer
            Dim diaNac, diaActual As Integer
            Dim anioNac, anioActual As Integer
            'Dim anios, meses, dias As Integer
            Dim diaMesAnterior As Integer

            diaNac = CDate(edadInicial).Day
            mesNac = CDate(edadInicial).Month
            anioNac = CDate(edadInicial).Year

            diaActual = CDate(edadFinal).Day
            mesActual = CDate(edadFinal).Month
            anioActual = CDate(edadFinal).Year

            anios = anioActual - anioNac
            meses = mesActual - mesNac
            dias = diaActual - diaNac

            'ajuste de negativo dias
            If dias < 0 Then
                meses -= 1
                Select Case mesActual
                    Case 1, 2, 4, 6, 8, 9, 11
                        diaMesAnterior = 31
                    Case 3
                        If bisiesto(anioActual) Then
                            diaMesAnterior = 29
                        Else
                            diaMesAnterior = 28
                        End If
                    Case 5, 7, 10, 12
                        diaMesAnterior = 30
                End Select
                dias = dias + diaMesAnterior
            End If
            'ajuste negativo mes
            If meses < 0 Then
                anios -= 1
                meses = meses + 12
            End If
            Return anios & " años " & meses & " meses " & dias & " días  "

        End Function

        Public Function bisiesto(anioActual As Integer) As Boolean
            Dim fbisiesto As Boolean = False
            If IsDate(Format("dd/MM/YYYY", "29/02/" + anioActual.ToString)) Then
                fbisiesto = True
            End If
            Return fbisiesto
        End Function

    Public Function InsertarVacunaPersona(ByVal per_Consecutivo As Long,
                                              ByVal vac_Id As Integer,
                                              ByVal dos_Id As Integer,
                                              ByVal pse_Id As Integer,
                                              ByVal cam_id As Integer,
                                              ByVal vac_FechaVacuna As Date,
                                              ByVal vac_actualizacion As Boolean,
                                              ByVal ins_Id As String,
                                              ByVal fun_idFunc As String,
                                              ByVal com_Id As Integer,
                                              ByVal pos_Id As Integer,
                                              ByVal vac_Lote As String,
                                              ByVal vac_EdadVacunaAnios As Integer,
                                              ByVal vac_EdadVacunaMeses As Integer,
                                              ByVal vac_EdadVacunaDias As Integer,
                                              ByVal vac_EdadVacunaTotalDias As Integer,
                                              ByVal cdm_idCondicion As Integer) As resultadoConsultaEntity Implements IpaiServicioSDS.insertarVacunaPersona


        Dim oResultado As New resultadoConsultaEntity
        Dim fValidacion As Boolean = True
        Dim mensajeValidacion As String = String.Empty 'Se agrega el 29/09/2017 Alejandro Muñoz para solucionar POS_ID = 0'


        If per_Consecutivo = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable per_Consecutivo no puede estar vacía"
            fValidacion = False
        End If

        If vac_Id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_Id no puede estar vacía"
            fValidacion = False
        End If
        If dos_Id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable dos_Id no puede estar vacía"
            fValidacion = False
        End If
        If pse_Id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable pse_Id no puede estar vacía"
            fValidacion = False
        End If
        If cam_id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable cam_id no puede estar vacía"
            fValidacion = False
        End If

        If vac_FechaVacuna = Date.MinValue Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_FechaVacuna no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(ins_Id) Then
            oResultado.resultado = False
            oResultado.errores += " La variable ins_Id no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(fun_idFunc) Then
            oResultado.resultado = False
            oResultado.errores += " La variable fun_idFunc no puede estar vacía"
            fValidacion = False
        End If
        If com_Id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable com_Id no puede estar vacía"
            fValidacion = False
        End If
        If pos_Id = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable pos_Id no puede estar vacía"
            fValidacion = False
        End If
        If String.IsNullOrEmpty(vac_Lote) Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_Lote no puede estar vacía"
            fValidacion = False
        End If
        If vac_EdadVacunaAnios = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_EdadVacunaAnios no puede estar vacía"
            fValidacion = False
        End If
        If vac_EdadVacunaMeses = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_EdadVacunaMeses no puede estar vacía"
            fValidacion = False
        End If

        If vac_EdadVacunaDias = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_EdadVacunaDias no puede estar vacía"
            fValidacion = False
        End If

        If vac_EdadVacunaTotalDias = 0 Then
            oResultado.resultado = False
            oResultado.errores += " La variable vac_EdadVacunaTotalDias no puede estar vacía"
            fValidacion = False
        End If

        'TODO: Faltan validaciones contra tablas de dominio

        If fValidacion Then
            Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
            Dim arParms() As SqlParameter = New SqlParameter(16) {}

            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_Consecutivo

            arParms(1) = New SqlParameter("@vac_Id", 8)
            arParms(1).Value = vac_Id

            arParms(2) = New SqlParameter("@dos_Id", 8)
            arParms(2).Value = dos_Id

            arParms(3) = New SqlParameter("@pse_Id", 8)
            arParms(3).Value = pse_Id

            arParms(4) = New SqlParameter("@cam_id", 8)
            If cam_id = 0 Then
                arParms(4).Value = System.DBNull.Value
            Else
                arParms(4).Value = cam_id
            End If

            arParms(5) = New SqlParameter("@vac_FechaVacuna", 31)
            arParms(5).Value = vac_FechaVacuna

            arParms(6) = New SqlParameter("@vac_actualizacion", 2)
            arParms(6).Value = vac_actualizacion

            arParms(7) = New SqlParameter("@ins_Id", 3)
            arParms(7).Value = ins_Id

            arParms(8) = New SqlParameter("@fun_idFunc", 12)
            arParms(8).Value = fun_idFunc

            arParms(9) = New SqlParameter("@com_Id", 8)
            arParms(9).Value = com_Id

            arParms(10) = New SqlParameter("@pos_Id", 20)
            arParms(10).Value = pos_Id

            arParms(11) = New SqlParameter("@vac_Lote", 22)
            arParms(11).Value = vac_Lote

            arParms(12) = New SqlParameter("@vac_EdadVacunaAnios", 8)
            arParms(12).Value = vac_EdadVacunaAnios

            arParms(13) = New SqlParameter("@vac_EdadVacunaMeses", 8)
            arParms(13).Value = vac_EdadVacunaMeses

            arParms(14) = New SqlParameter("@vac_EdadVacunaDias", 8)
            arParms(14).Value = vac_EdadVacunaDias

            arParms(15) = New SqlParameter("@vac_EdadVacunaTotalDias", 8)
            arParms(15).Value = vac_EdadVacunaTotalDias

            arParms(16) = New SqlParameter("@cdm_idCondicion", 8)
            If cdm_idCondicion <> 0 Then
                arParms(16).Value = cdm_idCondicion
            Else
                arParms(16).Value = System.DBNull.Value
            End If

            Try
                If CLng(SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarVacunaPersonaWS", arParms)) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("insertarVacunaPersona()", "CORRECTO", arParms)
                    insertarLogBD("insertarVacunaPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no ingresado."

                    insertarLog("insertarVacunaPersona()", "Registro no ingresado", arParms)
                    insertarLogBD("insertarVacunaPersona()", "Registro no ingresado", arParms)
                End If
            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("insertarVacunaPersona()", e.Message, arParms)
                insertarLogBD("insertarVacunaPersona()", e.Message, arParms)
            End Try
        End If

        Return oResultado
    End Function

    Public Function ActualizarEsquemaSeguimientoPersona(per_Consecutivo As Long) As resultadoConsultaEntity
        Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        If per_Consecutivo <> 0 Then


            Dim arParms() As SqlParameter = New SqlParameter(0) {}

            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_Consecutivo

            Try
                If CLng(SqlHelper.ExecuteNonQuery(cadena, "pa_ActualizarEsquemaSeguimientoCohorteWS", arParms)) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("ActualizarEsquemaSegumientoPersona()", "CORRECTO", arParms)
                    insertarLogBD("ActualizarEsquemaSegumientoPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no ingresado."

                    insertarLog("ActualizarEsquemaSegumientoPersona()", "Registro no ingresado", arParms)
                    insertarLogBD("ActualizarEsquemaSegumientoPersona()", "Registro no ingresado", arParms)
                End If
            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("ActualizarEsquemaSegumientoPersona()", e.Message, arParms)
                insertarLogBD("ActualizarEsquemaSegumientoPersona()", e.Message, arParms)
            End Try
            'Catch ex As Exception
            '    oResultado.resultado = False
            '    oResultado.errores = ex.Message.ToString()
            'End Try
        Else
            oResultado.resultado = False
            oResultado.errores = "El per_consecutivo fue enviado como 0"
        End If
        Return oResultado
    End Function

    Public Function ActualizarInstitucionParto(per_Consecutivo As Long, per_ParInstitucion As String) As resultadoConsultaEntity
        Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        If per_Consecutivo <> 0 And per_ParInstitucion <> "" Then


            Dim arParms() As SqlParameter = New SqlParameter(1) {}

            arParms(0) = New SqlParameter("@per_Consecutivo", 0)
            arParms(0).Value = per_Consecutivo

            arParms(1) = New SqlParameter("@per_ParInstitucion", 3)
            arParms(1).Value = per_ParInstitucion

            Try
                If CLng(SqlHelper.ExecuteNonQuery(cadena, "pa_ActualizarInstitucionPartoWS", arParms)) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("ActualizarInstitucionParto()", "CORRECTO", arParms)
                    insertarLogBD("ActualizarInstitucionParto()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Registro no ingresado."

                    insertarLog("ActualizarInstitucionParto()", "Registro no ingresado", arParms)
                    insertarLogBD("ActualizarInstitucionParto()", "Registro no ingresado", arParms)
                End If
            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("ActualizarInstitucionParto()", e.Message, arParms)
                insertarLogBD("ActualizarInstitucionParto()", e.Message, arParms)
            End Try
            'Catch ex As Exception
            '    oResultado.resultado = False
            '    oResultado.errores = ex.Message.ToString()
            'End Try

        Else
            oResultado.resultado = False
            oResultado.errores = "El per_consecutivo fue enviado como 0"
        End If
        Return oResultado
    End Function

    Public Function InsertarCausasNoVacunaPersona(ByVal cnv_Id As Integer,
                                                      ByVal vac_Id As Integer,
                                                      ByVal per_Consecutivo As Long,
                                                      ByVal dos_id As Long) As resultadoConsultaEntity

        Dim cadena As String = conexionBD 'My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity
        If per_Consecutivo <> 0 Then

            Dim arParms() As SqlParameter = New SqlParameter(3) {}

            arParms(0) = New SqlParameter("@cnv_Id", 8)
            arParms(0).Value = cnv_Id

            arParms(1) = New SqlParameter("@vac_Id", 8)
            arParms(1).Value = vac_Id

            arParms(2) = New SqlParameter("@per_Consecutivo", 0)
            arParms(2).Value = per_Consecutivo

            arParms(3) = New SqlParameter("@dos_id", 8)
            arParms(3).Value = dos_id

            Try
                If SqlHelper.ExecuteNonQuery(cadena, "pa_InsertarCausasNoVacunacionWS", arParms) > 0 Then
                    oResultado.resultado = True
                    oResultado.errores = String.Empty

                    insertarLog("InsertarCausasNoVacunaPersona()", "CORRECTO", arParms)
                Else
                    oResultado.resultado = False
                    oResultado.errores = "Ingreso no realizado."

                    insertarLog("InsertarCausasNoVacunaPersona()", "Registro no ingresado", arParms)
                End If

            Catch e As SqlException
                Dim errorMessage As String = e.Message
                Dim errorCode As Integer = e.ErrorCode
                oResultado.resultado = False
                oResultado.errores = errorMessage

                insertarLog("InsertarCausasNoVacunaPersona()", e.Message, arParms)
            End Try
            'Catch ex As Exception
            '    oResultado.resultado = False
            '    oResultado.errores = ex.Message.ToString()
            'End Try

        Else
            oResultado.resultado = False
            oResultado.errores = "El per_consecutivo fue enviado como 0"
        End If
        Return oResultado
    End Function

    ''' <summary>
    ''' Retorna Json { "per_Consecutivo": numero, ["per_TipoId":"Tipo", "per_Id","id"] }
    ''' </summary>
    ''' <param name="per_Consecutivo">Numero consecutivo de la persona</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function seleccionarPersonaIdentificacion(per_Consecutivo As Long) As String Implements IpaiServicioSDS.seleccionarPersonaIdentificacion
            '[pa_SeleccionarPersonaIdentificacion]
            Dim strReturn As String = ""
            Dim cadena As String = conexionBD
            Dim lParms As List(Of SqlParameter) = New List(Of SqlParameter)
            Dim comi = Chr(34)
            lParms.Add(New SqlParameter("@per_Consecutivo", 0) With {.Value = per_Consecutivo})

            Dim dsPersonaIdentificacion As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPersonaIdentificacion", lParms.ToArray())

            If dsPersonaIdentificacion IsNot Nothing AndAlso dsPersonaIdentificacion.Tables(0).Rows.Count > 0 Then
                For Each oRow As Data.DataRow In dsPersonaIdentificacion.Tables(0).Rows
                    Dim strItem As String = ""
                    'strItem += "per_Consecutivo" + ":" + oRow("per_Consecutivo").ToString() + ","
                    strItem += String.Concat(comi, "per_TipoId", comi, ":", comi, oRow("per_TipoId").ToString(), comi, ",")
                    strItem += String.Concat(comi, "per_Id", comi, ":", comi, oRow("per_Id").ToString(), comi, "")
                    'oRow("doc_Descripcion")
                    'oRow("userName")
                    'oRow("per_Institucion")
                    'oRow("fecha")
                    strReturn += ",{" + strItem + "}"
                Next
                If strReturn <> "" Then
                    strReturn = String.Concat(comi, "historico", comi, ":[", Replace(strReturn, ",", "", 1, 1, CompareMethod.Binary), "]")
                    strReturn = String.Concat("{", comi, "per_Consecutivo", comi, ":", per_Consecutivo, ",", strReturn, "}")
                End If
            End If


            Dim arParms() As SqlParameter = New SqlParameter(0) {}
            arParms(0) = New SqlParameter("@cnv_Id", 3)
            arParms(0).Value = strReturn

            insertarLog("seleccionarPersonaIdentificacion()", "CORRECTO", arParms)
            insertarLogBD("seleccionarPersonaIdentificacion()", "CORRECTO", arParms)

            Return strReturn
        End Function

    Public Function SeleccionarPersonaBusquedaId(per_Consecutivo As Long) As PersonaEntity Implements IpaiServicioSDS.SeleccionarPersonaBusquedaId
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim persona As New PersonaEntity

        Dim arParms() As SqlParameter = New SqlParameter(0) {}

        arParms(0) = New SqlParameter("@per_Consecutivo", 0)
        arParms(0).Value = per_Consecutivo

        Dim dsPersona As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPersonaBusquedaId", arParms)
        If dsPersona IsNot Nothing AndAlso dsPersona.Tables(0).Rows.Count > 0 Then
            Dim dr As Data.DataRow = dsPersona.Tables(0).Rows(0)
            persona.per_Consecutivo = per_Consecutivo
            persona.per_TipoId = If(IsDBNull(dr("per_TipoId")), persona.per_TipoId, CStr(dr("per_TipoId")))
            persona.per_Id = If(IsDBNull(dr("per_Id")), persona.per_Id, CStr(dr("per_Id")))
            persona.per_CertNacVivo = If(IsDBNull(dr("per_CertNacVivo")), persona.per_CertNacVivo, CStr(dr("per_CertNacVivo")))
            persona.per_CertDefuncion = If(IsDBNull(dr("per_CertDefuncion")), persona.per_CertDefuncion, CStr(dr("per_CertDefuncion")))
            persona.per_TipoIdM = If(IsDBNull(dr("per_TipoIdM")), persona.per_TipoIdM, CStr(dr("per_TipoIdM")))
            persona.per_IdM = If(IsDBNull(dr("per_IdM")), persona.per_IdM, CStr(dr("per_IdM")))
            persona.per_NumeroHijoM = If(IsDBNull(dr("per_NumeroHijoM")), persona.per_NumeroHijoM, CShort(dr("per_NumeroHijoM")))
            persona.primerNombre = If(IsDBNull(dr("per_Nombre1")), persona.primerNombre, CStr(dr("per_Nombre1")))
            persona.segundoNombre = If(IsDBNull(dr("per_Nombre2")), persona.segundoNombre, CStr(dr("per_Nombre2")))
            persona.primerApellido = If(IsDBNull(dr("per_Apellido1")), persona.primerApellido, CStr(dr("per_Apellido1")))
            persona.segundoApellido = If(IsDBNull(dr("per_Apellido2")), persona.segundoApellido, CStr(dr("per_Apellido2")))
            persona.perFechaNac = If(IsDBNull(dr("per_FechaNac")), persona.perFechaNac, CDate(dr("per_FechaNac")))
            persona.per_parInstitucion = If(IsDBNull(dr("per_parInstitucion")), persona.per_parInstitucion, CStr(dr("per_parInstitucion")))
            persona.per_FechaAlm = If(IsDBNull(dr("per_FechaAlm")), persona.per_FechaAlm, CDate(dr("per_FechaAlm")))
            persona.per_Func = If(IsDBNull(dr("per_Func")), persona.per_Func, CStr(dr("per_Func")))
            persona.per_Institucion = If(IsDBNull(dr("per_Institucion")), persona.per_Institucion, CStr(dr("per_Institucion")))
            persona.per_Estado = If(IsDBNull(dr("per_Estado")), persona.per_Estado, CInt(dr("per_Estado")))
            persona.cni_id = If(IsDBNull(dr("per_cni_id")), persona.cni_id, CInt(dr("per_cni_id")))
            persona.etn_idEtnia = If(IsDBNull(dr("per_idEtnia")), persona.etn_idEtnia, CShort(dr("per_idEtnia")))
            persona.gru_IdGrupo = If(IsDBNull(dr("per_IdGrupoPoblacional")), persona.gru_IdGrupo, CStr(dr("per_IdGrupoPoblacional")))
            'Alejandro Muñoz 26/10/2017 - Actualización cambio. Para que no haga homologación, ya que compensar envía ya los datos validados'
            persona.per_Genero = If(IsDBNull(dr("per_IdGenero")), persona.per_Genero, CStr(dr("per_IdGenero")))

            'If Not dr("per_IdGenero").GetType.ToString.Equals("System.DBNull") Then
            '    If CStr(dr("per_IdGenero")) = "M" Then
            '        persona.per_Genero = "H"
            '    ElseIf CStr(dr("per_IdGenero")) = "F" Then
            '        persona.per_Genero = "M"
            '    End If
            'End If
            'persona.per_Genero = persona.per_Genero
            persona.perGrupoSanguineo = If(IsDBNull(dr("per_IdGrupoSanguineo")), persona.perGrupoSanguineo, CStr(dr("per_IdGrupoSanguineo")))
            persona.perRh = If(IsDBNull(dr("per_IdRh")), persona.perRh, CStr(dr("per_IdRh")))
            persona.pes_Peso = If(IsDBNull(dr("Pes_Peso")), persona.pes_Peso, CStr(dr("Pes_Peso")))
        End If
        dsPersona.Dispose()

        insertarLog("SeleccionarPersonaBusquedaId()", "CORRECTO", arParms)
        insertarLogBD("SeleccionarPersonaBusquedaId()", "CORRECTO", arParms)

        Return persona
    End Function



    Public Function seleccionarPresentacionComercial(ByVal vac_id As Integer, ByVal Pertenencia_POS As Integer, ByVal Grupo_etareo As Integer) As VacunaCollection Implements IpaiServicioSDS.seleccionarPresentacionComercial
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity

        Dim arParms() As SqlParameter = New SqlParameter(2) {}

        arParms(0) = New SqlParameter("@vac_Id", 8)
        arParms(0).Value = vac_id

        arParms(1) = New SqlParameter("@com_POS_Id", 8)
        arParms(1).Value = Pertenencia_POS


        arParms(2) = New SqlParameter("@com_Grup_Id", 8)
        arParms(2).Value = Grupo_etareo

        Dim dsPresentacioncomercial As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPresentacionComercialPorVacuna", arParms)
        Dim PresentacionComerc As New VacunaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsPresentacioncomercial.Tables(0).Rows
            Dim Comercial As New VacunaEntity()
            Comercial.com_Id = CInt(dr("com_Id"))
            Comercial.com_nombre = CStr(dr("com_Descripcion"))

            PresentacionComerc.Add(Comercial)
        Next

        insertarLog("seleccionarPresentacionComercial()", "CORRECTO", arParms)

        dsPresentacioncomercial.Dispose()

        Return PresentacionComerc
    End Function


    Public Function seleccionarEsquemavacunasPAI(ByVal Grup_Etareo As Integer) As VacunaCollection Implements IpaiServicioSDS.seleccionarEsquemavacunasPAI
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity

        Dim arParms() As SqlParameter = New SqlParameter(0) {}

        arParms(0) = New SqlParameter("@Grup_etareo", 8)
        arParms(0).Value = Grup_Etareo

        Dim dsEsquema As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "Pa_SeleccionarEsquemaVacunacionPAi", arParms)
        Dim EsquemaVacunacion As New VacunaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsEsquema.Tables(0).Rows
            Dim Esquema As New VacunaEntity()
            Esquema.vac_id = CInt(dr("vac_id"))
            Esquema.vac_Nombre = CStr(dr("vac_nombre"))
            Esquema.Dos_Nombre = CStr(dr("dos_nombre"))
            Esquema.Dos_id = CInt(dr("dos_Id"))
            Esquema.grup_nombre = CStr(dr("grup_Nombre"))
            Esquema.grup_Id = CInt(dr("grup_id"))
            Esquema.pos_nombre = CStr(dr("pos_Descripcion"))
            Esquema.pos_Id = CInt(dr("vac_pos"))



            EsquemaVacunacion.Add(Esquema)
        Next

        insertarLog("seleccionarEsquemaPAI()", "CORRECTO", arParms)
        dsEsquema.Dispose()

        Return EsquemaVacunacion
    End Function


    Public Function seleccionarPresentacion(ByVal vac_id As Integer, ByVal Pertenencia_POS As Integer, ByVal Grupo_etareo As Integer) As VacunaCollection Implements IpaiServicioSDS.seleccionarPresentacion
        Dim cadena As String = conexionBD ' My.Settings.cadenaPai20
        Dim oResultado As New resultadoConsultaEntity

        Dim arParms() As SqlParameter = New SqlParameter(2) {}

        arParms(0) = New SqlParameter("@vac_Id", 8)
        arParms(0).Value = vac_id

        arParms(1) = New SqlParameter("@pse_POS", 8)
        arParms(1).Value = Pertenencia_POS


        arParms(2) = New SqlParameter("@pse_Grup_Id", 8)
        arParms(2).Value = Grupo_etareo

        Dim dsPresentacion As Data.DataSet = SqlHelper.ExecuteDataset(cadena, "pa_SeleccionarPresentacionPorVacuna", arParms)
        Dim Presentacion As New VacunaCollection()
        Dim dr As Data.DataRow
        For Each dr In dsPresentacion.Tables(0).Rows
            Dim present As New VacunaEntity()
            present.pse_Id = CInt(dr("pse_Id"))
            present.Pse_Nombre = CStr(dr("pse_Descripcion"))



            Presentacion.Add(present)
        Next

        insertarLog("seleccionarPresentacion()", "CORRECTO", arParms)
        dsPresentacion.Dispose()

        Return Presentacion
    End Function

End Class
