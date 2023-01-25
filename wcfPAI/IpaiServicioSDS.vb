Imports System.ServiceModel
Imports Sds.PAI.wsPaiEntity



<ServiceContract(Namespace:="http://Sds.SaludPublica")>
    Public Interface IpaiServicioSDS

        <OperationContract()>
        Function seleccionarPersonaBusqueda(ByVal TipoIdVacunado As String,
                                                  ByVal NumeroIdVacunado As String,
                                                  ByVal PrimerNombreVacunado As String,
                                                  ByVal SegundoNombreVacunado As String,
                                                  ByVal PrimerApellidoVacunado As String,
                                                  ByVal SegundoApellidoVacunado As String,
                                                  ByVal per_parInstitucion As String,
                                                  ByVal per_FechaNac As Date,
                                                  ByVal TipoIdentificacionMadre As String,
                                                  ByVal NumeroIdentificacionMadre As String,
                                                  ByVal PrimerNombreMadre As String,
                                                  ByVal SegundoNombreMadre As String,
                                                  ByVal PrimerApellidoMadre As String,
                                                  ByVal SegundoApellidoMadre As String,
                                                  ByVal grupoEtareo As Integer) As PersonaCollection

        <OperationContract()>
        Function seleccionarVacunasPersona(ByVal per_Consecutivo As Long) As VacunaCollection

        <OperationContract()>
        Function seleccionarUbicacionPersona(ByVal per_Consecutivo As Long) As UbicacionPersonaEntity

        <OperationContract()>
        Function seleccionarAfiliacionPersona(ByVal per_Consecutivo As Long) As PersonaAfiliacionEntity

        <OperationContract()>
        Function seleccionarTablaDominio(ByVal id_Tabla As Short) As TablaDominioCollection

        <OperationContract()>
        Function seleccionarEstadoContactenos(ByVal IdCaso As Long) As String

    <OperationContract()>
    Function seleccionarTablasConNovedades() As TablaDominioCollection

    <OperationContract()>
    Function seleccionarSeguimientoCohorte(mesNacimiento As Integer, anioNacimiento As Integer) As PersonaCohorteCollection

    <OperationContract()>
        Function seleccionarEsquemavacunasPAIPendiente(ByVal per_consecutivo As Long) As VacunaCollection

        <OperationContract()>
        Function actualizarPersona(ByVal per_consecutivo As Long,
                                 ByVal per_TipoId As String,
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
                                ByVal per_IdRh As String) As resultadoConsultaEntity


        <OperationContract()>
        Function insertarVacunaPersona(ByVal per_Consecutivo As Long,
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
                              ByVal cdm_idCondicion As Integer) As resultadoConsultaEntity

    <OperationContract()>
    Function insertarUbicacionPersona(ByVal per_Consecutivo As Long,
                                         ByVal dir_Direccion As String,
                                         ByVal dir_Barrio As String,
                                         ByVal dir_Codigo_direccion As String,
                                         ByVal dir_Upz As String,
                                         ByVal dir_CoordenadaX As String,
                                         ByVal dir_CoordenadaY As String,
                                         ByVal dir_Localidad As Integer,
                                         ByVal dir_Estrato As String,
                                         ByVal tel_Telefono As String,
                                         ByVal tel_Contacto As String,
                                         ByVal cor_correo As String,
                                         ByVal dir_mun_id As Integer,
                                         ByVal dir_dep_Id As Integer,
                                         ByVal dir_pais_Id As String,
                                         ByVal dir_zon_Id As Integer) As resultadoConsultaEntity

    <OperationContract()>
    Function insertarAfiliacionPersona(ByVal per_Consecutivo As Long,
                                          ByVal ase_id As String,
                                          ByVal reg_Id As Integer,
                                          ByVal tia_id As Integer) As resultadoConsultaEntity

    <OperationContract()>
    Function insertarContactenos(ByVal con_cat_id As Integer,
                                    ByVal con_mensaje As String,
                                    ByVal con_fun_idFunc As String) As resultadoConsultaEntity

    <OperationContract()>
    Function insertarSeguimientoCohorte(ByVal tsc_id As Integer,
                                           ByVal seg_Numero_Llamada As Integer,
                                           ByVal seg_Fecha_Llamada As Date,
                                           ByVal seg_Comunica As Boolean,
                                           ByVal seg_mnc As Integer,
                                           ByVal seg_Mensaje As Boolean,
                                           ByVal res_Id As Integer,
                                           ByVal mnv_Id As Integer,
                                           ByVal seg_Observaciones As String,
                                           ByVal seg_per_consecutivo As Long,
                                           ByVal seg_Activo As Boolean,
                                           ByVal seg_FechaSeguimientoPersonal As Date,
                                           ByVal seg_UserName As String) As resultadoConsultaEntity

    <OperationContract()>
    Function insertarPersonaVacuna(per_Id As String,
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
                                    per_Institucion As String) As resultadoConsultaEntity

    '--i 25/09/2013
    <OperationContract()>
    Function InsertarPersona(ByVal per_TipoId As String,
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
                            ByVal per_IdRh As String) As resultadoConsultaEntity

    <OperationContract()>
    Function seleccionarPersonaBusquedaAttr(TipoIdVacunado As String,
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
                                                bGetHdocum As Boolean) As PersonaCollection

    <OperationContract()>
        Function seleccionarPersonaIdentificacion(per_Consecutivo As Long) As String

    <OperationContract()>
    Function SeleccionarPersonaBusquedaId(per_Consecutivo As Long) As PersonaEntity

    <OperationContract()>
        Function seleccionarEsquemavacunasPAI(ByVal Grup_Etareo As Integer) As VacunaCollection

        <OperationContract()>
        Function seleccionarPresentacionComercial(ByVal vac_id As Integer, ByVal Pertenencia_POS As Integer, ByVal Grupo_etareo As Integer) As VacunaCollection

        <OperationContract()>
        Function seleccionarPresentacion(ByVal vac_id As Integer, ByVal Pertenencia_POS As Integer, ByVal Grupo_etareo As Integer) As VacunaCollection

    End Interface
