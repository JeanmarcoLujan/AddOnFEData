using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnFEData.Comunes
{
    class EstructuraDatos
    {
        #region _Attributes_

        int m_iErrCode = 0;
        string m_sErrMsg = "";
        private string m_sNombreAddon = Properties.Resources.NombreAddon;
        private string m_sDescripcionUDTAddon = Properties.Resources.DescripcionTablaAddon;
        private string m_sVersion = Properties.Resources.VersionAddon;

        #endregion

        #region _Constructor_

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        public EstructuraDatos()
        {
            try
            {
                if (ValidaVersion(m_sNombreAddon, m_sDescripcionUDTAddon, m_sVersion))
                {
                    RegistrarVersion(m_sNombreAddon, m_sVersion);
                    //CrearUDTAddOn();
                    CrearUDFAddOn();
                    //CrearUDOAddOn();
                    //PrecargarDatosAddOn();
                    //CrearAutorizacionesAddOn();
                    // GenerateDataStructureReports();
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion

        #region _Functions_

        /// <summary>
        /// Método para validar la versión instalada del AddOn dentro de la sociedad donde se está iniciando.
        /// </summary>
        /// <param name="NombreAddon">Nombre del AddOn que sale de los recursos del compilado.</param>
        /// <param name="Version">Versión del AddOn que sale de los recuros del compilado.</param>
        /// <returns>Returna TRUE o FALSE según el resultado de la operación de verificación en la sociedad.</returns>
        private bool ValidaVersion(string NombreAddon, string DescripUDTAddOn, string Version)
        {
            bool bRetorno = false;
            SAPbobsCOM.UserTable oUT = null;
            SAPbobsCOM.Recordset oRS = null;
            string NombreTabla = "";

            try
            {
                NombreTabla = NombreAddon.ToUpper();
                try
                {
                    oUT = Conexion.Conexion_SBO.m_oCompany.UserTables.Item(NombreTabla);
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("invalid field name")) oUT = null;
                    else throw ex;
                }

                if (oUT == null)
                {
                    CreaTablaMD(NombreTabla, DescripUDTAddOn, SAPbobsCOM.BoUTBTableType.bott_NoObject);
                    bRetorno = true;
                }
                else
                {
                    oRS = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRS.DoQuery(Consultas.ConsultaTablaConfiguracion(Conexion.Conexion_SBO.m_oCompany.DbServerType, NombreAddon, "", true));
                    if (oRS.RecordCount == 0)
                    {
                        bRetorno = true;
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Actualizará la esturctura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    else
                    {
                        if (int.Parse(Version.Replace(".", "").ToString()) > int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                        {
                            bRetorno = true;
                            Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Actualizará la esturctura de datos",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }

                        if (int.Parse(Version.Replace(".", "").ToString()) < int.Parse(oRS.Fields.Item("code").Value.ToString().Replace(".", "")))
                            Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Detectó una version superior para este Addon",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                FuncionesComunes.LiberarObjetoGenerico(oRS);
                oRS = null;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
            return bRetorno;
        }

        #endregion

        #region _Methods_

        /// <summary>
        /// Método para registar el AddOn y la versión del mismo dentro de la sociedad donde se está iniciando el AddOn.
        /// </summary>        
        /// <param name="NombreAddon">Nombre del AddOn que sale de los recursos del compilado.</param>
        /// <param name="Version">Versión del AddOn que sale de los recuros del compilado.</param>     
        private void RegistrarVersion(string NombreAddon, string Version)
        {
            SAPbobsCOM.UserTable oUT;
            string NombreTabla = "";
            try
            {
                NombreTabla = NombreAddon.ToUpper();
                oUT = Conexion.Conexion_SBO.m_oCompany.UserTables.Item(NombreTabla);
                oUT.Code = Version;
                oUT.Name = NombreAddon + " V-" + Version;
                m_iErrCode = oUT.Add();

                if (m_iErrCode == 0)
                    Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Se ingreso un nuevo registro a la BD ",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                else
                    Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error ingresar el registro en la tabla: "
                        + NombreTabla + ". Error: " + Conexion.Conexion_SBO.m_oCompany.GetLastErrorDescription().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de tablas de usuario (UDT) en la sociedad.
        /// </summary>
        private void CrearUDTAddOn()
        {
            try
            {
                //CreaTablaMD("MSS_BANCOVALIDO", "Formato Banco TXT", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //CreaTablaMD("MSS_PREGUARDADO_CAB", "Pre-Guardado CAB", SAPbobsCOM.BoUTBTableType.bott_Document);
                //CreaTablaMD("MSS_PREGUARDADO_LIN", "Pre-Guardado LIN", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de campos de usuarios (UDF) en la sociedad.
        /// </summary>
        private void CrearUDFAddOn()
        {
            try
            {


                string[] ValidValues = null;
                string[] ValidDescrip = null;

                ValidValues = new string[6] { "DP", "DE", "DS", "DR", "DA", "DI" };
                ValidDescrip = new string[6] { "Documento pendiente", "Documento con error", "Documento en seguimiento", "Documento rechazado", "Documento aprobado", "Documento interno" };

                CreaCampoMD("OINV", "MGS_FE_Estado", "MGS - FE Estado", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 4, SAPbobsCOM.BoYesNoEnum.tNO, ValidValues, ValidDescrip, null, "");
                CreaCampoMD("OINV", "MGS_FE_RespEnvio", "MGS - FE Respuesta Envío", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_EstatusDGII", "MGS - FE Estado DGII", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_MensajeDGII", "MGS - FE Mensaje DGII", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 250, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_PDF", "MGS - FE - PDF", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Link, 150, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_PDF", "MGS - FE - PDF2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Link, 150, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_QR", "MGS - FE - QR", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Link, 150, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_FechaFirma", "MGS - FE Fecha Firma", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_CodigoSeguridad", "MGS - FE Cod. Seguridad", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 30, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_Ref", "MGS - FE - Numero ref", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_RefNCF", "MGS - FE - NCF ref", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);
                CreaCampoMD("OINV", "MGS_FE_RefFecha", "MGS - FE - Fecha ref", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null);


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de objetos de usuarios (UDO) en la sociedad.
        /// </summary>
        private void CrearUDOAddOn()
        {

            string[] FindColumns = null;
            string[] ChildTables = null;

            try
            {
                FindColumns = new string[2] { "DocEntry", "DocNum" };
                ChildTables = new string[1] { "MSS_PREGUARDADO_LIN" };

                CreaUDOMD(
                    "Preguardado", //code
                    "Pre-Guardado", //name
                    "MSS_PREGUARDADO_CAB",//tablename 
                    FindColumns, //findcolumns
                    ChildTables, //childTable
                    SAPbobsCOM.BoYesNoEnum.tNO, //cancel 
                    SAPbobsCOM.BoYesNoEnum.tNO, //close
                    SAPbobsCOM.BoYesNoEnum.tNO, //delete
                    SAPbobsCOM.BoYesNoEnum.tNO, //createDefaul
                    null, //formColumns
                    SAPbobsCOM.BoYesNoEnum.tYES, //CanFind 
                    SAPbobsCOM.BoYesNoEnum.tNO, //Canlog
                    SAPbobsCOM.BoUDOObjType.boud_Document, //objettype
                    SAPbobsCOM.BoYesNoEnum.tYES, //manageseries, 
                    SAPbobsCOM.BoYesNoEnum.tNO,
                    SAPbobsCOM.BoYesNoEnum.tNO,
                    null,
                    null
                    );


            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para iniciar la creación de objetos de usuarios (UDO) en la sociedad.
        /// </summary>
        private void PrecargarDatosAddOn()
        {
            try
            {
                // REGISTRO DE VALORES VALIDOS EN TABLA DE TIPOS DE CUENTAS
                CargarValoresUDT("MSS_TIPOCUENTA", "C", "Corriente");
                CargarValoresUDT("MSS_TIPOCUENTA", "A", "Ahorro");
                CargarValoresUDT("MSS_TIPOCUENTA", "M", "Maestra");

                // REGISTRO DE VALORES VALIDOS EN TABLA DE TIPOS DE SN
                CargarValoresUDT("MSS_TIPOSN", "D", "Carnet Diplomatico");
                CargarValoresUDT("MSS_TIPOSN", "M", "Libreta Militar");
                CargarValoresUDT("MSS_TIPOSN", "E", "Carnet Ext.");
                CargarValoresUDT("MSS_TIPOSN", "P", "Pasaporte");
                CargarValoresUDT("MSS_TIPOSN", "J", "Juzgado/Resolucion");
                CargarValoresUDT("MSS_TIPOSN", "R", "RUC");
                CargarValoresUDT("MSS_TIPOSN", "L", "Libreta Electoral/DNI");
                CargarValoresUDT("MSS_TIPOSN", "S", "Sin Documento");

                // REGISTRO DE VALORES VALIDOS EN TABLA DE BANCOS VALIDOS
                CargarValoresUDT("MSS_BANCOVALIDO", "1", "Banco de Crédito del Peru");
                CargarValoresUDT("MSS_BANCOVALIDO", "2", "Banco Scotiabank");
                CargarValoresUDT("MSS_BANCOVALIDO", "3", "Banco Contineltal BBVA");
                CargarValoresUDT("MSS_BANCOVALIDO", "4", "Banco de la Nacion (DET)");
                CargarValoresUDT("MSS_BANCOVALIDO", "5", "Banco Interbank");
                CargarValoresUDT("MSS_BANCOVALIDO", "6", "Banco BanBif");
                CargarValoresUDT("MSS_BANCOVALIDO", "7", "Banco de Crédito del Peru II");
                CargarValoresUDT("MSS_BANCOVALIDO", "8", "Banco Contineltal BBVA II");
                CargarValoresUDT("MSS_BANCOVALIDO", "9", "Banco Scotiabank - planillas");
                CargarValoresUDT("MSS_BANCOVALIDO", "10", "Banco Pichincha");

                PreloadUDO("MSSL_PMD", null, new string[] { "Code", "Name", "U_MSSL_CAD", "U_MSSL_CPP", "U_MSSL_CPD", "U_MSSL_CPA" }, new string[] { "MSSL", "MSSL", "N", "N", "N", "N" }, null, null);

            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar autorizaciones del AddOn en la sociedad de SAP B1.
        /// </summary>
        private void CrearAutorizacionesAddOn()
        {
            try
            {
                //RegistrarAutorizaciones("MSS_PERM_APMD_0001", "AddOn MSS Asistente de Pagos Masivos y Detracciones", PermissionType.pt_father, "", "");
                //RegistrarAutorizaciones("MSS_PERM_APMD_0002", "Asistente de Pagos Masivos y Detracciones", PermissionType.pt_child, "MSS_PERM_APMD_0001", "");            
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }



        #endregion

        #region _Base Methods_

        /// <summary>
        /// Método para crear objetos de tipo Tablas de Usuarios (UDT) en la sociedad conectada.
        /// </summary>
        /// <param name="NombTabla">Nombre de la tabla de usuario.</param>
        /// <param name="DescTabla">Descripción de la tabla de usuario.</param>
        /// <param name="tipoTabla">Tipo de tabla de usuario.</param>
        /// <returns>Returna TRUE o FALSE como resultado de la operación de registro.</returns>
        private bool CreaTablaMD(string NombTabla, string DescTabla, SAPbobsCOM.BoUTBTableType tipoTabla)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;

            try
            {
                oUserTablesMD = null;
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!oUserTablesMD.GetByKey(NombTabla))
                {
                    oUserTablesMD.TableName = NombTabla;
                    oUserTablesMD.TableDescription = DescTabla;
                    oUserTablesMD.TableType = tipoTabla;
                    m_iErrCode = oUserTablesMD.Add();

                    if (m_iErrCode != 0)
                    {
                        Conexion.Conexion_SBO.m_oCompany.GetLastError(out m_iErrCode, out m_sErrMsg);
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error al crear  tabla: " + NombTabla,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    else
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Se ha creado la tabla: " + NombTabla,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    FuncionesComunes.LiberarObjetoGenerico(oUserTablesMD);
                    oUserTablesMD = null;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserTablesMD);
                oUserTablesMD = null;
            }
        }

        /// <summary>
        /// Método para crear objetos de tipo Campos de Usuarios (UDF) en la sociedad conectada.
        /// </summary>
        /// <param name="NombreTabla">Nombre de la tabla donde se creará el campo.</param>
        /// <param name="NombreCampo">Nombre del campo de usuario.</param>
        /// <param name="DescCampo">Descripción del campo de usario</param>
        /// <param name="TipoCampo">Tipo de campo de usuario.</param>
        /// <param name="SubTipo">Subtipo de campo de usuario.</param>
        /// <param name="Tamano">Tamaño del campo de usuario.</param>
        /// <param name="Obligatorio">Indicador de si el campo es obligatorio o no.</param>
        /// <param name="validValues">Arreglo de valores validos.</param>
        /// <param name="validDescription">Arreglo de descripción de valores validos.</param>
        /// <param name="valorPorDef">Valor por defecto.</param>
        /// <param name="tablaVinculada">Tabla vinculada al campo de usuario.</param>
        private void CreaCampoMD(string NombreTabla, string NombreCampo, string DescCampo, SAPbobsCOM.BoFieldTypes TipoCampo,
           SAPbobsCOM.BoFldSubTypes SubTipo, int Tamano, SAPbobsCOM.BoYesNoEnum Obligatorio, string[] validValues,
            string[] validDescription, string valorPorDef, string tablaVinculada)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;

            try
            {
                if (NombreTabla == null) NombreTabla = "";
                if (NombreCampo == null) NombreCampo = "";
                if (Tamano == 0) Tamano = 10;
                if (validValues == null) validValues = new string[0];
                if (validDescription == null) validDescription = new string[0];
                if (valorPorDef == null) valorPorDef = "";
                if (tablaVinculada == null) tablaVinculada = "";

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NombreTabla;
                oUserFieldsMD.Name = NombreCampo;
                oUserFieldsMD.Description = DescCampo;
                oUserFieldsMD.Type = TipoCampo;
                if (TipoCampo != SAPbobsCOM.BoFieldTypes.db_Date) oUserFieldsMD.EditSize = Tamano;
                oUserFieldsMD.SubType = SubTipo;

                if (tablaVinculada != "") oUserFieldsMD.LinkedTable = tablaVinculada;
                else
                {
                    if (validValues.Length > 0)
                    {
                        for (int i = 0; i <= (validValues.Length - 1); i++)
                        {
                            oUserFieldsMD.ValidValues.Value = validValues[i];
                            if (validDescription.Length > 0) oUserFieldsMD.ValidValues.Description = validDescription[i];
                            else oUserFieldsMD.ValidValues.Description = validValues[i];
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    oUserFieldsMD.Mandatory = Obligatorio;
                    if (valorPorDef != "") oUserFieldsMD.DefaultValue = valorPorDef;
                }

                m_iErrCode = oUserFieldsMD.Add();

                if (m_iErrCode != 0)
                {
                    Conexion.Conexion_SBO.m_oCompany.GetLastError(out m_iErrCode, out m_sErrMsg);
                    if ((m_iErrCode != -5002) && (m_iErrCode != -2035))
                        Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error al crear campo de usuario: " + NombreCampo
                            + "en la tabla: " + NombreTabla + " Error: " + m_sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                    Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Se ha creado el campo de usuario: " + NombreCampo
                            + " en la tabla: " + NombreTabla, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserFieldsMD);
                oUserFieldsMD = null;
            }
        }

        /// <summary>
        /// Método para crear objetos de tipo Objetos de Usuarios (UDO) en la sociedad conectada.
        /// </summary>
        /// <param name="s_Code"></param>
        /// <param name="s_Name"></param>
        /// <param name="s_TableName"></param>
        /// <param name="s_FindColumns"></param>
        /// <param name="s_ChildTables"></param>
        /// <param name="e_CanCancel"></param>
        /// <param name="e_CanClose"></param>
        /// <param name="e_CanDelete"></param>
        /// <param name="e_CanCreateDefaultForm"></param>
        /// <param name="s_FormColumns"></param>
        /// <param name="e_CanFind"></param>
        /// <param name="e_CanLog"></param>
        /// <param name="e_ObjectType"></param>
        /// <param name="e_ManageSeries"></param>
        /// <param name="e_EnableEnhancedForm"></param>
        /// <param name="e_RebuildEnhancedForm"></param>
        /// <param name="s_ChildFormColumns"></param>
        /// <param name="i_ChildBlock"></param>
        /// <returns></returns>
        private bool CreaUDOMD(string s_Code, string s_Name, string s_TableName, string[] s_FindColumns = null,
            string[] s_ChildTables = null, SAPbobsCOM.BoYesNoEnum e_CanCancel = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_CanClose = SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum e_CanDelete = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO, string[] s_FormColumns = null,
            SAPbobsCOM.BoYesNoEnum e_CanFind = SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum e_CanLog = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoUDOObjType e_ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData,
            SAPbobsCOM.BoYesNoEnum e_ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            string[] s_ChildFormColumns = null, int[] iChildBlock = null
            )
        {

            /* ,
            SAPbobsCOM.BoYesNoEnum e_EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            SAPbobsCOM.BoYesNoEnum e_RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO,
            string[] s_ChildFormColumns = null, int[] iChildBlock= null */

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            int i_Result = 0;
            string s_Result = "";
            int beginChild = 0;

            try
            {
                oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                oUserObjectsMD.Code = "";


                if (!oUserObjectsMD.GetByKey(s_Code))
                {
                    oUserObjectsMD.Code = s_Code;
                    oUserObjectsMD.Name = s_Name;
                    oUserObjectsMD.ObjectType = e_ObjectType;
                    oUserObjectsMD.TableName = s_TableName;
                    oUserObjectsMD.CanCancel = e_CanCancel;
                    oUserObjectsMD.CanClose = e_CanClose;
                    oUserObjectsMD.CanDelete = e_CanDelete;
                    oUserObjectsMD.CanCreateDefaultForm = e_CanCreateDefaultForm;
                    oUserObjectsMD.EnableEnhancedForm = e_EnableEnhancedForm;
                    oUserObjectsMD.RebuildEnhancedForm = e_RebuildEnhancedForm;
                    oUserObjectsMD.CanFind = e_CanFind;
                    oUserObjectsMD.CanLog = e_CanLog;
                    oUserObjectsMD.ManageSeries = e_ManageSeries;

                    if (s_FindColumns != null)
                    {
                        for (int i = 0; i < s_FindColumns.Length - 1; i++)
                        {
                            oUserObjectsMD.FindColumns.ColumnAlias = s_FindColumns[i].ToString();
                            oUserObjectsMD.FindColumns.Add();
                        }
                    }

                    if (s_ChildTables != null)
                    {
                        for (int j = 0; j < s_ChildTables.Length; j++)
                        {
                            oUserObjectsMD.ChildTables.TableName = s_ChildTables[j];
                            oUserObjectsMD.FindColumns.Add();
                        }
                    }

                    if (s_FormColumns != null)
                    {
                        oUserObjectsMD.UseUniqueFormType = SAPbobsCOM.BoYesNoEnum.tYES;

                        for (int k = 0; k < s_FormColumns.Length; k++)
                        {
                            oUserObjectsMD.FormColumns.FormColumnAlias = s_FormColumns[k];
                            oUserObjectsMD.FormColumns.Add();
                        }
                    }

                    if (s_ChildFormColumns != null)
                    {
                        if (s_ChildTables != null)
                        {
                            beginChild = 1;

                            for (int i = 0; i < s_ChildFormColumns.Length; i++)
                            {
                                oUserObjectsMD.FormColumns.SonNumber = beginChild;
                                oUserObjectsMD.FormColumns.FormColumnAlias = s_ChildFormColumns[i];
                                oUserObjectsMD.FormColumns.Add();

                                if (iChildBlock[(beginChild - 1)].Equals((i + 1)))
                                {
                                    beginChild = beginChild + 1;
                                }
                            }
                        }
                    }

                    i_Result = oUserObjectsMD.Add();

                    if (!i_Result.Equals(0))
                    {
                        Conexion.Conexion_SBO.m_oCompany.GetLastError(out i_Result, out s_Result);
                        FuncionesComunes.DisplayErrorMessages((i_Result.ToString() + " - " + s_Result).ToString(), System.Reflection.MethodBase.GetCurrentMethod());
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                return false;
            }
            finally
            {
                FuncionesComunes.LiberarObjetoGenerico(oUserObjectsMD);
                oUserObjectsMD = null;
            }
        }

        /// <summary>
        /// Método para registrar valores validos en tablas de usuarios dentro de SAP B1.
        /// </summary>
        /// <param name="s_TableName">Nombre de la tabla de usuario.</param>
        /// <param name="s_CodeValue">Código del valor, representado en la columna Code de la tabla de usuario.</param>
        /// <param name="s_NameValue">Nombre o descripción del valor, representado en la columna Name de la tabla de usuario.</param>
        private void CargarValoresUDT(string s_TableName, string s_CodeValue, string s_NameValue)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            int i_Result = 0;
            int i_Error = 0;
            string s_Error = "";

            try
            {
                oUserTable = Conexion.Conexion_SBO.m_oCompany.UserTables.Item(s_TableName);
                if (!oUserTable.GetByKey(s_CodeValue))
                {
                    oUserTable.Code = s_CodeValue;
                    oUserTable.Name = s_NameValue;
                    i_Result = oUserTable.Add();

                    if (i_Result != 0)
                    {
                        Conexion.Conexion_SBO.m_oCompany.GetLastError(out i_Error, out s_Error);
                        FuncionesComunes.DisplayErrorMessages((i_Error.ToString() + s_Error).ToString(), System.Reflection.MethodBase.GetCurrentMethod());
                    }
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        

        private void PreloadUDO(string udoName, string udoChildName, string[] columnsUDOMain, string[] valuesUDOMain, string[] columnsUDOChild, string[] valuesUDOChid)
        {
            Metadata m_oMetadata = null;
            string errorMessage = "";

            try
            {
                m_oMetadata = new Metadata(Conexion.Conexion_SBO.m_oCompany);

                if (!m_oMetadata.PreloadUDO(udoName, udoChildName, columnsUDOMain, valuesUDOMain, columnsUDOChild,
                                            valuesUDOChid))
                {
                    if (m_oMetadata.IsError)
                    {
                        errorMessage = "[" + m_oMetadata.ErrorCode.ToString() + "] " + m_oMetadata.ErrorMessage;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                m_oMetadata = null;
            }
        }


        #endregion
    }
}
