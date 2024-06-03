using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnFEData.Comunes
{
    class Metadata
    {
        #region _Attributes_

        private SAPbobsCOM.Company oCompany;
        private int m_iErrCode = 0;
        private string m_sErrMsg = "";
        private bool m_bErr = false;

        #endregion

        #region _Constructor_

        public Metadata(SAPbobsCOM.Company oCmpny)
        {
            oCompany = oCmpny;
        }

        #endregion

        #region _Properties_

        public int ErrorCode
        {
            get { return m_iErrCode; }
        }
        public string ErrorMessage
        {
            get { return m_sErrMsg; }
        }
        public bool IsError
        {
            get { return m_bErr; }
        }

        #endregion

        #region _Metodos_

        

        public bool ValidRS(string query)
        {
            SAPbobsCOM.Recordset oRecordset = null;

            try
            {
                m_bErr = false;

                ValidarEstructuraDeConsulta(ref query);

                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(query);

                if (oRecordset.RecordCount > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                m_bErr = true;
                m_iErrCode = -9999;
                m_sErrMsg = ex.Message;
                return false;
            }
            finally
            {
                Comunes.FuncionesComunes.LiberarObjetoGenerico(oRecordset);
                oRecordset = null;
            }
        }


        private void ValidarEstructuraDeConsulta(ref string sQuery)
        {
            //Este método quita el último punto y coma de la consulta, el cual es un caracter no válido para algunas versiones de HANA.
            try
            {
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    if (true)
                    { // Versión de referencia tomado inicialmente del cliente EXPLORA
                        if (sQuery.LastIndexOf(';') != -1)
                        {
                            sQuery = sQuery.Remove(sQuery.LastIndexOf(';'));
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        


        public bool PreloadRS(string query)
        {
            SAPbobsCOM.Recordset oRecordset = null;
            try
            {
                m_bErr = false;

                ValidarEstructuraDeConsulta(ref query);

                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(query);

                return true;
            }
            catch (Exception ex)
            {
                m_bErr = true;
                m_iErrCode = -9999;
                m_sErrMsg = ex.Message;
                return false;
            }
            finally
            {
                Comunes.FuncionesComunes.LiberarObjetoGenerico(oRecordset);
                oRecordset = null;
            }
        }


        public bool PreloadUDO(string udoName, string udoChildName, string[] columnsUDOMain, string[] valuesUDOMain, string[] columnsUDOChild, string[] valuesUDOChid)
        {
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oUDOMain = null;
            SAPbobsCOM.GeneralData oUDOChild = null;
            SAPbobsCOM.GeneralDataCollection oUDOChildren = null;

            try
            {
                m_bErr = false;

                oCompanyService = (SAPbobsCOM.CompanyService)oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(udoName);
                oUDOMain = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                for (int i = 0; i < columnsUDOMain.Length; i++)
                {
                    oUDOMain.SetProperty(columnsUDOMain[i], valuesUDOMain[i]);
                }

                if (udoChildName != null)
                {
                    oUDOChildren = oUDOMain.Child(udoChildName);
                    oUDOChild = oUDOChildren.Add();

                    for (int i = 0; i < columnsUDOChild.Length; i++)
                    {
                        oUDOChild.SetProperty(columnsUDOChild[i], valuesUDOChid[i]);
                    }
                }

                oGeneralService.Add(oUDOMain);
                return true;
            }
            catch (Exception ex)
            {

                m_bErr = true;
                m_iErrCode = -9999;
                m_sErrMsg = ex.Message;
                return false;
            }
            finally
            {
                oCompanyService = null;
                oGeneralService = null;
                oUDOMain = null;
                oUDOChild = null;
                oUDOChildren = null;
            }
        }
        #endregion
    }
}
