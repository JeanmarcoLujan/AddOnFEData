using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOnFEData.Comunes
{
    class Eventos_SBO
    {
        #region Constructores

        /// <summary>
        /// Constructor de la clase.
        /// </summary>
        public Eventos_SBO()
        {
            try
            {
                Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title.Replace(" #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString(), "");
                Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title + " #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString();
                RegistrarEventos();
                RegistrarFiltros();
                RegistrarMenu();
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Conectado con exito",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }

        #endregion

        #region Eventos

        void m_SBO_Appl_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                        RegistrarMenu();
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
        void m_SBO_Appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //switch (pVal.FormTypeEx)
                //{
                //    case "Frm_APMD_P1":
                //        Formularios.Frm_APMD_P1 oFrm_APMD_P1 = null;
                //        oFrm_APMD_P1 = new MSS_Asistente_PMD.Formularios.Frm_APMD_P1(false);
                //        oFrm_APMD_P1.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                //        oFrm_APMD_P1 = null;
                //        break;
                //}
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
        void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion

        #region Metodos

        /// <summary>
        /// Método para registrar la opción del menú dentro del formulario de menus de SAP B1.
        /// </summary>
        private void RegistrarMenu()
        {
            try
            {
                //CreaMenu("MSS_APMD", "Asistente Pagos Masivos y Detracciones", "43538", SAPbouiCOM.BoMenuType.mt_STRING);
                CreaMenu("MSS_MPMD", "Pagos Masivos y Detracciones", "43538", SAPbouiCOM.BoMenuType.mt_POPUP); //43520
                CreaMenu("MSS_CPMD", "Definiciones PMD", "MSS_MPMD", SAPbouiCOM.BoMenuType.mt_STRING);
                CreaMenu("MSS_APMD", "Asistente Pagos Masivos y Detracciones", "MSS_MPMD", SAPbouiCOM.BoMenuType.mt_STRING);//43538

            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar los eventos de la aplicación en SAP B1.
        /// </summary>
        private void RegistrarEventos()
        {
            try
            {
                Conexion.Conexion_SBO.m_SBO_Appl.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(m_SBO_Appl_AppEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(m_SBO_Appl_FormDataEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(m_SBO_Appl_ItemEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(m_SBO_Appl_MenuEvent);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar los filtros en SAP B1.
        /// </summary>
        private void RegistrarFiltros()
        {
            SAPbouiCOM.EventFilter oEF = null;
            SAPbouiCOM.EventFilters oEFs = null;
            try
            {
                oEFs = new SAPbouiCOM.EventFilters();
                oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
                oEF = oEFs.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                oEF.AddEx("Frm_APMD_P1");
                oEF.AddEx("Frm_APMD_P2");
                
                Conexion.Conexion_SBO.m_SBO_Appl.SetFilter(oEFs);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar opciones de menú en el menu principal de SAP B1.
        /// </summary>
        /// <param name="uniqueId"></param>
        /// <param name="name"></param>
        /// <param name="principalMenuId"></param>
        /// <param name="type"></param>
        private void CreaMenu(string uniqueId, string name, string principalMenuId, SAPbouiCOM.BoMenuType type)
        {
            SAPbouiCOM.MenuCreationParams objParams;
            SAPbouiCOM.Menus objSubMenu;

            try
            {
                objSubMenu = Conexion.Conexion_SBO.m_SBO_Appl.Menus.Item(principalMenuId).SubMenus;

                if (Conexion.Conexion_SBO.m_SBO_Appl.Menus.Exists(uniqueId) == false)
                {
                    objParams = (SAPbouiCOM.MenuCreationParams)Conexion.Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objParams.Type = type;
                    objParams.UniqueID = uniqueId;
                    objParams.String = name;
                    objParams.Position = -1;
                    objSubMenu.AddEx(objParams);
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion
    }
}
