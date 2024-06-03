using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnFEData.Conexion
{
    class Conexion_SBO
    {
        #region Atributos

        public static SAPbouiCOM.Application m_SBO_Appl = null;
        public static SAPbobsCOM.Company m_oCompany = null;

        #endregion

        #region Constructores

        public Conexion_SBO()
        {
            try
            {
                ObtenerAplicacion();
                ConectarCompany();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(AddOnFEData.Properties.Resources.NombreAddon + " Error: Conexion_SBO.cs > Conexion_SBO(): " + ex.Message, "Aceptar",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }

        }

        #endregion

        #region Metodos

        private void ObtenerAplicacion()
        {
            try
            {
                string strConexion = ""; //variable que almacena el codigo de identificacion de conexion con SBO
                string[] strArgumentos = new string[4];
                SAPbouiCOM.SboGuiApi oSboGuiApi = null; //Variable para obtener la instacia activa de SBO

                oSboGuiApi = new SAPbouiCOM.SboGuiApi();//Instancia nueva para la gestion de la conexion
                strArgumentos = System.Environment.GetCommandLineArgs();//obtenemos el codigo de conexion del entorno configurado en "Propiedades -> Depurar -> Argumentos de la linea de comandos"

                if (strArgumentos.Length > 0)
                {
                    if (strArgumentos.Length > 1)
                    {
                        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                        if (strArgumentos[0].LastIndexOf("\\") > 0) strConexion = strArgumentos[1];
                        else strConexion = strArgumentos[0];
                    }
                    else
                    {
                        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                        if (strArgumentos[0].LastIndexOf("\\") > -1) strConexion = strArgumentos[0];
                        else
                        {
                            System.Windows.Forms.MessageBox.Show(AddOnFEData.Properties.Resources.NombreAddon + " Error en: Conexion_SBO.cs > ObtenerAplicacion(): SAP Business One no esta en ejecucion", "Aceptar",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);//mensaje de error por no tener SBO activo
                        }
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(AddOnFEData.Properties.Resources.NombreAddon + " Error en: Conexion_SBO.cs > ObtenerAplicacion(): SAP Business One no esta en ejecucion", "Aceptar",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);//mensaje de error por no tener SBO activo
                }

                oSboGuiApi.Connect(strConexion);//Establecemos la conexion
                m_SBO_Appl = oSboGuiApi.GetApplication(-1);//Asignamos la conexion a la aplicacion
            }
            catch (Exception ex)
            {
                {
                    System.Windows.Forms.MessageBox.Show(AddOnFEData.Properties.Resources.NombreAddon + " Error en: Conexion_SBO.cs > ObtenerAplicacion(): " + ex.Message, "Aceptar",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);//mensaje de error por no tener SBO activo
                }
            }
        }

        public static void ConectarCompany()
        {
            string sCookie = "", sErrMsg = "";
            int iRet = 0, iErrCode = 0;
            try
            {
                if (m_oCompany == null)
                {
                    m_oCompany = new SAPbobsCOM.Company(); //creamos una nueva instacia del objeto company
                    sCookie = m_oCompany.GetContextCookie();
                    iRet = m_oCompany.SetSboLoginContext(m_SBO_Appl.Company.GetConnectionContext(sCookie));//Conectamos a la compañia de la instacia de SBO que se esta ejecutando
                    if (iRet == 0)
                    {
                        iRet = m_oCompany.Connect();//Establecemos la conexion con la compañia
                        if (iRet != 0)//validamos que no se hayan producido errores
                        {
                            m_oCompany.GetLastError(out iErrCode, out sErrMsg);//obtenemos el error que se produjo
                            Comunes.FuncionesComunes.LiberarObjetoGenerico(m_oCompany);
                            m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error: Conexion_SBO.cs > ConectarCompany(): " + sErrMsg,
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                else
                {
                    iRet = m_oCompany.Connect();//Establecemos la conexion con la compañia
                    if (iRet != 0)//validamos que no se hayan producido errores
                    {
                        m_oCompany.GetLastError(out iErrCode, out sErrMsg);//obtenemos el error que se produjo
                        Comunes.FuncionesComunes.LiberarObjetoGenerico(m_oCompany);
                        m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error: Conexion_SBO.cs > ConectarCompany(): " + sErrMsg,
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }

            }
            catch (Exception ex)
            {
                m_oCompany.GetLastError(out iErrCode, out sErrMsg);//obtenemos el error que se produjo
                Comunes.FuncionesComunes.LiberarObjetoGenerico(m_oCompany);
                m_SBO_Appl.StatusBar.SetText(AddOnFEData.Properties.Resources.NombreAddon + " Error: Conexion_SBO.cs > ConectarCompany():" + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void DesconectarCompany()
        {
            try
            {
                m_oCompany.Disconnect();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("AddOnFEData.Properties.Resources.NombreAddon +  Error en: Conexion_SBO.cs > DesconectarCompany(): " + ex.Message, "Aceptar",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);//mensaje de error por no tener SBO activo
            }
        }

        #endregion
    }
}
