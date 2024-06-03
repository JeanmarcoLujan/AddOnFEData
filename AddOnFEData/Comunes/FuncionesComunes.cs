using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnFEData.Comunes
{
    class FuncionesComunes
    {
        #region Metodos

        public static void LiberarObjetoGenerico(Object objeto)
        {
            try
            {
                if (objeto != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(AddOnFEData.Properties.Resources.NombreAddon + " Error Liberando Objeto: " + ex.Message);
            }
        }

        /// <summary>
        /// Método para mostrar los errores a través de la barra de mansajes de SAP B1 en la sociedad conectada.
        /// </summary>
        /// <param name="messageError">Descripción del mensaje de error.</param>
        /// <param name="oMethodBase">Objeto Reflection con el detalle de la clase donde se generó el error.</param>
        public static void DisplayErrorMessages(string messageError, System.Reflection.MethodBase oMethodBase)
        {
            string className = string.Empty;
            string methodName = string.Empty;

            try
            {
                className = oMethodBase.DeclaringType.Name.ToString();
                methodName = oMethodBase.Name.ToString();

                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Properties.Resources.NombreAddon +
                    " Error: " + className + ".cs > " + methodName + "(): " + messageError,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            catch (Exception ex)
            {
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(Properties.Resources.NombreAddon + " Error: FuncionesComunes.cs > DisplayErrorMessages(): "
                    + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        



        #endregion
    }
}
