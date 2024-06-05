using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOnFEData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Conexion.Conexion_SBO oConexion = new AddOnFEData.Conexion.Conexion_SBO();

            if ((oConexion != null) && (Conexion.Conexion_SBO.m_oCompany.Connected))
            {
                Conexion.Conexion_SBO.m_oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Comunes.EstructuraDatos oEstructuraDatos = new Comunes.EstructuraDatos();
                Comunes.Eventos_SBO oEventos = new AddOnFEData.Comunes.Eventos_SBO();
                GC.KeepAlive(oConexion);
                GC.KeepAlive(oEventos);
                Application.Run();
            }
            else
                Application.Exit();
        }
    }
}
