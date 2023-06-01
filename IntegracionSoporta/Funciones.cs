using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Data.Odbc;
using System.Globalization;

namespace IntegracionSoporta
{
    public class Funciones
    {
        public void ObtenerCredenciales(ref Credenciales credenciales)
        {
            XmlNodeList Configuracion;
            XmlNodeList lista;
            String sPath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory); //this.GetType().Assembly.Location); System.Reflection.Assembly.GetExecutingAssembly().Location
            XmlDocument xDoc;

            try
            {
                string contenido = String.Empty;
                contenido = File.ReadAllText(sPath + "\\Config.xml");

                xDoc = new XmlDocument();

                xDoc.Load(sPath + "\\Config.xml");

                Configuracion = xDoc.GetElementsByTagName("Configuracion");

                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("ServidorSAP");

                foreach (XmlElement nodo in lista)
                {
                    var i = 0;
                    var sServer = nodo.GetElementsByTagName("Servidor");
                    var sLicencia = nodo.GetElementsByTagName("ServidorLicencia");
                    var sUserDB = nodo.GetElementsByTagName("UsuarioDB");
                    var sPassDB = nodo.GetElementsByTagName("PasswordDB");
                    var nPass = nodo.GetElementsByTagName("PasswordSAP");
                    var nUser = nodo.GetElementsByTagName("UsuarioSAP");
                    var nTipoDB = nodo.GetElementsByTagName("TipoDB");
                    var nMostrarLineError = nodo.GetElementsByTagName("MostrarLineError");
                    var nBaseDatos = nodo.GetElementsByTagName("BaseDatos");
                    var nArticulo = nodo.GetElementsByTagName("Articulo");
                    var nFecha = nodo.GetElementsByTagName("FechaInicio");

                    credenciales.servidor = (String)(sServer[i].InnerText);
                    credenciales.servidorLicencia = (String)(sLicencia[i].InnerText);
                    credenciales.usuarioDB = (String)(sUserDB[i].InnerText);
                    credenciales.passwordDB = (String)(sPassDB[i].InnerText);
                    credenciales.usuarioSAP = (String)(nUser[i].InnerText);
                    credenciales.passwordSAP = (String)(nPass[i].InnerText);
                    credenciales.tipoDB = (String)(nTipoDB[i].InnerText);
                    credenciales.mostrarLineError = (((String)nMostrarLineError[i].InnerText) == "Y" ? true : false);
                    credenciales.baseDatos = (String)(nBaseDatos[i].InnerText);
                    credenciales.articulo = (String)(nArticulo[i].InnerText);

                    var fec = (String)(nFecha[i].InnerText);
                    if (fec == "")
                        credenciales.fechaInicio = DateTime.Now.Date;
                    else
                        credenciales.fechaInicio = DateTime.Parse(fec, CultureInfo.InvariantCulture, DateTimeStyles.None);
                }

                lista = ((XmlElement)Configuracion[0]).GetElementsByTagName("Soporta");
                foreach (XmlElement nodo in lista)
                {
                    var i = 0;
                    var sServer = nodo.GetElementsByTagName("Servidor");
                    var sUsuario = nodo.GetElementsByTagName("Usuario");
                    var sPassword = nodo.GetElementsByTagName("Password");
                    var sBase = nodo.GetElementsByTagName("Base");
                    var sTop = nodo.GetElementsByTagName("TOP");

                    credenciales.servidorSoporta = (String)(sServer[i].InnerText);
                    credenciales.usuarioSoporta = (String)(sUsuario[i].InnerText);
                    credenciales.passwordSoporta = (String)(sPassword[i].InnerText);
                    credenciales.baseSoporta = (String)(sBase[i].InnerText);
                    credenciales.top = ((String)(sTop[i].InnerText) == "" ? 20 : Convert.ToInt32((String)(sTop[i].InnerText)));
                }

            }
            catch (Exception ex)
            {
                AddLog("Error al Obtener Credenciales XML: " + ex.Message + ", Trace: " + ex.StackTrace);
            }
        }

        //Funcion registra log
        public void AddLog(String Mensaje, String Prefijo = "")
        {
            StreamWriter Arch;
            //Exe: String := 
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String NomArch;
            String NomArchB;
            NomArch = "\\LOGs\\ProgramLog_" + (Prefijo.Trim().Length > 0 ? Prefijo.Trim() + "_" : "") + String.Format("{0:yyyy-MM-dd}", DateTime.Now) + ".log";
            Arch = new StreamWriter(sPath + NomArch, true);
            NomArchB = sPath + "\\LOGs\\ProgramLog_" + String.Format("{0:yyyy-MM-dd}", DateTime.Now.AddDays(-1)) + ".log";
            //Elimina archivo del dia anterior
            //if (System.IO.File.Exists(NomArchB))
            //    System.IO.File.Delete(NomArchB);

            try
            {
                Arch.WriteLine(String.Format("{0:dd-MM-yyyy HH:mm:ss}", DateTime.Now) + " " + Mensaje);
            }
            finally
            {
                Arch.Flush();
                Arch.Close();
            }
        }

        public Boolean ConectarSAP(ref Credenciales credenciales, ref SAPbobsCOM.Company oCompany, ref ListBox lbText)
        {
            string sErrMsg = null;
            int lErrCode = 0;
            int lRetCode;

            try
            {
                if (credenciales.tipoDB == "2019")
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                else
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

                oCompany.UseTrusted = false;
                oCompany.Server = credenciales.servidor;
                oCompany.LicenseServer = credenciales.servidorLicencia;
                oCompany.DbUserName = credenciales.usuarioDB;
                oCompany.DbPassword = credenciales.passwordDB;

                oCompany.CompanyDB = credenciales.baseDatos;
                oCompany.UserName = credenciales.usuarioSAP;
                oCompany.Password = credenciales.passwordSAP;

                // Connecting to a company DB
                lRetCode = oCompany.Connect();

                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lRetCode, out sErrMsg);
                    AddLog("ConexionSAPB1: Error al conectar DIAPI - " + sErrMsg);
                    lbText.Items.Add("ConexionSAPB1: Error al conectar DIAPI - " + sErrMsg);
                    return false;
                }
                else
                {
                    AddLog("Conectado a " + credenciales.baseDatos);
                    lbText.Items.Add("Conectado a " + oCompany.CompanyName);
                    return true;
                }
            }
            catch (Exception ex)
            {

                return false;
            }

        }

        public Boolean ConectarSoporta(ref Credenciales credenciales, ref ListBox lbText, ref OdbcConnection connSql)
        {
            try
            {
                connSql = new OdbcConnection();
                connSql.ConnectionString =
                              "Driver={SQL Server};" +
                              @"Server=" + credenciales.servidorSoporta + ";" +
                              "DataBase=" + credenciales.baseSoporta + ";" +
                              "Uid=" + credenciales.usuarioSoporta + ";" +
                              "Pwd=" + credenciales.passwordSoporta + ";";
                connSql.Open();
                connSql.Close();
                lbText.Items.Add("Conectado a Soporta");
                return true;

            }
            catch (Exception ex)
            {
                lbText.Items.Add("ConectarSoporta: Error al conectar - " + ex.Message);
                return false;
            }

        }

    }
}
