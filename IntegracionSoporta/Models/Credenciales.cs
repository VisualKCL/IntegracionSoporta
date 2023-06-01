using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegracionSoporta
{
    public class Credenciales
    {
        public string servidor { get; set; }
        public string servidorLicencia { get; set; }
        public string usuarioSAP { get; set; }
        public string passwordSAP { get; set; }
        public string usuarioDB { get; set; }
        public string passwordDB { get; set; }
        public string baseDatos { get; set; }
        public string tipoDB { get; set; }
        public Boolean mostrarLineError { get; set; }
        public string servidorSoporta { get; set; }
        public string usuarioSoporta { get; set; }
        public string passwordSoporta { get; set; }
        public string baseSoporta { get; set; }
        public string articulo { get; set; }
        public DateTime fechaInicio { get; set; }
        public Int32 top { get; set; }

    }
}
