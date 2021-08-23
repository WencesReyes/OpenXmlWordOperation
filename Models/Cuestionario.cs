using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OpenXMLWordOperations.Models
{
    public class Cuestionario
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int NumDocumento { get; set; }
        public string Nacionalidad { get; set; }
        public string Cargo { get; set; }
        public string Correo { get; set; }
        public string Telefono { get; set; }
        public string Sexo { get; set; }
        public string Domicilio { get; set; }
        public DateTime DOB { get; set; }
    }
}
