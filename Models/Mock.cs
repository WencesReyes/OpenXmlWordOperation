using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OpenXMLWordOperations.Models
{
    public class Mock
    {
        public static IEnumerable<Cuestionario> GetCuestionariosMock()
        {
            List<Cuestionario> cuestionarios = new List<Cuestionario>
            {
                new Cuestionario
                {
                    Nombre = "Wenceslao",
                    Apellido = "Reyes",
                    Cargo = "Developer .NET",
                    Correo = "wenceslao@gmail.com",
                    DOB = new DateTime(1997,8,27),
                    Domicilio = "Chapulin #1435",
                    Nacionalidad = "Mexicano",
                    NumDocumento = 12,
                    Sexo = "M",
                    Telefono = "3124562311"
                },
                new Cuestionario
                {
                    Nombre = "Chuy",
                    Apellido = "García",
                    Cargo = "Developer JS",
                    Correo = "chuy@gmail.com",
                    DOB = new DateTime(1996,3,25),
                    Domicilio = "Alberto Issac #1254",
                    Nacionalidad = "Mexicano",
                    NumDocumento = 15,
                    Sexo = "M",
                    Telefono = "3124561123"
                },
            };

            return cuestionarios;
        }
    }
}
