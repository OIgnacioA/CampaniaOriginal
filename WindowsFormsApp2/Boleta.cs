using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class Boleta
    {
        public string Mail { get; set; }

        public string RazonSocial { get; set; }

        public string Cuit { get; set; }

        List<string> DatosObjetos { get; set; }

        int cantidad { get; set; }
        public int Cantidad { get { return cantidad; } }

        public Boleta(string pMail, string pRazonSocial, string pCuit)
        {
            this.Mail = pMail;
            this.RazonSocial = pRazonSocial;
            this.Cuit = pCuit;
            this.DatosObjetos = new List<string>();
        }

        public bool Igual(string pMail, string pRazonSocial, string pCuit)
        {
            return ((Mail == pMail)&&(RazonSocial == pRazonSocial)&&(Cuit == pCuit));
        }

        public bool Igual(string pMail)
        {
            return (Mail == pMail);
        }

        public void Agregar(string pDatos)
        {
            this.DatosObjetos.Add(pDatos);
            this.cantidad++;
        }

        public string Datos
        {
            get
            {
                string resultado = string.Empty;
                foreach (string item in DatosObjetos)
                {
                    resultado += item;
                }
                return resultado;
            }
        }
    }
}
