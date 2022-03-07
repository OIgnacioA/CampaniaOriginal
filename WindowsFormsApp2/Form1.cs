using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        string txtOrigen = string.Empty;
        string txtDestino = string.Empty;

        string mailAux = string.Empty;
        string razonsocialAux = string.Empty;
        string cuitAux = string.Empty;

        string mail = string.Empty;
        string objeto = string.Empty;
        string objetoFormateado = string.Empty;
        string razonsocial = string.Empty;
        string porcentaje = string.Empty;
        string anio = string.Empty;
        string cuota = string.Empty;
        string cuotaNumero = string.Empty;
        string fechaVencimiento = string.Empty;
        string fechaVencimientoNumero = string.Empty;
        string montoCuota = string.Empty;
        string montoAnual = string.Empty;
        //string codigoElectronico = string.Empty;
        string debitoCredito = string.Empty;
        string buenContribuyente = string.Empty;
        string cuit = string.Empty;
        string cuitFormateado = string.Empty;
        string medioPago = string.Empty;
        string plantas = string.Empty;
        string impuesto = string.Empty;
        string fechaOpcion = string.Empty;
        string descuento = string.Empty;

        string nombreImpuesto = string.Empty;
        
        string datosObjeto = string.Empty;

        string impuestoLiquidar = string.Empty;

        string directorioOrigen = @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\";
        string directorioDestino = @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Destino\";

        public Form1()
        {
            InitializeComponent();
            habilitarGenerar();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool seguir = false;
            try
            {
                int cantidad = Convert.ToInt32(this.txtCantidad.Text);
                seguir = true;
            }
            catch (Exception) { }

            if (seguir)
            {                
                impuesto = this.Impuesto.Text;
                if (ConAnual.Checked)
                {
                    cuota = this.Cuota.Text + " y Saldo Anual";
                }
                else
                {
                    cuota = this.Cuota.Text;
                }

                anio = this.Anio.Text.Trim();
                descuento = this.txtDescuento.Text;
                txtDestino = this.Origen.FileName;
                txtDestino = txtDestino.Replace(this.Origen.InitialDirectory, "");
                this.procesar();
            }
            else
            {
                this.txtCantidad.Focus();
                MessageBox.Show("Ingrese la cantidad de suscripciones a procesar.", "Boleta Electrónica");
            }
        }        

        private void procesar()
        {
            int cantidadAleer = Convert.ToInt32(this.txtCantidad.Text);
            this.barraGenerados.Maximum = cantidadAleer;
            this.barraLeidos.Maximum = cantidadAleer;
            int counter = 0;
            int contador = 0;
            int distintos = 0;
            int escritos = 0;
            int cantidadMailIgual = 0;
            string line;
            string mailLinea = string.Empty;
            string ultimoMail = string.Empty;
             
            string datosTodosObjetos = string.Empty;
            int cantidadCorte = Convert.ToInt32(this.txtCantidadCorte.Text);
            int cantidadArchivosGenerados = 1;

            string nombreArchivoGenerado = string.Format("{0}-Parte-{1}.csv", txtDestino, cantidadArchivosGenerados);

            fechaOpcion = this.FechaOpcion.Value.ToLongDateString().Replace(",", "");
            fechaVencimiento = fechaOpcion;

            StreamWriter sw = new System.IO.StreamWriter(nombreArchivoGenerado);
            StreamWriter swResultados = new System.IO.StreamWriter(txtDestino + "-Informe.txt");
            
            //genero reporte.
            swResultados.Write("Se generearon los siguientes archivos:");
            swResultados.WriteLine();
            swResultados.Write(string.Format("Archivo ** {0} ** ", nombreArchivoGenerado));

            

            this.EscribirCabecera(sw);

            System.IO.StreamReader file = new System.IO.StreamReader(txtOrigen);

            while ((line = file.ReadLine()) != null)
            {
                this.LeerLinea(line);
                
                //veo si es la primer linea.
                if (mailAux == string.Empty)
                {
                    mailAux = mail;
                    razonsocialAux = razonsocial;
                    cuitAux = cuit;
                    //ultimoMail = mail;
                }

                this.ArmarDatosMail();
                
                //veo si hay que agrupar el mail o no.
                if ((mail != mailAux) || (razonsocial != razonsocialAux))// || (cuit != cuitAux))
                {
                    //por las dudas que el emblue pueda usarse.
                    if ((this.DiferenciarMails.Checked) && (mailAux == ultimoMail))
                    {
                         cantidadMailIgual++;
                    }
                    else
                    {
                         cantidadMailIgual = 0;
                    }
                    
                    if (cantidadMailIgual == 0)
                    {
                        mailLinea = string.Format("{0}|", mailAux);
                    }
                    else
                    {
                        mailLinea = string.Format("{0}+{1}|", mailAux, cantidadMailIgual.ToString());
                    }


                    ultimoMail = mailAux;
                    mailLinea += string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", razonsocialAux, this.formatearCuit(cuitAux), fechaVencimiento, fechaOpcion, anio, cuota, impuesto, datosTodosObjetos, descuento);

                    if (escritos == cantidadCorte)
                    {
                        swResultados.Write(string.Format("Con {0} suscripciones y {1} mails para enviar", contador, escritos));
                        swResultados.WriteLine();

                        escritos = 0;
                        contador = 0;
                        cantidadArchivosGenerados++;
                        sw.Flush();
                        sw.Close();
                        nombreArchivoGenerado = string.Format("{0}-Parte-{1}.csv", txtDestino, cantidadArchivosGenerados);
                        sw = new System.IO.StreamWriter(nombreArchivoGenerado);
                        swResultados.Write(string.Format("Archivo ** {0} **", nombreArchivoGenerado));                        
                        this.EscribirCabecera(sw);
                    }
                    distintos++;
                    escritos++;
                    if (distintos <= this.barraGenerados.Maximum)
                    {
                        this.barraGenerados.Value = distintos;                        
                    }                    
                    sw.Write(mailLinea);
                    sw.WriteLine();

                    mailAux = mail;
                    razonsocialAux = razonsocial;
                    cuitAux = cuit;
                    
                    datosTodosObjetos = datosObjeto;                    
                    datosObjeto = string.Empty;
                }
                else
                {
                    datosTodosObjetos += datosObjeto;                    
                }
                counter++;
                contador++;
                if (counter <= this.barraLeidos.Maximum)
                {
                    this.barraLeidos.Value = counter;
                }                
            }


            //por las dudas que el emblue pueda usarse.
            if ((this.DiferenciarMails.Checked) && (mailAux == ultimoMail))
            {
                cantidadMailIgual++;
            }
            else
            {
                cantidadMailIgual = 0;
            }

            if (cantidadMailIgual == 0)
            {
                mailLinea = string.Format("{0}|", mailAux);
            }
            else
            {
                mailLinea = string.Format("{0}+{1}|", mailAux, cantidadMailIgual.ToString());
            }


            ultimoMail = mailAux;
            mailLinea += string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", this.formatearCuit(cuitAux), objetoFormateado, fechaVencimiento, fechaOpcion, anio, cuota, impuesto, datosTodosObjetos, descuento);


            sw.Write(mailLinea);
            sw.WriteLine();
            mailAux = mail;
            razonsocialAux = razonsocial;
            cuitAux = cuit;
            distintos++;
            escritos++;            

            swResultados.Write(string.Format("Con {0} suscripciones y {1} mails para enviar", contador, escritos));
            swResultados.WriteLine();

            swResultados.Flush();
            swResultados.Close();
            sw.Flush();
            sw.Close();
            file.Close();
            string mensaje = string.Empty;
            if (counter != cantidadAleer)
            {
                mensaje = string.Format("La cantidad de suscripciones configuradas ({0}) y es distinta a la cantidad de registros leidos ({1}). De todas maneras se generaron {2} mails para enviar.", cantidadAleer, counter, distintos);
                MessageBox.Show(mensaje, "Cantidad de registros ERRONEA!!");
            }
            else
            {
                mensaje = string.Format("Se leyeron {0} suscripciones y se generaron {1} mails para enviar. Armar bases?", counter.ToString(), distintos.ToString());
                if (MessageBox.Show(mensaje, "Control Totales ATENCION", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.InformarArchivosGenerados();
                }
            }
            
        }

        private string formatearCuit(string pCuit)
        {
            string cuitFormateado = string.Empty;
            if (pCuit.Length == 11)
            {
                string primeraParte = pCuit.Substring(0, 2);
                string dni = pCuit.Substring(2, 8);
                string digito = pCuit.Substring(10, 1);
                cuitFormateado = string.Format("Cuit: {0}-{1}-{2}", primeraParte, dni, digito);
            }
            return cuitFormateado;
        }

        private void InformarArchivosGenerados()
        {
            DirectoryInfo di = new DirectoryInfo(".\\");            
            string zip = string.Format("{0}.zip", txtDestino);            
            FileInfo[] archivos = di.GetFiles(txtDestino + "*");
            StreamReader r;
            using (FileStream zipToOpen = new FileStream(zip, FileMode.Create))
            {                
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    foreach (FileInfo fileToCompress in archivos)
                    {                        
                            ZipArchiveEntry readmeEntry = archive.CreateEntry(fileToCompress.Name);
                            using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
                            {
                                r = new StreamReader(fileToCompress.Open(FileMode.Open, FileAccess.Read, FileShare.None));
                                writer.WriteLine(r.ReadToEnd());
                                r.Close();
                                r.Dispose();
                            }                     
                    }
                }
            }
            foreach (FileInfo fileToCompress in archivos)
            {
                File.Delete(fileToCompress.FullName);
            }


            string mensaje = string.Format("Se generó el archivo {0}\\{1} con los datos para el envío de las campañas. Colocar dicho archivo en {2} y avisar a Mesa de ayuda.", di.FullName, zip, directorioDestino);
            MessageBox.Show(mensaje);
        }

        private void EscribirCabecera(StreamWriter pSw)
        {
            if (this.ConCabecera.Checked)
            {
                pSw.Write("email|nombre|cuit|fechav|fechao|anio|cuota|impuesto|datos|descuento");
                pSw.WriteLine();
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.Origen.ShowDialog();
            txtOrigen = this.Origen.FileName;
            this.habilitarGenerar();
        }        

        private void habilitarGenerar()
        {
            this.txtArchivoOrigen.Text = txtOrigen;            
            this.Generar.Enabled = (txtOrigen != string.Empty);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Anio.Text = DateTime.Now.Year.ToString();
        }

        private void Impuesto_SelectedIndexChanged(object sender, EventArgs e)
        {
            directorioOrigen = @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\";
            directorioDestino = @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Destino\";
            List<string> cuotas;
            switch (this.Impuesto.SelectedIndex)
            {
                case 0:
                {
                    cuotas = new List<string>(){ "1", "2", "3", "4", "5" };
                    directorioOrigen += @"Automotores\";
                    directorioDestino += @"Automotores\";
                    impuestoLiquidar = "1";
                    nombreImpuesto = "Automotor";
                        this.txtDescuento.Text = "20";

                    break;
                }
                case 1:
                {
                    cuotas = new List<string>() { "1", "2" };
                    directorioOrigen += @"Embarcaciones\";
                    directorioDestino += @"Embarcaciones\";
                    nombreImpuesto = "Embarcaciones";
                    impuestoLiquidar = "3";
                        this.txtDescuento.Text = "20";
                        break;
                    
                }
                case 2:
                {
                    cuotas = new List<string>() { "1", "2", "3", "4", "5" };
                    directorioOrigen += @"Edificado\";
                    directorioDestino += @"Edificado\";
                    nombreImpuesto = "Edificado";
                    impuestoLiquidar = "0";
                        this.txtDescuento.Text = "20";
                        break;
                }
                case 3:
                {
                    cuotas = new List<string>() { "1", "2", "3", "4"};
                    directorioOrigen += @"Baldio\";
                    directorioDestino += @"Baldio\";
                    nombreImpuesto = "Baldio";
                    impuestoLiquidar = "0";
                        this.txtDescuento.Text = "20";
                        break;
                }
                case 4:
                {
                    cuotas = new List<string>() { "1", "2", "3", "4" };
                    directorioOrigen += @"Rural\";
                    directorioDestino += @"Rural\";
                    impuestoLiquidar = "0";
                    nombreImpuesto = "Rural";
                        this.txtDescuento.Text = "20";
                        break;
                }
                case 5:
                {
                    cuotas = new List<string>() { "1", "2", "3", "4" };
                    directorioOrigen += @"Complementario\";
                    directorioDestino += @"Complementario\";
                    nombreImpuesto = "Complementario";
                    impuestoLiquidar = "10";
                        this.txtDescuento.Text = "20";
                        break;
                }

                default:
                { 
                    cuotas = new List<string>() { "0" };
                    break;
                }
            }
            this.Origen.InitialDirectory = directorioOrigen;                
            this.Cuota.DataSource = cuotas;
            this.Cuota.Refresh();
            this.Cuota.SelectedIndex = 0;

        }

        private void LeerLinea(string line)
        {            
            switch (this.Impuesto.SelectedIndex)
            {
                case 0:
                case 1:
                    {
                        mail = line.Substring(0, 255).TrimEnd(' ').ToLower();
                        objeto = line.Substring(255, 11).TrimEnd(' ');
                        objetoFormateado = objeto.ToUpper();
                        razonsocial = line.Substring(266, 60).TrimEnd(' ');
                        porcentaje = string.Empty;
                        fechaVencimiento = Convert.ToDateTime(line.Substring(334, 10).TrimEnd(' ')).ToLongDateString().Replace(",", "");
                        fechaVencimientoNumero = line.Substring(334, 10).TrimEnd(' ');
                        montoCuota = line.Substring(345, 17).Trim(' ');
                        montoAnual = line.Substring(362, 16).Trim(' ');
                        //codigoElectronico = line.Substring(378, 14).Trim(' ');
                        debitoCredito = line.Substring(392, 1).Trim(' ');
                        buenContribuyente = line.Substring(393, 1).Trim(' ');
                        cuit = line.Substring(394, 11).TrimEnd(' ');                                               
                        break;
                    }
                case 2:
                case 3:
                case 4:
                    {
                        //Leo Inmo
                        mail = line.Substring(0, 255).TrimEnd(' ').ToLower();
                        objeto = line.Substring(255, 11).TrimEnd(' ');
                        objetoFormateado = formatearObjetoInmobiliario(objeto);
                        razonsocial = line.Substring(266, 60).TrimEnd(' ');
                        porcentaje = string.Empty;
                        fechaVencimiento = Convert.ToDateTime(line.Substring(334, 10).TrimEnd(' ')).ToLongDateString().Replace(",", "");
                        fechaVencimientoNumero = line.Substring(334, 10).TrimEnd(' ');
                        montoCuota = line.Substring(345, 17).Trim(' ');
                        montoAnual = line.Substring(362, 16).Trim(' ');
                        //codigoElectronico = line.Substring(378, 14).Trim(' ');
                        debitoCredito = line.Substring(392, 1).Trim(' ');
                        buenContribuyente = line.Substring(393, 1).Trim(' ');
                        cuit = line.Substring(394, 11).TrimEnd(' ');
                        break;
                    }
                case 5:
                    {
                        //Leo Comple
                        mail = line.Substring(0, 255).TrimEnd(' ').ToLower();
                        objeto = line.Substring(255, 11).TrimEnd(' ');
                        objetoFormateado = formatearCuit(objeto); 
                        razonsocial = line.Substring(275, 60).TrimEnd(' ');                                                
                        plantas = line.Substring(347, 1).Trim(' ');
                        debitoCredito = line.Substring(348, 1).Trim(' ');
                        buenContribuyente = line.Substring(349, 1).Trim(' ');
                        cuit = objeto;
                        break;
                    }

                default:
                    
                    break;
            }
            TextInfo myTI = CultureInfo.CurrentCulture.TextInfo;
            razonsocial = myTI.ToTitleCase(razonsocial);
            
        }

        private string formatearObjetoInmobiliario(string pObjeto)
        {
            string partido = pObjeto.Substring(0, 3).TrimEnd(' ');
            string partida = pObjeto.Substring(3, 6).TrimEnd(' ');
            string digito = pObjeto.Substring(9, 1).TrimEnd(' ');
            return string.Format("{0}-{1}-{2}", partido, partida, digito);
        }       

        private void ArmarDatosMail()
        {
            switch (debitoCredito)
            {
                case "1":
                    {
                        medioPago = "Débito en Cuenta";
                        break;
                    }
                case "2":
                    {
                        medioPago = "Tarjeta de Crédito";                        
                        break;
                    }
                default:
                    {                        
                        if (this.Impuesto.SelectedIndex == 5)
                        {
                            medioPago = "<a href=\"https://sso.arba.gov.ar/Login/login?service=http%3A%2F%2Fwww4.arba.gov.ar%2FLiqPredet%2Fsso%2FInicioLiquidacionIIC.do%3FFrame%3DNO%26origen%3DWEB%26imp%3D10%26cuit%3D\">Ingresar</a>";
                        }
                        else
                        {
                            medioPago = string.Format("<a href=\"http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D{0}%26opc%3DLIC%26Frame%3DSI%26oi%3D{1}\">Ingresar</a>", impuestoLiquidar, objeto);
                        }                        
                        break;
                    }                    
            }

            if (this.Impuesto.SelectedIndex == 5)
            {
                datosObjeto = "<tr class='datos'>";
                datosObjeto += string.Format("<td class='gris'>{0}</td>", objetoFormateado);
                datosObjeto += string.Format("<td class='amarillo'>Cuota {0}</td>", cuotaNumero);
                //datosObjeto += string.Format("<td class='amarillo'>{0}</td>", montoCuota);
                datosObjeto += string.Format("<td class='gris'>{0}</td>", medioPago);
                datosObjeto += "</tr>";
            }
            else
            {
                if (ConAnual.Checked)
                {
                    datosObjeto = "<tr class='datos'>";
                    datosObjeto += string.Format("<td rowspan='2' class='gris'>{0}</td>", objetoFormateado);
                    datosObjeto += string.Format("<td class='amarillo'>Cuota {0}</td>", cuotaNumero);
                    datosObjeto += string.Format("<td class='amarillo'>{0}</td>", montoCuota);
                    datosObjeto += string.Format("<td rowspan='2' class='gris'>{0}</td>", medioPago);
                    datosObjeto += "</tr>";
                    datosObjeto += "<tr class='datos'><td class='blanco'>Anual</td>";
                    datosObjeto += string.Format("<td class='blanco'>{0}</td>", montoAnual);
                    datosObjeto += "</tr>";
                }
                else
                {
                    datosObjeto = "<tr class='datos'>";
                    datosObjeto += string.Format("<td class='gris'>{0}</td>", objetoFormateado);
                    datosObjeto += string.Format("<td class='amarillo'>Cuota {0}</td>", cuotaNumero);
                    datosObjeto += string.Format("<td class='amarillo'>{0}</td>", montoCuota);
                    datosObjeto += string.Format("<td class='gris'>{0}</td>", medioPago);
                    datosObjeto += "</tr>";
                }
            }

        }
    }
}

