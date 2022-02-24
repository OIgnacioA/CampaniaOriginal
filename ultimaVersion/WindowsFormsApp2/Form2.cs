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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace WindowsFormsApp2
{
    public partial class Form2 : Form
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
        string planta = string.Empty;
        string plantaDescri = string.Empty;
        string impuesto = string.Empty;
        string fechaOpcion = string.Empty;
        bool flagg2 = true; 
        string Desk = string.Empty;
        //string descuento = string.Empty;

        string nombreImpuesto = string.Empty;
        
        string datosObjeto = string.Empty;

        string impuestoLiquidar = string.Empty;

        string directorioOrigen = @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\";
        string directorioDestino = "\\arba.gov.ar\\DE\\GGTI\\Gerencia de Produccion\\Mantenimiento\\Boleta Electronica\\Destino\\";

        public Form2()
        {
         
            CheckForIllegalCrossThreadCalls = false;
 
            InitializeComponent();
            //Centrar
            this.StartPosition = FormStartPosition.CenterScreen;
            ModoOriginal.Checked = true; 
            LabelZip.Enabled = false;
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
                txtDestino = this.Origen.FileName;

                ///////////////////////////////////////

               if (ModoOriginal.Checked  == true) {


                    txtDestino = txtDestino.Replace(this.Origen.InitialDirectory, "");


               } else if (ModoNuevo.Checked == true){

                    txtDestino = ObtenerNombre(txtDestino);

               }

                Console.WriteLine("txtdestino   recortado-------------" + txtDestino);
                // txtDestino = ObtenerNombre(txtDestino);

                /////////////////////////////////////////////


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

            string ArgumentoreaderNuevo = string.Empty;
            string ArgumentoOpcionCheck1 = string.Empty;
            string ArgumentoOpcionCheck2 = string.Empty;

            string datosTodosObjetos = string.Empty;
            int cantidadCorte = Convert.ToInt32(this.txtCantidadCorte.Text);
            int cantidadArchivosGenerados = 1;


            string nombreArchivoGenerado = string.Format("{0}-Parte-{1}.csv", txtDestino, cantidadArchivosGenerados);

            fechaOpcion = this.FechaOpcion.Value.ToLongDateString().Replace(",", "");
            fechaVencimiento = fechaOpcion;



            ////////////////////////////////////////////////////
       



            if (ModoOriginal.Checked == true)
            {
                ArgumentoOpcionCheck1 = nombreArchivoGenerado;
                ArgumentoOpcionCheck2 = txtDestino + "-Informe.txt";
            }
            else if (ModoNuevo.Checked == true)
            {
                ArgumentoreaderNuevo = directorioDestino + "\\" + nombreArchivoGenerado;
                ArgumentoOpcionCheck1 = ArgumentoreaderNuevo;
                ArgumentoOpcionCheck2 = directorioDestino + txtDestino + "-Informe.txt";

            }

            /////////////////////////////////////////////////////
            ///


            StreamWriter sw = new System.IO.StreamWriter(ArgumentoOpcionCheck1);
            StreamWriter swResultados = new System.IO.StreamWriter(ArgumentoOpcionCheck2);



            //genero reporte.
            swResultados.Write("Se generearon los siguientes archivos:");
            swResultados.WriteLine();
            swResultados.Write(string.Format("Archivo ** {0} ** ", nombreArchivoGenerado));


            this.EscribirCabecera(sw);

            System.IO.StreamReader file = new System.IO.StreamReader(txtOrigen);



            Thread t1 = new Thread(InformarArchivosGenerados);


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
                    mailLinea += string.Format("{0}|Cuit: {1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", razonsocialAux, this.formatearCuit(cuitAux), fechaVencimiento, fechaOpcion, anio, cuota, impuesto, datosTodosObjetos, porcentaje);

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



                        ////////////////////////////////////////////////

                        if (ModoOriginal.Checked == true)
                        {
                            ArgumentoOpcionCheck1 = nombreArchivoGenerado;
                         
                        }
                        else if (ModoNuevo.Checked == true)
                        {
                            ArgumentoreaderNuevo = directorioDestino + "\\" + nombreArchivoGenerado;
                            ArgumentoOpcionCheck1 = ArgumentoreaderNuevo;
                           
                        }

                        sw = new System.IO.StreamWriter(ArgumentoOpcionCheck1); //Original : nombreArchivoGenerado.
                        swResultados.Write(string.Format("Archivo ** {0} **", nombreArchivoGenerado));                    


                        ////////////////////////////////////////////




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
            mailLinea += string.Format("{0}|Cuit: {1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", razonsocialAux, this.formatearCuit(cuitAux), fechaVencimiento, fechaOpcion, anio, cuota, impuesto, datosTodosObjetos, porcentaje);

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

            ///////////////////////////////////////////////////////
            ///
         
            if (ModoNuevo.Checked == true)
            {

                mensaje = string.Format("Se leyeron {0} suscripciones y se generaron {1} mails para enviar. Armar bases?", counter.ToString(), distintos.ToString());
                if (MessageBox.Show(mensaje, "Control Totales ATENCION", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    this.barraGenerados.Value = cantidadAleer;

                    t1.Start();

                }

            }
           
            /////////////////////////////////////////////////
        }

        private string formatearCuit(string pCuit)
        {
            string cuitFormateado = string.Empty;
            if (pCuit.Length == 11)
            {
                string primeraParte = pCuit.Substring(0, 2);
                string dni = pCuit.Substring(2, 8);
                string digito = pCuit.Substring(10, 1);
                cuitFormateado = string.Format("{0}-{1}-{2}", primeraParte, dni, digito);
            }           
            return cuitFormateado;
        }


        public string ObtenerPath(string nom)
        {   
            int cont = 0; 
            string name = "";
            string Navn = "";

            for ( int i = nom.Length - 1; i > 0; i--)
            {

                if (nom[i] == '\\')
                {
                    break;
                }

                name += nom[i];
                cont++;
            }


            int diff = nom.Length - cont; ;

            for (int j = 0; j < diff; j++)
            {

                Navn += nom[j];
            }


            return Navn;
        }


        public string ObtenerNombre(string nom)
        {
           
            string name = "";
            string Navn = "";
            

            for (int i = nom.Length - 1; i > 0; i--)
            {
                if (nom[i] == '\\')
                {
                    break; 
                }
                name += nom[i];
            }

            for (int j = name.Length - 1; j > -1; j--)
            {

                Navn += name[j];
            }


            return Navn;
        }

        public void Reloj()
        {

            BarraReloj.Maximum = 10;

            while (flagg2 == true)
            {

                for (int i = 0; i < 11; i++)
                {

                    Thread.Sleep(100);

                    BarraReloj.Value = i;

                }
            }
            BarraReloj.Value = 0;
        }

        private void InformarArchivosGenerados()
        {

            LabelZip.Enabled = true; 
            Thread t3 = new Thread(Reloj);

            string Desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Desk += "\\"+ txtDestino + ".zip";


            // Argumentos de CreateFromDirectory: (desde donde tomar los archivos, donde generar el Zip - con ".zip" al final.

            ModifyProgressBarColor.SetState(BarraReloj, 3);

            t3.Start();

            ZipFile.CreateFromDirectory(directorioDestino, Desk);

            
            flagg2 = false;
            BarraReloj.Value = 10;



            string mensaje = string.Format("Se generó el archivo .Zip: ' {0}.Zip ', con los datos para el envío de las campañas y el archivo Original. en un .Zip ubicado en el Escritorio de su PC",  txtDestino);
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

            Console.WriteLine("directorioDestino-------------" + directorioDestino);
            Console.WriteLine("directorioOrigen-------------" + directorioOrigen);
            Console.WriteLine("txtOrigen-------------" + txtOrigen);
        }        

        private void habilitarGenerar()
        {
            this.txtArchivoOrigen.Text = txtOrigen;            
            this.Generar.Enabled = (txtOrigen != string.Empty);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
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
                    directorioOrigen += @"Automotores\";
                    directorioDestino += @"Automotores\";
                    impuestoLiquidar = "1";
                    nombreImpuesto = "Automotor";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/Liquidacion.asp?imp=1&opc=LIC&oi=" + "{0}";
                        break;
                }
                case 1:
                {
                    directorioOrigen += @"Embarcaciones\";
                    directorioDestino += @"Embarcaciones\";
                    nombreImpuesto = "Embarcaciones";
                    impuestoLiquidar = "3";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/Liquidacion.asp?imp=3&opc=LIC&oi=" + "{0}";
                    break;
                    
                }
                case 2:
                {
                    directorioOrigen += @"Edificado\";
                    directorioDestino += @"Edificado\";
                    nombreImpuesto = "Edificado";
                    impuestoLiquidar = "0";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/Liquidacion.asp?imp=0&opc=LIC&oi=" + "{0}";
                    break;
                }
                case 3:
                {
                    directorioOrigen += @"Baldio\";
                    directorioDestino += @"Baldio\";
                    nombreImpuesto = "Baldio";
                    impuestoLiquidar = "0";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/Liquidacion.asp?imp=0&opc=LIC&oi=" + "{0}";
                    break;
                }
                case 4:
                {
                    directorioOrigen += @"Rural\";
                    directorioDestino += @"Rural\";
                    impuestoLiquidar = "0";
                    nombreImpuesto = "Rural";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/Liquidacion.asp?imp=0&opc=LIC&oi=" + "{0}";
                    break;
                }
                case 5:
                {
                    directorioOrigen += @"Complementario\";
                    directorioDestino += @"Complementario\";
                    nombreImpuesto = "Complementario";
                    impuestoLiquidar = "10";
                    txturl.Text = "https://www.arba.gov.ar/Aplicaciones/LiqPredet.asp?imp=10&Frame=NO&origen=WEB&opc=IIC";
                    break;
                }

                default:
                { 
                    cuotas = new List<string>() { "0" };
                    break;
                }
            }
            this.Origen.InitialDirectory = directorioOrigen;

            directorioDestino += "Destino2022";
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
                        porcentaje = "20";
                        anio = "2020";
                        cuota = "3";
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
                        //mail = line.Substring(0, 255).TrimEnd(' ').ToLower();
                        //objeto = line.Substring(255, 11).TrimEnd(' ');
                        //objetoFormateado = formatearCuit(objeto);
                        //razonsocial = line.Substring(275, 60).TrimEnd(' ');
                        //planta = line.Substring(347, 1).Trim(' ');
                        //debitoCredito = line.Substring(348, 1).Trim(' ');
                        //buenContribuyente = line.Substring(349, 1).Trim(' ');
                        //cuit = objeto;
                        this.LeerLineaNuevo(line);
                        break;
                    }

                default:

                    break;
            }
            TextInfo myTI = CultureInfo.CurrentCulture.TextInfo;
            razonsocial = myTI.ToTitleCase(razonsocial);

        }


        private void LeerLineaNuevo(string line)
        {            
            mail = line.Substring(0, 120).TrimEnd(' ').ToLower();
            objeto = line.Substring(120, 11).TrimEnd(' ');
            objetoFormateado = this.formatearObjeto(objeto);
            planta = line.Substring(131, 1).TrimEnd(' ');
            razonsocial = line.Substring(132, 60).TrimEnd(' ');
            porcentaje = line.Substring(192, 2).TrimEnd(' ');
            anio = line.Substring(194, 4).TrimEnd(' ');
            cuota = line.Substring(198, 2).TrimEnd(' ');
            fechaVencimiento = Convert.ToDateTime(line.Substring(200, 10).TrimEnd(' ')).ToLongDateString().Replace(",", "");
            fechaVencimientoNumero = line.Substring(200, 10).TrimEnd(' ');
            montoCuota = line.Substring(210, 17).Trim(' ');
            montoAnual = line.Substring(227, 17).Trim(' ');
            debitoCredito = line.Substring(244, 1).Trim(' ');
            cuit = line.Substring(245, 11).TrimEnd(' ');

            //Le pongo si es con anual o no.
            if (ConAnual.Checked)
            {
                cuota = cuota + " y Saldo Anual";
            }

            switch (planta)
            {
                case "B":
                    {
                        plantaDescri = "Baldio";
                        break;
                    }
                case "E":
                    {
                        plantaDescri = "Edificado";
                        break;
                    }
                case "R":
                    {
                        plantaDescri = "Rural";
                        break;
                    }
                default:
                    break;
            }
            

            TextInfo myTI = CultureInfo.CurrentCulture.TextInfo;
            razonsocial = myTI.ToTitleCase(razonsocial);
            
        }

        private string formatearObjeto(string pObjeto)
        {
            string resultado;
            switch (this.Impuesto.SelectedIndex)
            {
                case 0:
                    {
                        //Es auto
                        resultado = pObjeto;
                        break;
                    }
                case 1:
                    {
                        //Es Emba
                        resultado = pObjeto;
                        break;

                    }
                case 2:
                    {
                        //Es Edificado
                        resultado = formatearObjetoInmobiliario(pObjeto);                       
                        break;
                    }
                case 3:
                    {
                        //Es Baldio
                        resultado = formatearObjetoInmobiliario(pObjeto);
                        break;
                    }
                case 4:
                    {
                        //Es Rural
                        resultado = formatearObjetoInmobiliario(pObjeto);
                        break;
                    }
                case 5:
                    {
                        //Es Complementario
                        resultado = formatearCuit(pObjeto);
                        break;
                    }

                default:
                    {
                        resultado = pObjeto;
                        break;
                    }
            }
            return resultado;
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
                case "D":
                    {
                        medioPago = "Débito en Cuenta";                        
                        break;
                    }
                case "2":
                    {
                        medioPago = "Tarjeta de Crédito";                        
                        break;
                    }
                case "0":
                case "C":
                    {
                         medioPago = string.Format("<a href=\"" + txturl.Text + "\">Ingresar</a>", objeto); 
                        /*
                        if (this.Impuesto.SelectedIndex == 5)
                        {
                            //medioPago = "<a href=\"https://sso.arba.gov.ar/Login/login?service=http%3A%2F%2Fwww4.arba.gov.ar%2FLiqPredet%2Fsso%2FInicioLiquidacionIIC.do%3FFrame%3DNO%26origen%3DWEB%26imp%3D10%26cuit%3D\">Ingresar</a>";
                            medioPago =  "<a href=\"https://www.arba.gov.ar/aplicaciones/LiqPredet.asp?imp=10&Fame=NO&origen=WEB&op=IIC\">Ingresar</a>";
                        }
                        else
                        {
                            medioPago = string.Format("<a href=\"http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D{0}%26opc%3DLIC%26Frame%3DSI%26oi%3D{1}\">Ingresar</a>", impuestoLiquidar, objeto);
                        }      
                        */
                        break;
                    }
             /*   case "C":
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
                    }*/
                default:
                    {
                        // medioPago = "ERROR";
                        // break;
                        medioPago = string.Format("<a href=\"" + txturl.Text + "\">Ingresar</a>", objeto);
                        break;
                    }
            }

            if (this.Impuesto.SelectedIndex == 5)
            {
                datosObjeto = "<tr class='datos'>";
                datosObjeto += string.Format("<td class='gris'>{0} - {1}</td>", objetoFormateado, plantaDescri);
                datosObjeto += string.Format("<td class='amarillo'>Cuota {0}</td>", cuotaNumero);
                datosObjeto += string.Format("<td class='amarillo'>{0}</td>", montoCuota);
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

    public static class ModifyProgressBarColor
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)] static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr w, IntPtr l);

        public static void SetState(this ProgressBar pBar, int state) { SendMessage(pBar.Handle, 1040, (IntPtr)state, IntPtr.Zero); }
    }
}

