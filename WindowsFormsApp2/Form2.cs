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
    public partial class Form2 : Form
    {
        //--Variables de campo. 
        #region

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
        //string descuento = string.Empty;

        string nombreImpuesto = string.Empty;
        
        string datosObjeto = string.Empty; // OBJETO DONDE SE DEPOCITA EL RESULTADO FINAL
        /// </summary>

        string impuestoLiquidar = string.Empty;

        string directorioOrigen = @"C:\Users\oscar.avendano\Desktop\aplicacion Campaña\Archivos de Prueba"; 
                                   // @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\";
        string directorioDestino = @"C:\Users\oscar.avendano\Desktop\aplicacion Campaña\Archivos de Prueba\sehent"; // @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Destino\";

        #endregion //--Variables de campo. 

        public Form2()
        {
            InitializeComponent();
            habilitarGenerar();
        }

        // --- configuracion Datos Base: 


        private void button2_Click(object sender, EventArgs e) // Boton Origen.
        {

            //Ñ 11
            #region

            //Console.WriteLine("ver el path " + Path.GetDirectoryName(this.Origen.InitialDirectory));

            /* --- El valor de busqueda se encuentra en:  "Form2.Designer.cs"
             
                     DialogResult se puede usar desde el código que 
                     mostró un cuadro de diálogo para determinar si 
                     un usuario aceptó ( true ) o canceló
                     ( false ) el cuadro de diálogo.
            

            Ñ 11 -IMPORTANTE: 

                La ruta para el Showdialog estaba mal escrita,
            asi que no llevaba a ningun lado,Mientras esto permaneció así 
            siempre ABRIA LA MSIMA VENTANA, aparentemente por defecto quizá 
            la ultima que funcionó en la busqueda. 
            
            */
            #endregion

            Origen.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";// Ñ:)

            DialogResult dr = this.Origen.ShowDialog();

            txtOrigen = this.Origen.FileName; // TxtOrigen  toma el valor del .txt elegido en la ventana emergente: 

            this.habilitarGenerar(); //mostrar Path por pantalla y habilitar boton "generar"
        }

        private void habilitarGenerar()
        {
            this.txtArchivoOrigen.Text = txtOrigen;  // vuelca le valor de txtorigen (el path del txt que se clickea en la ventana                                             emergente (Boton Origen) y lo muestra por pantalla en el cuadro de texto.       
            this.Generar.Enabled = (txtOrigen != string.Empty); //( impuesto esta bueno: boton true, si txt es true ;)  )
        }

        private void button1_Click(object sender, EventArgs e) //Ñ4 boton Generar. 
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
                txtDestino = this.Origen.FileName; //Ñ 3 Objeto del formulario -button.text("Origen") (tipo OpenFileDialog- ventana emergente -) : da x ej:  \\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\Edificado\20150519-3-CO.TXT

      
                txtDestino = txtDestino.Replace(this.Origen.InitialDirectory, ""); //Ñ 5:  quedaría : "\Edificado\20150519-3-CO.TXT" 

             

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
            int cantidadAleer = Convert.ToInt32(this.txtCantidad.Text); // "Cant. Suscripciones"

            //barras de carga-------------------------------- :3
            this.barraGenerados.Maximum = cantidadAleer;
                 this.barraLeidos.Maximum = cantidadAleer;
            //-----------------------------------------------


            int counter = 0;
            int contador = 0;
            int distintos = 0;
            int escritos = 0;
            int cantidadMailIgual = 0;
            string line;
            string mailLinea = string.Empty;
            string ultimoMail = string.Empty;
             
            string datosTodosObjetos = string.Empty;
            int cantidadCorte = Convert.ToInt32(this.txtCantidadCorte.Text); //Ñ7 -  "150000"
            int cantidadArchivosGenerados = 1; //incrementa en linea 194


            string nombreArchivoGenerado = string.Format("{0}-Parte-{1}.csv", txtDestino, cantidadArchivosGenerados); //Ñ txtDestino a estas alturas vale:(por ejemplo) "\Edificado\20150519-3-CO.TXT" Viene desde : " private void button1_Click" (viene de la ventana emergente) 

            fechaOpcion = this.FechaOpcion.Value.ToLongDateString().Replace(",", ""); //Ñ1: pase de valor de fecha en formulario a codigo. "FechaOpcion" es el objeto. 
            fechaVencimiento = fechaOpcion;// ¿porqué no lo pasa directamente a -fechaVencimiento- ?


            // StreamWriter!!! ----------------------------

            StreamWriter sw = new System.IO.StreamWriter(nombreArchivoGenerado); //Ñ2 nombreArchivogenerado es - para estas alturas - un path!!    
            StreamWriter swResultados = new System.IO.StreamWriter(txtDestino + "-Informe.txt");
            
            //genero reporte.
            swResultados.Write("Se generearon los siguientes archivos:");
            swResultados.WriteLine();
            swResultados.Write(string.Format("Archivo ** {0} ** ", nombreArchivoGenerado)); //Ñ 9 y "Con 11401 suscripciones y 5983 mails para enviar" donde está?


            this.EscribirCabecera(sw); //Ñ 10 va al check box a ver si agrega la cabecerao no:  "email|nombre|cuit|fechav|fechao|anio|cuota|impuesto|datos|descuento""

          
            System.IO.StreamReader file = new System.IO.StreamReader(txtOrigen); //el StreamReader toma por argumento la ubicacion de un txt. Un path. (TxtOrigen tomo ese valor en: el boton Origen) : ""\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\Automotores\20150422-3-dt.TXT"

            /*
            el objeto "file", de tipo streamReader: toma "txtOrigen" como argumento: 
            por esto en el metodo ReadLine devuleve linea a linea
            el contenido del TxT (aunque solo muestra el mail en la impreison por pantalla) 
             
             */


            while ((line = file.ReadLine()) != null) //  
            {
                

                this.LeerLinea(line); // acá se envia el contenido de "line" a un método donde se trabaja con el combobox de "tipo de impuesto" 
                                      // alli con el método "TrimEnd(' ')"    se seleccionan diferentes partes de la linea del Txt. 

              

                //veo si es la primer linea.
                if (mailAux == string.Empty) //(mailAux) String declarado en primeras Lineas. [Hasta acá siempre lo recibe vacio la primera vez]
                {
                    mailAux = mail; //(mail) String declarado en primeras Lineas; vuelve con valor desde "LeerLinea(Line)".
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
            mailLinea += string.Format("{0}|Cuit: {1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}", this.formatearCuit(cuitAux), objetoFormateado, fechaVencimiento, fechaOpcion, anio, cuota, impuesto, datosTodosObjetos, porcentaje);
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
                cuitFormateado = string.Format("{0}-{1}-{2}", primeraParte, dni, digito);
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

    

        private void Impuesto_SelectedIndexChanged(object sender, EventArgs e)
        {

            Console.WriteLine("acá es 1) " + directorioOrigen);

            directorioOrigen = @"C:\Users\oscar.avendano\Desktop\aplicacion Campaña\Archivos de Prueba\sehent\"; // Ñ 11 paarentemente si se escribe mal la direccion de origen, y esta no daa ningun lado lleva a la ultima direccion que funcionó en la bsuqueda.

                              // Original  -> @"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Origen\";

            directorioDestino = @"C:\Users\oscar.avendano\Desktop\aplicacion Campaña\Archivos de Prueba\sehent\";

                               // Original  ->@"\\arba.gov.ar\DE\GGTI\Gerencia de Produccion\Mantenimiento\Boleta Electronica\Destino\";

            List<string> cuotas;

            

            switch (this.Impuesto.SelectedIndex)
            {
                case 0:
                {
                        directorioOrigen += @"Automotores\";
                        directorioDestino += @"Automotores\";

                        Console.WriteLine("acá es 2) " + directorioOrigen);

                        impuestoLiquidar = "1";
                    nombreImpuesto = "Automotor";
                    txturl.Text = "http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D1%26opc%3DLIC%26Frame%3DSI%26oi%3D" + "{0}";
                                  

                        break;
                }
                case 1:
                {
                    directorioOrigen += @"Automotores\"; //@"Embarcaciones\";
                    directorioDestino += @"Automotores\"; //@"Embarcaciones\";
                    nombreImpuesto = "Embarcaciones";
                    impuestoLiquidar = "3";
                    txturl.Text = "http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D3%26opc%3DLIC%26Frame%3DSI%26oi%3D" + "{0}";
                    break;
                    
                }
                case 2:
                {
                    directorioOrigen += @"Automotores\"; //@"Edificado\";
                    directorioDestino += @"Automotores\"; //@"Edificado\";
                        nombreImpuesto = "Edificado";
                    impuestoLiquidar = "0";
                    txturl.Text = "http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D0%26opc%3DLIC%26Frame%3DSI%26oi%3D" + "{0}";
                    break;
                }
                case 3:
                {
                    directorioOrigen += @"Automotores\"; //@"Baldio\";
                        directorioDestino += @"Automotores\"; //@"Baldio\";
                        nombreImpuesto = "Baldio";
                    impuestoLiquidar = "0";
                    txturl.Text = "http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D0%26opc%3DLIC%26Frame%3DSI%26oi%3D" + "{0}";
                    break;
                }
                case 4:
                {
                    directorioOrigen += @"Automotores\"; //@"Rural\";
                        directorioDestino += @"Automotores\"; //@"Rural\";
                        impuestoLiquidar = "0";
                    nombreImpuesto = "Rural";
                    txturl.Text = "http://www.arba.gov.ar/AplicacionesFrame.asp?url=Aplicaciones%2FLiquidacion%2Easp%3Fimp%3D0%26opc%3DLIC%26Frame%3DSI%26oi%3D" + "{0}";
                    break;
                }
                case 5:
                {
                    directorioOrigen += @"Automotores\"; //@"Complementario\";
                    directorioDestino += @"Automotores\"; //@"Complementario\";
                    nombreImpuesto = "Complementario";
                    impuestoLiquidar = "10";
                    txturl.Text = "https://www.arba.gov.ar/aplicaciones/LiqPredet.asp?imp=10&Fame=NO&origen=WEB&op=IIC";
                    break;
                }

                default:
                { 
                    cuotas = new List<string>() { "0" };
                    break;
                }
            }

            Console.WriteLine("acá es 3) " + directorioOrigen);

            this.Origen.InitialDirectory = directorioOrigen;  // IMPORTANTE : Origen es de tipo openFileDialog. Asi q posee el método InitialDirectory, que carga la busqueda principal de la ventana emergente del boton "origen"  abriendo una ventana diferente segun cual sea el impuesto elegido.             
        }
       
        
        private void LeerLinea(string line)
        {
            switch (this.Impuesto.SelectedIndex)
            {
                case 0:
                case 1:
                    {

                        /*The String.TrimEnd method removes characters 
                         * from the end of a string, 
                         * creating a new string object
                        ------
                        "String.Substring"  toma el primer parametro 
                        y parte desde ese caracter, 
                        por distancia igual al segundo 
                        parametro y lo devuleve. 

                        ---------

                        el archivo de texto tiene las columnas separadas
                        por las distancias marcadas en el TrimEnd.

                        */

                        mail = line.Substring(0, 255).TrimEnd(' ').ToLower();
                             Console.WriteLine("------------->" + mail);
                        objeto = line.Substring(255, 11).TrimEnd(' ');
                             Console.WriteLine("------------->" + objeto);
                        objetoFormateado = objeto.ToUpper();
                             Console.WriteLine("------------->" + objetoFormateado);
                        razonsocial = line.Substring(266, 60).TrimEnd(' ');
                             Console.WriteLine("------------->" + razonsocial);
                        porcentaje = string.Empty;
                        fechaVencimiento = Convert.ToDateTime(line.Substring(334, 10).TrimEnd(' ')).ToLongDateString().Replace(",", "");
                             Console.WriteLine("------------->" + fechaVencimiento);
                        fechaVencimientoNumero = line.Substring(334, 10).TrimEnd(' ');
                             Console.WriteLine("------------->" + fechaVencimientoNumero);
                        montoCuota = line.Substring(345, 17).Trim(' ');
                             Console.WriteLine("------------->" + montoCuota);
                        montoAnual = line.Substring(362, 16).Trim(' ');
                             Console.WriteLine("------------->" + montoAnual);


                        //codigoElectronico = line.Substring(378, 14).Trim(' ');
                       /* debitoCredito = line.Substring(392, 1).Trim(' ');
                            Console.WriteLine("------------->" + debitoCredito);
                        buenContribuyente = line.Substring(393, 1).Trim(' ');
                            Console.WriteLine("------------->" + buenContribuyente);
                        cuit = line.Substring(394, 11).TrimEnd(' ');
                            Console.WriteLine("------------->" + cuit);

                        porcentaje = "20";
                        anio = "2020";
                        cuota = "3";*/

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


        //---------------- x ver. 

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

            //MAIN!!!     (DONDE ESTAN LSO TXT QUE TIENEN DISCRIMINADO EL DEBITO O CREDITO??)

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

        private void EscribirCabecera(StreamWriter pSw)
        {
            if (this.ConCabecera.Checked)
            {
                pSw.Write("email|nombre|cuit|fechav|fechao|anio|cuota|impuesto|datos|descuento");
                pSw.WriteLine();
            }

        }

    }
}

