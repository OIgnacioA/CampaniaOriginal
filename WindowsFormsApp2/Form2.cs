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
        string fullPath;
        //string descuento = string.Empty;
        string CarpetaDestino = string.Empty;
        string nombreImpuesto = string.Empty;      
        string datosObjeto = string.Empty;
        string impuestoLiquidar = string.Empty;
        string nombreDeZip = string.Empty;
        int counter = 1;
        int conterror = 0;

        int contzip = 0;
        




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
                fullPath = txtOrigen; 

               //////////////////////////////////////////////////////////////////////////


               if ((ModoOriginal.Checked  == true) || (ModoNuevo.Checked == true))  {

                    txtDestino = ObtenerNombre(txtOrigen);

               } 

               //////////////////////////////////////////////////////////////////////////

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


            CarpetaDestino = ObtenerPath(fullPath);

            if (ModoOriginal.Checked == true)
            {
                ArgumentoOpcionCheck1 = CarpetaDestino + "\\" + nombreArchivoGenerado;
                ArgumentoOpcionCheck2 = fullPath + "-Informe.txt";
            }
            else if (ModoNuevo.Checked == true)
            {
                ArgumentoreaderNuevo = directorioDestino + "\\" + nombreArchivoGenerado;
                ArgumentoOpcionCheck1 = ArgumentoreaderNuevo;
                ArgumentoOpcionCheck2 = directorioDestino + "\\" + txtDestino + "-Informe.txt";

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
                            ArgumentoOpcionCheck1 = CarpetaDestino + "\\" + nombreArchivoGenerado;
                         
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

                mensaje = string.Format("La cantidad de suscripciones configuradas ({0}) y es distinta a la cantidad de registros leidos ({1}). De todas maneras se generaron {2} mails para enviar. ¿Armar bases?", cantidadAleer, counter, distintos);

                if (MessageBox.Show(mensaje, "Cantidad de registros ERRONEA!!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    this.barraGenerados.Value = cantidadAleer;

                    t1.Start(); //InformarArchivosGenerados()

                }

            }
            else
            {
                mensaje = string.Format("Se leyeron {0} suscripciones y se generaron {1} mails para enviar.   ¿Armar bases?", counter.ToString(), distintos.ToString());


                if (MessageBox.Show(mensaje, "Control Totales ATENCION", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    this.barraGenerados.Value = cantidadAleer;

                    t1.Start(); //InformarArchivosGenerados()

                }
            }

            Console.WriteLine("cantidad de errores: " + conterror);

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
            int cont2 = 0;
            string Navn = "";
            bool go = false; 

            for (int i = nom.Length - 1; i > 0; i--)
            {

                if (nom[i] == '\\')
                {
                    go = true; 
                    break;
                }

            }


            if (go == true)
            {

                for (int i = nom.Length - 1; i > 0; i--)
                {
                    if (nom[i] == '\\')
                    {
                        cont2++; 
                    }

                    cont++;

                    if (cont2 == 1) {
                        break;
                    }
                }

                int diff = nom.Length - cont; ;

                for (int j = 0; j < diff; j++)
                {
                    Navn += nom[j];
                }

                    return Navn;

            }else { return nom; }

        }


        public string ObtenerNombre(string nom)
        {
           
            string name = "";
            string Navn = "";
            

            for (int i = nom.Length - 1; i > -1; i--)
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

           /* string Desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Desk += "\\"+ txtDestino + ".zip";*/

            ModifyProgressBarColor.SetState(BarraReloj, 3);


            //////////////////////////////////////////////////////////////////////


            if (ModoOriginal.Checked == true)
            {
                t3.Start();

                InformarArchivosGenerados_Original();

            }
            else if (ModoNuevo.Checked == true)
            {
                t3.Start();
             
                DirectoryInfo di = new DirectoryInfo(directorioDestino); //(".\\");
                FileInfo[] archivos = di.GetFiles(txtDestino + "*");


                ///////// comprobar que no exista ya un zip: cambiar nombre:

                         string zip = VerificarZipPrevio(directorioDestino, archivos);

                //////// string.Format("{0}\\{1}.zip", directorioDestino, txtDestino); // --aqui se establece la direccion destino del zip 


                StreamReader r;

                using (FileStream zipToOpen = new FileStream(zip, FileMode.Create))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        foreach (FileInfo fileToCompress in archivos)
                        {
                            string esZip = fileToCompress.Name;

                            if (esZip.Contains(".zip")) { /* XD */ }
                            else
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
                }

                foreach (FileInfo fileToCompress in archivos)
                {
                    string esZip = fileToCompress.Name;

                    if (esZip.Contains(".zip")) { }
                    else
                    {
                        File.Delete(fileToCompress.FullName);
                    }
                }


              flagg2 = false;
              BarraReloj.Value = 10;

              string mensaje2 = string.Format("Se generó el archivo .Zip: ' {0} ', con los datos para el envío de las campañas y el archivo Original.ubicación en: {1} ", nombreDeZip, directorioDestino);
              MessageBox.Show(mensaje2);


        }

              ///////////////////////////////////////////////////////////////////

       
    }


    private void InformarArchivosGenerados_Original()
        {

            string txtDestin = "";

            FileInfo fileinfo = new FileInfo(fullPath);
            long peso = fileinfo.Length; 
            

            DirectoryInfo di = new DirectoryInfo(CarpetaDestino); //(".\\");

            txtDestin = ObtenerNombre(txtDestino);

            FileInfo[] archivos = di.GetFiles(txtDestin + "*"); //modelo de busqueda: asi se seleccionan archivos con cierto nombre. 
           

            /////// comprobar que no exista ya un zip: cambiar nombre:

                   string zip = VerificarZipPrevio(CarpetaDestino, archivos);

            /////// string zip = string.Format("{0}\\{1}.zip", CarpetaDestino, txtDestino); // --aqui se establece la direccion destino del zip 



            StreamReader r;

            using (FileStream zipToOpen = new FileStream(zip, FileMode.Create)) 
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))  
                {
                    foreach (FileInfo fileToCompress in archivos)
                    {
                        
                        if (fileToCompress.Length == peso) { /* XD */ }
                        else {

                             string esZip = fileToCompress.Name;

                            if (esZip.Contains(".zip")) { /* XD */ }
                            else
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
                }
            }

           foreach (FileInfo fileToCompress in archivos)
           {
                string esZip = fileToCompress.Name; 

                if (fileToCompress.Length == peso) { /* XD */ }
                else
                {
                    if (esZip.Contains(".zip")) { }
                    else
                    {
                        File.Delete(fileToCompress.FullName);
                    }
                }

           }

            flagg2 = false;
            BarraReloj.Value = 10;

            string mensaje = string.Format("Se generó el archivo ' {0} ' con los datos para el envío de las campañas. Colocar dicho archivo en {1} y avisar a Mesa de ayuda.", nombreDeZip, directorioDestino);
            MessageBox.Show(mensaje);

     }

    public string VerificarZipPrevio(string Fuente, FileInfo[] ficheros)
    {

        int tempN = 0;
        string tempS = "";
        string tempS2 = "";
        string zip = "";
        int cont2 = 1;
        int contParentesis = 0;



            DirectoryInfo dire = new DirectoryInfo(Fuente);
            FileInfo[] files = dire.GetFiles("*.zip");
            string NombreArchivo = "";

            foreach (FileInfo file in files) {

                NombreArchivo = file.Name; 


                if (NombreArchivo.Contains("_"))
                { //se busca el archivo con nombre de zip.


                    for (int l = NombreArchivo.Length; l > -1; l--)
                    {

                        if (NombreArchivo[l - 1] == ')')
                        { //Busca el parentesis------charAt mide contando el cero. Por eso se resta uno.

                            int temp2 = l;

                            ///Medir distancia entre parentesis
                            while (NombreArchivo[temp2 - 2] != '(')
                            {
                                contParentesis++;
                                temp2--;
                            }

                            temp2 = l;

                            /// Tomar elementos ENTRE los parentesis
                            while (cont2 <= contParentesis)
                            {

                                tempS += NombreArchivo[temp2 - 2];

                                temp2--;
                                cont2++;
                            }

                            /// Voltear numero dentro de los parentesis
                            if (contParentesis > 1)
                            {

                                for (int n = tempS.Length - 1; n > -1; n--)
                                {

                                    tempS2 += tempS[n];
                                }
                            }
                            else { tempS2 = tempS; }

                            tempN = Int32.Parse(tempS2);
                            tempS2 = "";
                            tempS = "";
                            temp2 = 0;
                            cont2 = 1;
                            contParentesis = 0;

                            ///Seleccion del número mas grande entre los zips.

                            if (contzip < tempN) { contzip = tempN; }
                            //System.out.println("--------------->: " + ficheros[k].getName().charAt(l - 2));
                            break;

                        }
                    }
                }
            }
        

        //////////////////////////////////////////////////////////////////////////////////
        contzip++;

        if (ModoOriginal.Checked == true)
        {

            nombreDeZip = txtDestino + "_(" + contzip + ").zip";
            zip = CarpetaDestino + "\\" + nombreDeZip;

        }
        else if (ModoNuevo.Checked == true)
        {

            nombreDeZip = txtDestino + "_(" + contzip + ").zip";
            zip = directorioDestino + "\\" + nombreDeZip;

        }
        ////////////////////////////////////////////////////////////////////////////////
        contzip = 0;
        return zip;
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

            ModoOriginal.Enabled = false;
            ModoNuevo.Enabled = false;

            /*  Console.WriteLine("directorioDestino-------------" + directorioDestino);
              Console.WriteLine("directorioOrigen-------------" + directorioOrigen);
              Console.WriteLine("txtOrigen-------------" + txtOrigen);*/
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


                        try
                        {


                            cuit = line.Substring(394, 11).TrimEnd(' ');



                        }
                        catch (Exception e) {

                            cuit = string.Empty; 
                             Console.WriteLine("------------> Error de tipo ( " + e.Message + " ) En la      Línea: " + counter.ToString());

                            conterror++;
                            
                        
                        }

                        Console.WriteLine("Cuit: " + cuit + ": en linea: " + counter.ToString());


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

                        try
                        {

                            cuit = line.Substring(394, 11).TrimEnd(' ');

                        }
                        catch (Exception e)
                        {

                            cuit = string.Empty; 
                            Console.WriteLine("------------> Error de tipo ( " + e.Message + " ) En la      Línea: " + counter.ToString());

                            conterror++; // conteo de erroes de lectura

                        }

                        Console.WriteLine("Cuit: " + cuit + ": en linea: " + counter.ToString());

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


            try
            {

                cuit = line.Substring(245, 11).TrimEnd(' ');

            }
            catch (Exception e)
            {

                cuit = string.Empty; 
                Console.WriteLine("------------> Error de tipo ( " + e.Message + " ) En la      Línea: " + counter.ToString());

                conterror++;
            }

            Console.WriteLine("Cuit: " + cuit + ": en linea: " + counter.ToString());




            //Le pongo si es con anual o no.
            if (ConAnual.Checked)
            {
                cuota = cuota + " y Saldo Anual";

                Console.WriteLine("------>",cuota);
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

