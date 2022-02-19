using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace AuxContabilidad
{
    public partial class Form1 : Form
    {
        string strRutaExcelOrig = "", strRutaExcel = "", strRutaFacturas = "", strRutaFacturasResultado = "", strRutaActualSub = "";
        string strFactura = "", strPoliza = "", strPolizaBuscar = "", strGrupo = "",strGrupoAux="",strRutaAux="", strRutaAuxPoliza = "", strRutaArchivo="",strNombreExcel="",strMes="",strEncontrado="";
        int intCantidadRegistros = 0, intColFactura=0,intColPoliza=0,intColGrupo=0,intColMes=0,intColEncontrado=0,intConteo=0;
        
        string nombre, ip, user, password;
        string[] valoresConexion;
        string[] valoresIP;
        string[] valoresUsuarios;
        string[] valoresPassword;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var item in listaDocumentos)
            {
                MessageBox.Show(item);
            }
        }

        bool bValidacion = true;

        private void btnConectar_Click(object sender, EventArgs e)
        {
            conexion();
        }

        DateTime dtFechaArchivoYYYYMM;

        private void cbxNombre_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtIP.Text = valoresIP[cbxNombre.SelectedIndex];
            txtUser.Text = valoresUsuarios[cbxNombre.SelectedIndex];
            txtPass.Text = valoresPassword[cbxNombre.SelectedIndex];
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //limpiarCredenciales();
            KillAllTasksByName("Excel");

        }
        private void KillAllTasksByName(string taskName)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Si le pide ayuda a Goku cerrara todos los procesos de excel abiertos ya que necesita esa energia para" +
                    " crear una genkidama. Le recomendamos guardar su trabajo antes de seguir, desea continuar?", "Mensaje de advertencia MORTAL", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (Process proceso in Process.GetProcessesByName(taskName))
                    {
                        proceso.Kill();
                    }
                }
                else if (dialogResult == DialogResult.No)
                {
                    MessageBox.Show("Hmmmmm");
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }
        void limpiarCredenciales()
        {
            Properties.Settings.Default.NOMBRE = String.Empty;
            Properties.Settings.Default.IP = String.Empty;
            Properties.Settings.Default.USER = String.Empty;
            Properties.Settings.Default.PASSWORD = String.Empty;
            Properties.Settings.Default.Save();
        }

        bool existConnection()
        {
            bool exist = false;
            for (int i = 0; i < valoresConexion.Length; i++)
            {
                if (valoresConexion[i].ToString() == cbxNombre.Text)
                {
                    exist = true;
                    i = valoresConexion.Length;
                }
                else
                {
                    exist = false;
                }
            }
            return exist;
        }

        DirectoryInfo diDir,dirDirFacturas;
        Task tTarea;
        FileInfo[] fiArchivos;
        WebClient webclient = new WebClient();

        string [] listaDocumentos;
        public Form1()
        {
            InitializeComponent();
            nombre = Properties.Settings.Default.NOMBRE;
            ip = Properties.Settings.Default.IP;
            user = Properties.Settings.Default.USER;
            password = Properties.Settings.Default.PASSWORD;
            readConnections();
        }
        void AumentarProgreso(int intPorcentaje)
        {
            Invoke(new Action(() => pbProgreso.Value = intPorcentaje));
            Invoke(new Action(() => lblPorcentaje.Text = intPorcentaje.ToString() + " %"));
        }

        void readConnections()
        {
            try
            {
                valoresConexion = nombre.Split('|');
                valoresIP = ip.Split('|');
                valoresUsuarios = user.Split('|');
                valoresPassword = password.Split('|');

                for (int i = 0; i < valoresConexion.Length; i++)
                {
                    cbxNombre.Items.Add(valoresConexion[i].ToString());
                }

                cbxNombre.SelectedIndex = 0;

            }
            catch (Exception)
            {

            }
        }

        private void btnFindExcel_Click(object sender, EventArgs e)
        {
            ofdFindExcel.Filter = "Excel Files|*.xlsx";
            if (ofdFindExcel.ShowDialog() == DialogResult.OK)
            {
                txtRutaExcel.Text = ofdFindExcel.FileName;
                strNombreExcel = ofdFindExcel.SafeFileName;
            }
        }

        public static String[] FTPListTree(String FtpUri, String User, String Pass)
        {
            try
            {

                List<String> files = new List<String>();
                Queue<String> folders = new Queue<String>();
                folders.Enqueue(FtpUri);

                while (folders.Count > 0)
                {
                    String fld = folders.Dequeue();
                    List<String> newFiles = new List<String>();

                    FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(fld);
                    ftp.Credentials = new NetworkCredential(User, Pass);
                    ftp.UsePassive = false;
                    ftp.Method = WebRequestMethods.Ftp.ListDirectory;
                    using (StreamReader resp = new StreamReader(ftp.GetResponse().GetResponseStream()))
                    {
                        String line = resp.ReadLine();
                        while (line != null)
                        {
                            newFiles.Add(line.Trim());
                            line = resp.ReadLine();
                        }
                    }

                    ftp = (FtpWebRequest)FtpWebRequest.Create(fld);
                    ftp.Credentials = new NetworkCredential(User, Pass);
                    ftp.UsePassive = false;
                    ftp.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                    using (StreamReader resp = new StreamReader(ftp.GetResponse().GetResponseStream()))
                    {
                        String line = resp.ReadLine();
                        while (line != null)
                        {
                            if (line.Trim().ToLower().StartsWith("d") || line.Contains(" <DIR> "))
                            {
                                String dir = newFiles.First(x => line.EndsWith(x));
                                newFiles.Remove(dir);
                                folders.Enqueue(fld + dir + "/");
                            }
                            line = resp.ReadLine();
                        }
                    }
                    files.AddRange(from f in newFiles select fld + f);
                }
                return files.ToArray();

            }
            catch (Exception)
            {
                MessageBox.Show("Credenciales no validas");
                return null;
            }
        }

        void conexion()
        {
            try
            {
                string url = "ftp://" + txtIP.Text;
                listaDocumentos = FTPListTree(url + "/", txtUser.Text, txtPass.Text);
                MessageBox.Show("Conectado exitosamente");
                if (!existConnection())
                {
                    createCredentials();
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnFindFolder_Click(object sender, EventArgs e)
        {
            if (fbdFindFacturas.ShowDialog() == DialogResult.OK)
            {
                txtRutaFacturas.Text = fbdFindFacturas.SelectedPath;
            }
        }

        void createCredentials()
        {
            if (Properties.Settings.Default.NOMBRE == String.Empty)
            {
                Properties.Settings.Default.NOMBRE = cbxNombre.Text;
            }
            else
            {
                Properties.Settings.Default.NOMBRE = Properties.Settings.Default.NOMBRE + "|" + cbxNombre.Text;
            }

            if (Properties.Settings.Default.IP == String.Empty)
            {
                Properties.Settings.Default.IP = txtIP.Text;
            }
            else
            {
                Properties.Settings.Default.IP = Properties.Settings.Default.IP + "|" + txtIP.Text;
            }

            if (Properties.Settings.Default.USER == String.Empty)
            {
                Properties.Settings.Default.USER = txtUser.Text;
            }
            else
            {
                Properties.Settings.Default.USER = Properties.Settings.Default.USER + "|" + txtUser.Text;
            }

            if (Properties.Settings.Default.PASSWORD == String.Empty)
            {
                Properties.Settings.Default.PASSWORD = txtPass.Text;
            }
            else
            {
                Properties.Settings.Default.PASSWORD = Properties.Settings.Default.PASSWORD + "|" + txtPass.Text;
            }

            Properties.Settings.Default.Save();
        }

        private async void btnProceder_Click(object sender, EventArgs e)
        {
            if (Validar())
            {
                strRutaExcelOrig = txtRutaExcel.Text;
                strRutaFacturas = txtRutaFacturas.Text;
                KillAllTasksByName("Excel");
                if (CrearRutaResultado())
                {
                    if (CopiarArchivoExcel())
                    {
                        tTarea = new Task(() => Procesar());
                        tTarea.Start();
                        await tTarea;

                        Process.Start(new ProcessStartInfo("explorer.exe", "/select, \"" + strRutaActualSub));
                        MessageBox.Show("Termino");
                    }
                }
            }
            else
            {
                MessageBox.Show("Verifique las rutas");
            }
        }

        bool Validar()
        {
            bValidacion = true;
            if (txtRutaExcel.Text.CompareTo("") == 0 || txtRutaFacturas.Text.CompareTo("") == 0)
            {
                bValidacion = false;
            }
            return bValidacion;
        }

        bool CrearRutaResultado()
        {
            bValidacion = true;
            try
            {
                strRutaFacturasResultado = strRutaFacturas + "\\RESULTADO";
                diDir = new DirectoryInfo(strRutaFacturasResultado);

                if (!diDir.Exists) //Si el directorio no existe, lo creo
                {
                    diDir.Create();
                }
            }
            catch (Exception ex)
            {
                bValidacion = false;
                MessageBox.Show(ex.ToString());
            }
            return bValidacion;
        }

        bool CopiarArchivoExcel()
        {
            bValidacion = true;
            try
            {
                strRutaAux = strRutaFacturasResultado + "\\" + strNombreExcel;
                if (!File.Exists(strRutaAux))
                {
                    File.Copy(strRutaExcelOrig, strRutaAux);
                }
                strRutaExcel = strRutaAux;
                //MessageBox.Show(strRutaExcel);
            }
            catch (Exception)
            {
                bValidacion = false;
            }
            return bValidacion;
        }

        async Task<Thread> Procesar()
        {
            //Creo el objeto de el archivo excel
            Excel.Application oXL = new Excel.Application();
            //Lo oculto porque por defecto lo deja abierto
            oXL.Visible = false;
            //Abro el archivo de la ruta
            Excel._Workbook wb = oXL.Workbooks.Open(strRutaExcel);
            //Me coloco en su primer hoja
            Excel._Worksheet excelSheet = wb.ActiveSheet;
            bool recorrer = true;
            //dynamic dynRenglonActual = "";
            intCantidadRegistros = excelSheet.UsedRange.Rows.Count - 1;
            intCantidadRegistros += 3;
            int intFila = 5;
            intColFactura = 2;
            intColPoliza = 4;
            intColMes = 5;
            intColGrupo = 6;
            intColEncontrado = 7;
            dirDirFacturas = new DirectoryInfo(strRutaFacturas);
            //Recorrer los datos
            recorrer = true;
            while (recorrer & (intFila<=intCantidadRegistros))
            {
                try
                {
                    strFactura = excelSheet.Cells[intFila, intColFactura].Value;
                    if (Object.ReferenceEquals(null, strFactura))
                    {
                        strFactura = "";
                    }

                    strMes = excelSheet.Cells[intFila, intColMes].Value;
                    if (Object.ReferenceEquals(null, strMes))
                    {
                        strMes = "";
                    }

                    strPoliza = excelSheet.Cells[intFila, intColPoliza].Value;
                    if (Object.ReferenceEquals(null, strPoliza))
                    {
                        strPoliza = "";
                    }

                    strGrupo = excelSheet.Cells[intFila, intColGrupo].Value.ToString();
                    if (Object.ReferenceEquals(null, strGrupo))
                    {
                        strGrupo = "";
                    }

                    try
                    {
                        strEncontrado = excelSheet.Cells[intFila, intColEncontrado].Value.ToString();
                        if (Object.ReferenceEquals(null, strEncontrado))
                        {
                            strEncontrado = "";
                        }
                    }
                    catch (Exception)
                    {
                        strEncontrado = "";
                    }


                    if (strFactura.CompareTo("") != 0 & strMes.CompareTo("") != 0 & strPoliza.CompareTo("") != 0 & strGrupo.CompareTo("") != 0)
                    {
                        if (strEncontrado.CompareTo("") == 0 || strEncontrado.CompareTo("0") == 0)
                        {
                            //if(listaDocumentos.Contains(strFactura))
                            //{
                            int intIndexTest = 0;
                            List<string> list = listaDocumentos.ToList();
                            var n = list.Where(i => i.Contains(strFactura));

                            intIndexTest = list.IndexOf(n.First());
                            //intIndexTest = list.IndexOf(strFactura);

                            //intIndexTest = Array.FindIndex(listaDocumentos, valor); 
                            strRutaArchivo = listaDocumentos[intIndexTest];

                            strGrupoAux = strGrupo.Last().ToString();
                            strRutaActualSub = strRutaFacturasResultado + "\\" + strMes + "\\" + strGrupoAux + "\\" + strGrupo;
                            diDir = new DirectoryInfo(strRutaActualSub);

                            if (!diDir.Exists) //Si el directorio no existe, lo creo
                            {
                                diDir.Create();
                            }

                            strRutaAux = strRutaActualSub + "\\" + strPoliza + "_" + strFactura + ".pdf";
                            strRutaAuxPoliza = strRutaActualSub + "\\" + strPoliza + "_0.pdf";
                            if (!File.Exists(strRutaAux))
                            {
                                DescargarFacturaFTP();
                                //if (DescargarFacturaFTP())
                                //{
                                //    MessageBox.Show("A huevo");
                                //}
                            }
                            excelSheet.Cells[intFila, intColEncontrado].Value = "1";
                            if (!File.Exists(strRutaAuxPoliza))
                            {
                                DescargarPolizaPDF();
                            }
                            //}

                            //Busco el archivo por su nombre completo
                            /*fiArchivos = dirDirFacturas.GetFiles(strFactura + ".pdf", SearchOption.AllDirectories);
                            if (fiArchivos.Length > 0)
                            {
                                strRutaArchivo = fiArchivos[0].FullName.ToString();

                                strGrupoAux = strGrupo.Last().ToString();
                                strRutaActualSub = strRutaFacturasResultado + "\\" + strMes + "\\" + strGrupoAux + "\\" + strGrupo;
                                diDir = new DirectoryInfo(strRutaActualSub);

                                if (!diDir.Exists) //Si el directorio no existe, lo creo
                                {
                                    diDir.Create();
                                }

                                strRutaAux = strRutaActualSub + "\\" + strPoliza + "_" + strFactura + ".pdf";
                                strRutaAuxPoliza = strRutaActualSub + "\\" + strPoliza + "_0.pdf";
                                if (!File.Exists(strRutaAux))
                                {
                                    File.Copy(strRutaArchivo, strRutaAux);
                                }
                                excelSheet.Cells[intFila, intColEncontrado].Value = "1";
                                    if (!File.Exists(strRutaAuxPoliza))
                                    {
                                        DescargarPolizaPDF();
                                    }*/
                        }
                        else { excelSheet.Cells[intFila, intColEncontrado].Value = "0"; }
                    }


                    //Thread.Sleep(200);
                    intConteo++;
                    AumentarProgreso(((intConteo * 100) / (intCantidadRegistros - 3)));
                    intFila++;
                }
                catch (Exception e)
                {
                    recorrer = false;
                    MessageBox.Show(e.ToString());
                }
            }

            //excelSheet.SaveAs(strRutaExcel);
            
            if (oXL != null)
            {
                oXL.Quit();
            }
            
            return Thread.CurrentThread;
        }

        bool valor()
        {
            if (listaDocumentos.Contains(strFactura))
            {
                return true;    
            }
            else
            {
                return false;
            }
        }

        void DescargarFacturaFTP()
        {
            try
            {
                using (WebClient request = new WebClient())
                {
                    request.Credentials = new NetworkCredential("Noan", "ItSp2021/");
                    byte[] fileData = request.DownloadData(strRutaArchivo);

                    using (FileStream file = File.Create(strRutaAux))
                    {
                        file.Write(fileData, 0, fileData.Length);
                        file.Close();
                    }
                    //MessageBox.Show("Download Complete");
                }

                //byte[] fileData = request.DownloadData(strRutaArchivo);

                //using (FileStream file = File.Create(strRutaAux))
                //{
                //    file.Write(fileData, 0, fileData.Length);
                //    file.Close();
                //}
            }
            catch (Exception)
            {
            }
            
        }
        void DescargarPolizaPDF()
        {
            try
            {
                if (strPoliza.StartsWith("TRA"))
                {
                    strPolizaBuscar = strPoliza.Substring(strPoliza.Length - 6, 6);
                    //strPolizaBuscar = strPoliza.Substring(7, 6);
                    webclient.DownloadFile("http://inspect.mht-jcv.com/html/Empleado/Polizas/pdfpolizasgsus.php?numPoliza=" + strPolizaBuscar, strRutaAuxPoliza);
                }
            }
            catch (Exception e)
            {
            }
            
        }
    }
}
