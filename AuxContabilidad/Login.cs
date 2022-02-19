using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AuxContabilidad
{
    public partial class Login : Form
    {
        //Clases
        //Variables
        string path = Directory.GetCurrentDirectory();
        //List<string> listDocuments = new List<string>();
        //WebClient request;
        string nombre, ip, user, password;
        string[] valoresConexion;
        string[] valoresIP;
        string[] valoresUsuarios;
        string[] valoresPassword;

        public Login()
        {
            InitializeComponent();

            nombre = Properties.Settings.Default.NOMBRE;
            ip = Properties.Settings.Default.IP;
            user = Properties.Settings.Default.USER;
            password = Properties.Settings.Default.PASSWORD;
            readConnections();
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

        private void cbxNombre_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtIP.Text = valoresIP[cbxNombre.SelectedIndex];
            txtUser.Text = valoresUsuarios[cbxNombre.SelectedIndex];
            txtPass.Text = valoresPassword[cbxNombre.SelectedIndex];
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            conexion();
        }

        void conexion()
        {
            string url = "ftp://" + txtIP.Text;

            //Form1 form = new Form1(FTPListTree(url + "/", txtUser.Text, txtPass.Text));
            //form.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*for (int i = 0; i < listDocuments.Count; i++)
            {
                MessageBox.Show(listDocuments[i].ToString());
                //Console.WriteLine(listDocuments[i].Name + " is " + listOfUsers[i].Age + " years old");
            }*/
        }

        void updateCredentials()
        {
            try
            {
                for (int i = 0; i < valoresConexion.Length; i++)
                {
                    if (valoresConexion[i].ToString() == cbxNombre.Text)
                    {
                        valoresIP[i] = txtIP.Text;
                        valoresUsuarios[i] = txtUser.Text;
                        valoresPassword[i] = txtPass.Text;
                    }
                }
            }
            catch (Exception)
            {

            }
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

            }catch (Exception){}
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
    }
}
