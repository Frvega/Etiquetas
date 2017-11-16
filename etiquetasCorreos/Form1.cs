using MsgKit;
using MsgKit.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace etiquetasCorreos
{
    public partial class Form1 : Form
    {

        private FolderBrowserDialog folderBrowserDialog;
        private OpenFileDialog openFileDialog1;

        private RichTextBox richTextBox1;

        private MainMenu mainMenu1;
        private MenuItem fileMenuItem, openMenuItem;
        private MenuItem folderMenuItem, closeMenuItem;

        private string openFileName, folderName;

        private bool fileOpened = false;


        public Form1()
        {
            //MsgReader.Outlook.Storage.Message outlookMsg = new OutlookStorage.Message(@"C:\test.msg");
            //DisplayMessage(outlookMsg);
            var etiquetas = new object[] {
            "[nombreGDCPM]",
            "[correoGDCPM]",
            "[nombreGerenciaFSW]",
            "[correoGerenciaFSW]",
            "[nombreLiderDesarollo]",
            "[CorreoLiderDesarrollo]",
            "[nombreDesarrollador]",
            "[correoDesarrollador]",
            "[nombreClienteFSW]",
            "[correoClienteFSW]",
            "[nombreCM]",
            "[correoCM]"};
            InitializeComponent();
            dataGridView1.RowCount = etiquetas.Length;
            for(int i=0;i< etiquetas.Length; i++)
            {


                dataGridView1.Rows[i].Cells[0].Value = etiquetas[i];
                
              
            }

            this.mainMenu1 = new System.Windows.Forms.MainMenu();
            this.fileMenuItem = new System.Windows.Forms.MenuItem();
            this.openMenuItem = new System.Windows.Forms.MenuItem();
            this.folderMenuItem = new System.Windows.Forms.MenuItem();
            this.closeMenuItem = new System.Windows.Forms.MenuItem();

            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();

            this.mainMenu1.MenuItems.Add(this.fileMenuItem);
            this.fileMenuItem.MenuItems.AddRange(
                                new System.Windows.Forms.MenuItem[] {
                                                                 
                                                                 this.folderMenuItem});
            this.fileMenuItem.Text = "File";

            
            this.folderMenuItem.Text = "Select Directory...";
            this.folderMenuItem.Click += new System.EventHandler(this.folderMenuItem_Click);

            

            

            this.folderBrowserDialog1.Description =
                "Select the directory that you want to use as the default.";

            this.folderBrowserDialog1.ShowNewFolderButton = false;

            this.folderBrowserDialog1.RootFolder = Environment.SpecialFolder.DesktopDirectory;

            this.richTextBox1.AcceptsTab = true;
            this.richTextBox1.Location = new System.Drawing.Point(8, 8);
            this.richTextBox1.Anchor = AnchorStyles.Top | AnchorStyles.Left |
                                       AnchorStyles.Bottom | AnchorStyles.Right;

            this.Menu = this.mainMenu1;
            this.Text = "RTF Document Browser";


        }

        private void openMenuItem_Click(object sender, System.EventArgs e)
        {
            if (!fileOpened)
            {
                openFileDialog1.InitialDirectory = folderBrowserDialog1.SelectedPath;
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                openFileDialog1.FileName = null;
            }

            // Display the openFile dialog.
            DialogResult result = openFileDialog1.ShowDialog();

            
        }


        private void closeMenuItem_Click(object sender, System.EventArgs e)
        {
            richTextBox1.Text = "";
            fileOpened = false;

            closeMenuItem.Enabled = false;
        }


        private void folderMenuItem_Click(object sender, System.EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                folderName = folderBrowserDialog1.SelectedPath;
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                           }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            {
                {
                    {
                        var a = dataGridView1.Rows.Count;
                        string[] textoreplace;
                        textoreplace = new string[a];
                        for (int i = 0; i < etiquetas.Length; i++)
                        {
                            if (Convert.ToString(dataGridView1.Rows[i].Cells[1].Value) != "")
                            {
                                textoreplace[i] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                            }
                            else
                            {
                                MessageBox.Show("Favor de No Dejar Celdas en Blanco");
                                break;
                            }

                            //MessageBox.Show(Convert.ToString(textoreplace[i]));


                        }





                        var A2 = "<b>From: [nombreGDCPM] <[correoGDCPM]></b> <b>Sent: [fecha] </b> <b>To: [nombreClienteFSW] < [correoClienteFSW] ></b> <b>Subject: [tema]</b> Hola [nombreClienteFSW],<br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto].<br><br> Los integrantes del equipo son:<br><br> GDC PM: [nombreGDCPM]<br> Lider de Proyecto: [nombreLiderDesarollo]<br> Desarrollador: [nombreDesarrollador]<br> Y el gerente de Fábrica de Software: [nombreGerenciaFSW] <br> Cliente: [nombreClienteFSW]<br><br><br> Saludos<br><br><hr> From: [nombreGDCPM] <[correoGDCPM]> Sent: [fecha] To: [nombreClienteFSW] <[correoClienteFSW]>; [nombreGerenciaFSW] <[correoGerenciaFSW]>; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > Subject: [tema] Hola [nombre], Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. [pathDocGuiaAjuste] la guía de ajuste no aplica en este caso. sgcservicio.clientes@neoris.com Saludos From: [nombreClienteFSW] <[correoClienteFSW]> Sent: [fecha] To: [nombreGerenciaFSW] <[correoGerenciaFSW]> Subject: [tema] Hola [nombreGerenciaFSW], La estimación fue aprobada. Saludos. From: [nombreGDCPM] <[correoGDCPM]> Sent: [fecha] To: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]>; [nombreGerenciaFSW] <[correoGerenciaFSW]>; [nombreClienteFSW] <[correoClienteFSW]> Subject: [tema] Hola [nombreLiderDesarollo], Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. Saludos. From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]> Sent: [fecha] To: [nombreGDCPM] <[correoGDCPM]>; [nombreGerenciaFSW] <[correoGerenciaFSW]> Subject: [tema] Hola [nombreGDCPM] Anexo path donde está el documento de la estimación: [pathDocEstimacion] Y el de la calculadora: [pathDocCalculadora] ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. Saludos. From: [nombreGDCPM]<[correoGDCPM]> Sent: [fecha] To: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]>; [nombreClienteFSW] <[correoClienteFSW]> Subject: [tema] Hola [nombreLiderDesarollo], [nombreClienteFSW] Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: [path] Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): [pathRepositorioNeoris] También fue versionado en el siguiente path de TFS: [pathTFS] En este mismo path puede versionar el documento de estimación. El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): 23291 CapacidadEntrega-CC I1 @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. Saludos.";
                        var A5_2_ = "<b>From: [nombreGDCPM] < [correoGDCPM] ></b><br> <b>Sent: [fecha]</b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] ></b><br> <b>Subject: [tema]</b><br><br><br> Hola [nombreClienteFSW],<br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto].<br><br> Los integrantes del equipo son:<br><br> GDC PM: [nombreGDCPM]<br> Lider de Proyecto: [nombreLiderDesarollo]<br> Desarrollador: [nombreDesarrollador]<br> Y el gerente de Fábrica de Software: [nombreGerenciaFSW] <br> Cliente: [nombreClienteFSW]<br><br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b> <br> <b>Sent: [fecha] </b> <br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com>; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br><br> [pathDocGuiaAjuste] <br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos <br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]>; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; </b><br> <b>Subject: [tema] </b><br><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> [pathDocEstimacion] <br><br> Y el de la calculadora: <br><br> [pathDocCalculadora] <br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br> <div style=”color:#1F497D”> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> [pathDocReview] <br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> [pathRequerimientoEnRepositorioNeoris] <br><br> También fue versionado en el siguiente path de TFS: <br> [pathRequerimientoEnTFS] <br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1 <br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br> </div>";

                        var A1_2 = "<b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto]. <br><br> Los integrantes del equipo son: <br><br> <b>GDC PM: </b> [nombreGDCPM]. <br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>Desarrollador:</b> [nombreDesarrollador]. <br> <b>Y el gerente de Fábrica de Software:</b> [nombreGerenciaFSW]. <br> <b>Cliente:</b> [nombreClienteFSW]. <br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br> <b> [pathDocGuiaAjuste] </b><br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos<br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> <b> [pathDocEstimacion] </b><br><br> Y el de la calculadora: <br> <b> [pathDocCalculadora] </b><br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> <b> [pathPeerReview] </b><br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> <b> [pathRepositorioNeoris] </b><br><br> También fue versionado en el siguiente path de TFS: <br><br> <b> [pathRepositorioTFS] </b><br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1<br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br><br>";
                        var A2_2 = "<b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreGerenciaFSW], [nombreGDCPM] <br> Confirmo de enterado y estoy de acuerdo con tus comentarios. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto]. <br><br> Los integrantes del equipo son: <br><br> <b>GDC PM: </b> [nombreGDCPM]. <br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>Desarrollador:</b> [nombreDesarrollador]. <br> <b>Y el gerente de Fábrica de Software:</b> [nombreGerenciaFSW]. <br> <b>Cliente:</b> [nombreClienteFSW]. <br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br> <b> [pathDocGuiaAjuste] </b><br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos<br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> <b> [pathDocEstimacion] </b><br><br> Y el de la calculadora: <br> <b> [pathDocCalculadora] </b><br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> <b> [pathPeerReview] </b><br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> <b> [pathRepositorioNeoris] </b><br><br> También fue versionado en el siguiente path de TFS: <br><br> <b> [pathRepositorioTFS] </b><br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1<br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br><br>";
                        var A3_2 = "<b>From: [nombreGerenciaFSW] < [correoGerenciaFSW] >;</b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreCM] < [CorreoCM] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola Equipo, buen día<br><br> Enterado y adelante con las fechas y comentarios expuestos en los correos. Agrego a la persona CM para la generación del repositorio. <br><br> CM (Configuration Manager): [nombreCM]. <br><br> @[nombreLiderDesarollo]: por favor coordinate con [nombreCM] para el path del repositorio. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto]. <br><br> Los integrantes del equipo son: <br><br> <b>GDC PM: </b> [nombreGDCPM]. <br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>Desarrollador:</b> [nombreDesarrollador]. <br> <b>Y el gerente de Fábrica de Software:</b> [nombreGerenciaFSW]. <br> <b>Cliente:</b> [nombreClienteFSW]. <br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br> <b> [pathDocGuiaAjuste] </b><br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos<br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> <b> [pathDocEstimacion] </b><br><br> Y el de la calculadora: <br> <b> [pathDocCalculadora] </b><br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> <b> [pathPeerReview] </b><br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> <b> [pathRepositorioNeoris] </b><br><br> También fue versionado en el siguiente path de TFS: <br><br> <b> [pathRepositorioTFS] </b><br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1<br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br><br>";
                        var A4_2 = "<b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreClienteFSW] < [correoClienteFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreCM] < [correoCM] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreCM], <br><br> Estaremos trabajando en el path original: <br> <b> [pathElegidoParaProyecto] </b><br><br> Si gustas en base a dicho path por favor ayúdanos con el plan de la configuración. <br><br> Saludos y gracias por el apoyo. <br><br><hr> <b>From: [nombreGerenciaFSW] < [correoGerenciaFSW] >;</b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreCM] < [CorreoCM] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola Equipo, buen día<br><br> Enterado y adelante con las fechas y comentarios expuestos en los correos. Agrego a la persona CM para la generación del repositorio. <br><br> CM (Configuration Manager): [nombreCM]. <br><br> @[nombreLiderDesarollo]: por favor coordinate con [nombreCM] para el path del repositorio. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto]. <br><br> Los integrantes del equipo son: <br><br> <b>GDC PM: </b> [nombreGDCPM]. <br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>Desarrollador:</b> [nombreDesarrollador]. <br> <b>Y el gerente de Fábrica de Software:</b> [nombreGerenciaFSW]. <br> <b>Cliente:</b> [nombreClienteFSW]. <br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br> <b> [pathDocGuiaAjuste] </b><br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos<br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> <b> [pathDocEstimacion] </b><br><br> Y el de la calculadora: <br> <b> [pathDocCalculadora] </b><br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> <b> [pathPeerReview] </b><br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> <b> [pathRepositorioNeoris] </b><br><br> También fue versionado en el siguiente path de TFS: <br><br> <b> [pathRepositorioTFS] </b><br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1<br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br><br>";
                        var A5_2 = "<b>From: [nombreCM] <[correoCM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreClienteFSW] < [correoClienteFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreDesarrollador] < [correoDesarrollador] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola Equipo, <br><br> De acuerdo y anexo el path del plan de configuración a como lo conversamos en llamada por teléfono: <br> <b> [pathPlanConfiguracion] </b><br><br> A su vez veo que ya se versionaron los documentos y por este medio valido que efectivamente fueron versionados correctamente. <br><br> Saludos y gracias por el apoyo. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreClienteFSW] < [correoClienteFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreCM] < [correoCM] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreCM], <br><br> Estaremos trabajando en el path original: <br> <b> [pathElegidoParaProyecto] </b><br><br> Si gustas en base a dicho path por favor ayúdanos con el plan de la configuración. <br><br> Saludos y gracias por el apoyo. <br><br><hr> <b>From: [nombreGerenciaFSW] < [correoGerenciaFSW] >;</b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreCM] < [CorreoCM] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola Equipo, buen día<br><br> Enterado y adelante con las fechas y comentarios expuestos en los correos. Agrego a la persona CM para la generación del repositorio. <br><br> CM (Configuration Manager): [nombreCM]. <br><br> @[nombreLiderDesarollo]: por favor coordinate con [nombreCM] para el path del repositorio. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> El inicio del proyecto es del [fechaInicioProyecto] al [fechaFinProyecto]. <br><br> Los integrantes del equipo son: <br><br> <b>GDC PM: </b> [nombreGDCPM]. <br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>Desarrollador:</b> [nombreDesarrollador]. <br> <b>Y el gerente de Fábrica de Software:</b> [nombreGerenciaFSW]. <br> <b>Cliente:</b> [nombreClienteFSW]. <br><br> Saludos<br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; Servicio Clientes NSF <sgcservicio.clientes@neoris.com > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Una vez revisado nuestro proceso interno (Guía de Ajustes y el Plan de Calidad), anexo el mail para servicio a clientes con la finalidad de proporcionar mayor control y seguimiento a peticiones. <br> <b> [pathDocGuiaAjuste] </b><br><br> la guía de ajuste no aplica en este caso. <br><br> sgcservicio.clientes@neoris.com <br><br> Saludos<br><br><hr> <b>From: [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGerenciaFSW], <br><br> La estimación fue aprobada. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreClienteFSW] < [correoClienteFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Es correcto la estimación es correcta y cuenta con el Vo.Bo de GDC PM y el Gerente de Fábrica. @[nombreClienteFSW]: nos puedes ayudar con la aprobación por favor. <br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br> Hola [nombreGDCPM], <br><br> Anexo path donde está el documento de la estimación: <br> <b> [pathDocEstimacion] </b><br><br> Y el de la calculadora: <br> <b> [pathDocCalculadora] </b><br><br> ´ La calculadora tiene el Vo.Bo. de [nombreGerenciaFSW] (Gerente de Fabrica) y [nombreGDCPM] (GDC PM), visto hace unos momentos en su lugar y a quienes agregó en este correo como evidencia de su Vo.Bo. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreClienteFSW] < [correoClienteFSW] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreClienteFSW] <br><br> Posterior al review de la documentación de entrada, anexo el path donde se encuentra dicho documento de peer review: <br> <b> [pathPeerReview] </b><br><br> Favor de realizar la estimación en base al documento de entrada, se proporcionó en la siguiente ruta y también fue versionado en TFS. <br><br> Anexo el path donde se encuentra el requerimiento (ya en repositorio de Neoris): <br> <b> [pathRepositorioNeoris] </b><br><br> También fue versionado en el siguiente path de TFS: <br><br> <b> [pathRepositorioTFS] </b><br><br> En este mismo path puede versionar el documento de estimación. <br><br> El ticket y nombre del proyecto oficial es el siguiente (el número es un identificador único proporcionado por el cliente y para mayor referencia con el mismo): <br><br> 23291 CapacidadEntrega-CC I1<br><br> @[nombreClienteFSW]: el requerimiento contiene información suficiente para realizar una estimación, por lo cual procedemos a realizar dicho documento. <br><br> Saludos. <br><br>";
                        var B1_2 = "<b>From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Nos puedes apoyar con el plan de pruebas Por Favor. <br><br> Saludos. <br>"; 
                        var B2_2 = "<b>From: [nombreClienteFSW] <[correoClienteFSW]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Anexo el path del plan de pruebas, favor de agregar los cambios necesarios, como validación de cada actividad por parte del desarrollador y vo.bo. del líder de proyectos. <br> <b> [pathPlanPruebas] </b><br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Nos puedes apoyar con el plan de pruebas Por Favor. <br><br> Saludos. <br><br><hr>";
                        var B4_2 = "<b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreGerenciaFSW] <br><br> El equipo persiste como se mencionó en correos anteriores, <br><br> <b>Lider de Proyecto: </b>[nombreLiderDesarollo] <br> <b>GDC PM: </b>[nombreGDCPM] <br> <b>Desarrollador: </b>[nombreDesarrollador] <br> <b>CM: </b>[nombreCM] <br> <b>Cliente: </b>[nombreClienteFSW] <br><br> Solicito el plan de desarrollo a detalle para llevar el seguimiento de cada actividad. <br><br> Saludos. <br>";
                        var B5_2 = "<b>From: [nombreGerenciaFSW] < [correoGerenciaFSW] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [asunto] </b><br><br><br> Hola [nombreGDCPM], <br><br> De acuerdo con el equipo solicitado y ya han sido asignados duante el tiempo descrito en la estimación o propuesta de servicio. <br><br> <b>Lider de Proyecto:</b> [nombreLiderDesarollo]. <br> <b>GDC PM: </b>[nombreGDCPM]. <br> <b>Desarrollador: </b>[nombreDesarrollador]. <br> <b>CM: </b>[nombreCM]. <br> <b>Cliente: </b>[nombreClienteFSW]. <br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreGerenciaFSW] <br><br> El equipo persiste como se mencionó en correos anteriores, <br><br> <b>Lider de Proyecto: </b>[nombreLiderDesarollo] <br> <b>GDC PM: </b>[nombreGDCPM] <br> <b>Desarrollador: </b>[nombreDesarrollador] <br> <b>CM: </b>[nombreCM] <br> <b>Cliente: </b>[nombreClienteFSW] <br><br> Solicito el plan de desarrollo a detalle para llevar el seguimiento de cada actividad. <br><br> Saludos. <br>";
                        var B6_2 = "<b>From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGDCPM] <[correoGDCPM]>; [nombreGerenciaFSW] <[correoGerenciaFSW]>; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreGDCPM], <br><br> Anexo el path donde se estará versionando el plan de trabajo: <br> <b> [pathPlanTrabajo] </b><br><br> Saludos. <br><br><hr> <b>From: [nombreGDCPM] < [correoGDCPM] > </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreLiderDesarollo] < [CorreoLiderDesarrollo] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], [nombreGerenciaFSW] <br><br> El equipo persiste como se mencionó en correos anteriores, <br><br> <b>Lider de Proyecto: </b>[nombreLiderDesarollo] <br> <b>GDC PM: </b>[nombreGDCPM] <br> <b>Desarrollador: </b>[nombreDesarrollador] <br> <b>CM: </b>[nombreCM] <br> <b>Cliente: </b>[nombreClienteFSW] <br><br> Solicito el plan de desarrollo a detalle para llevar el seguimiento de cada actividad. <br><br> Saludos. <br>";
                        var B3_2 = "<b>From: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGDCPM] < [correoGDCPM] >; [nombreGerenciaFSW] < [correoGerenciaFSW] > </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreGerenciaFSW], <br><br> De acuerdo con el plan, lo estaremos actualizando semanalmente una vez inicie el desarrollo. <br><br> Saludos. <br><br><hr> <b>From: [nombreClienteFSW] <[correoClienteFSW]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreLiderDesarollo] < [CorreoLiderDesarrollo] >; [nombreGerenciaFSW] < [correoGerenciaFSW] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreLiderDesarollo], <br><br> Anexo el path del plan de pruebas, favor de agregar los cambios necesarios, como validación de cada actividad por parte del desarrollador y vo.bo. del líder de proyectos. <br> <b> [pathPlanPruebas] </b><br><br> Saludos. <br><br><hr> <b>From: [nombreLiderDesarollo] <[CorreoLiderDesarrollo]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] < [correoClienteFSW] >; [nombreGDCPM] < [correoGDCPM] >; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Nos puedes apoyar con el plan de pruebas Por Favor. <br><br> Saludos. <br>";
                        var B7_2 = "<b>From: [nombreClienteFSW] <[correoClienteFSW]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] <[correoGerenciaFSW]>; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreGerenciaFSW], <br><br> Enterado y de acuerdo con el plan y fechas proporcionadas. <br><br> Saludos. <br><br>";
                        var B8_2 = "<b>From: [nombreClienteFSW] <[correoClienteFSW]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreGerenciaFSW] <[correoGerenciaFSW]>; </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreGerenciaFSW], <br><br> Enterado y de acuerdo con el plan y fechas proporcionadas. <br><br> Saludos. <br><br><hr> <b>From: [nombreGerenciaFSW] <[correoGerenciaFSW]> </b><br> <b>Sent: [fecha] </b><br> <b>To: [nombreClienteFSW] <[correoClienteFSW]> </b><br> <b>Subject: [tema] </b><br><br><br> Hola [nombreClienteFSW], <br><br> Anexo el path donde se estará versionando el plan de trabajo: <br> <b>[pathPlanTrabajo]</b> <br><br> La fecha inicio y fin del proyecto es la siguiente: <br> <b>Inicio:</b> [fechaInicioProyecto]. <br> <b>Fin:</b> [fechaFinProyecto]. <br> Favor de confirmar si estás de acuerdo. <br><br> Saludos. <br><br><hr>";

                        for (int i = 0; i < etiquetas.Length; i++)
                        {
                            var neew = Convert.ToString(textoreplace[i]);
                            var old = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                            A5_2 = A5_2.Replace(old, neew);

                        }


                        using (var email = new Email(
         new Sender("", ""),
         new Representing("", ""),
         ""))
                        {
                            //email.SentOn = new DateTime(2010,5,2,9,30,15,DateTimeKind.Local);
                            email.Recipients.AddTo("", "");
                            //email.Recipients.AddCc("crocodile@neverland.com", "The evil ticking crocodile");
                            email.Subject = "";
                            email.BodyText = "Cuerpo del Mensaje [dsd]";
                            email.BodyHtml = A5_2;
                            email.Importance = MessageImportance.IMPORTANCE_HIGH;
                            email.IconIndex = MessageIconIndex.ReadMail;
                            //email.Attachments.Add(@"d:\crocodile.jpg");
                            email.Save( @textBox1.Text+"\\Generados\\"+"correo1.msg");

                            // Show the E-mail
                            System.Diagnostics.Process.Start(@textBox1.Text + "\\Generados\\" + "correo1.msg");
                            }

                         

                        //                       if (textBox1.Text != string.Empty)
                        //                       {
                        //string[] files = Directory.GetFiles(textBox1.Text, "*.msg", SearchOption.AllDirectories);


                        //                       string result;
                        //                       //MessageBox.Show(files.ToArray().ToString());
                        //                       var i = 0;
                        //                       foreach (var file in files)
                        //                       {
                        //                           string text = File.ReadAllText(file);

                        //                           //File.WriteAllText(file, text);
                        //                           //MessageBox.Show(text);

                        //                           using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(textBox1.Text + "\\CorreosGenerados\\" + i + ".msg"))

                        //                           {text = text.Replace("[tema]", "FW: HolaPrueba Replace()");
                        //                               File.WriteAllText(file, text);
                        //                               escritor.WriteLine(text);
                        //                               escritor.Close();

                        //                                   //textBox3.Text = text;

                        //                               }
                        //                               //result = Path.ChangeExtension(textBox1.Text + "\\" + i + ".txt", ".msg");
                        //                               //MessageBox.Show(result);
                        //                               i++;
                        //                       }
                        //                       }
                        //                       else
                        //                       {
                        //                           MessageBox.Show("Introduzca un Directorio");
                        //                       }

                    }
                }
            }
        
            
            
            //    MailMessage msg = new MailMessage();
            //    msg.Dispose();
            //    MessageBox.Show("From: {0}", msg.From.ToString());
            //textBox1 = msg.From();
            //textBox1.Text = System.IO.File.ReadAllText("C:\\Users\\Roberto Vega\\Desktop\\FW_SolicitudDesarrolloARCA.msg");
            //textBox1.Text = System.IO.File.ReadAllText("C:\\Users\\Roberto Vega\\Desktop\\text.txt");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //System.IO.File.WriteAllText("C:\\Users\\Roberto Vega\\Desktop\\FW_SolicitudDesarrolloARCA.msg", textBox1.Text);
            ////System.IO.File.WriteAllText("C:\\Users\\Roberto Vega\\Desktop\\text.txt",textBox1.Text);
            //System.IO.StreamReader file = new System.IO.StreamReader("C:\\Users\\Roberto Vega\\Desktop\\FW_SolicitudDesarrolloARCA.msg");
            //file.Close();

           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //System.IO.StreamReader file = new System.IO.StreamReader("C:\\Users\\Roberto Vega\\Desktop\\FW_SolicitudDesarrolloARCA.msg");
            //file.Dispose();
            //file.Close();
            

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void propertyGrid1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
    }
}
