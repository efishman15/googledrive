using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;

namespace GoogleDrive
{
    public partial class frmMain : Form
    {
        Drive drive;
        public frmMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            drive = new Drive();
            //drive.BuildPresentationsList(ConfigurationManager.AppSettings["rootFolderId"]);
            //drive.SavePresentationsList();
        }
    }
}
