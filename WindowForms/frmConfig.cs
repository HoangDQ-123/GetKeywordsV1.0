using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
namespace GetKeywords
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            InitVar.v_VolMax = Convert.ToInt32(txtVolMax.Text);
            InitVar.v_speed = Convert.ToInt32(txtSpeed.Text);
            InitVar.v_LevelSearch = Convert.ToInt32(txtLevelSearch.Text);
            InitVar.v_LevelDif = Convert.ToInt32(txtLevelDif.Text);
            InitVar.v_VolMin = Convert.ToInt32(txtVolMin.Text);
            InitVar.v_AutoSave = Convert.ToInt32(txtAutoSave.Text);
            InitVar.v_TimeDownload = Convert.ToInt32(txtTimeDownload.Text);
            InitVar.v_TimeDownload2 = Convert.ToInt32(txtTimeDownload2.Text);
            InitVar.v_TimeSearch    = Convert.ToInt32(txtTimeSearch.Text);  
            InitVar.v_MaxCountDownload = Convert.ToInt32(txtMaxCountDownload.Text);

            InitVar.SaveFileConfig(InitVar.pathConfig);

            this.Close();
        }

        

        private void frmConfig_Load(object sender, EventArgs e)
        {
            InitVar.OpenFileConfig(InitVar.pathConfig);

            txtVolMax.Text = Convert.ToString(InitVar.v_VolMax);
            txtSpeed.Text = Convert.ToString(InitVar.v_speed);
            txtLevelSearch.Text = Convert.ToString(InitVar.v_LevelSearch);
            txtLevelDif.Text = Convert.ToString(InitVar.v_LevelDif);
            txtVolMin.Text = Convert.ToString(InitVar.v_VolMin);
            txtAutoSave.Text = Convert.ToString(InitVar.v_AutoSave);
            txtTimeDownload.Text = Convert.ToString(InitVar.v_TimeDownload);
            txtTimeDownload2.Text = Convert.ToString(InitVar.v_TimeDownload2);
            txtTimeSearch.Text = Convert.ToString(InitVar.v_TimeSearch);
            txtMaxCountDownload.Text = Convert.ToString(InitVar.v_MaxCountDownload);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
           //Form1.InitVar_Plan();
        }
    }
}
