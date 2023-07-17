using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WF.BASE.Helper;
using WF.BASE.Service;

namespace WF.BASE
{
    public partial class BASE : Form
    {
        public BASE()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            createdAt.Format = DateTimePickerFormat.Custom;
            createdAt.CustomFormat = "dd/MM/yyyy";
            createdAt.Value = ConvertExtensions.FirstDayOfMonth(DateTime.Now);
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            btnGetData.Enabled = false;
            var token = txtToken.Text;
            if (string.IsNullOrEmpty(token))
            {
                MessageBox.Show("Vui lòng nhập TOKEN", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGetData.Enabled = true;
            }
            else
            {
                var result = new BaseService().GetAllRequest(new Models.Base.Request.GetAllRequest());
                StartLoading(true);
            }
        }

        private void StartLoading(bool start)
        {
            loading.Maximum = 1000;
            loading.Minimum = 0;
            loading.Step = 1;
            loading.Style = ProgressBarStyle.Continuous;
            for (int i = 0; i < 1000; i++)
            {
                loading.Value = i;
            }
            if (start)
            {
                loading.Value = 1000;
            }
        }
    }
}
