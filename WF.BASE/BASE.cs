using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
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
            var token = txtToken.Text;
            if (string.IsNullOrEmpty(token))
                MessageBox.Show("Vui lòng nhập TOKEN", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                var result = new BaseService().GetAllRequest(new Models.Base.Request.GetAllRequest());
            }
        }
    }
}
