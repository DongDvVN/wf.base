using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
        static string currentPathFolder = "C://WF.BASE";
        public BASE()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //createdAt.Format = DateTimePickerFormat.Custom;
            //createdAt.CustomFormat = "dd/MM/yyyy";
            //createdAt.Value = ConvertExtensions.FirstDayOfMonth(DateTime.Now);
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
                var result = new BaseService().GetAllRequest(new Models.Base.Request.GetAllRequest() { Token = token });
                MessageBox.Show(result, "KẾT QUẢ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                StartLoading(true);
                btnGetData.Enabled = true;
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

        private void ExportExcel()
        {
            List<string> lstHeaderLoan = new List<string> { "loai_bds", "ho_ten", "email", "sdt", "ten_cong_ty", "ma_nhan_vien", "chuc_vu", "quyen_so_huu", "tinh_tp", "quan_huyen", "phuong_xa", "dia_chi", "tieu_de", "ten_du_an", "lo_block", "so_can_ho", "ten_chung_cu", "dien_tich", "phong_ngu", "phong_tam", "do_rong_duong", "ket_cau_nha", "huong_nha", "huong_dat", "nam_o_tang", "chieu_rong", "chieu_dai", "huong_cua_chinh", "noi_that", "huong_ban_cong", "gia_tien", "mo_ta", "thuong_luong", "phi_sang_ten", "giay_to_phap_ly", "tien_ich_xung_quanh", "tien_ich_khac", "trang_thiet_bi" };
            List<int> lstWidthLoan = new List<int> { 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20 };
            List<ExcelHorizontalAlignment> lstAlign = new List<ExcelHorizontalAlignment> { ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           };
            List<System.Drawing.Color> lstColor = new List<System.Drawing.Color> { Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                  };
            byte[] fileContents;
            ExcelPackage package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Dữ liệu");
            ws.Cells.Style.Font.Name = "Times New Roman";
            ws.Cells.Style.Font.Size = 11;
            // Format All Cells
            ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            // Format All Column
            for (int i = 1; i <= lstHeaderLoan.Count; i++)
            {
                ws.Column(i).Width = lstWidthLoan[i - 1];
                ws.Column(i).Style.HorizontalAlignment = lstAlign[i - 1];
                ws.Column(i).Style.Font.Color.SetColor(lstColor[i - 1]);
            }
            int iPosHeader = 1;
            // mapping title header
            ws.Cells[iPosHeader, 1, iPosHeader, lstHeaderLoan.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[iPosHeader, 1, iPosHeader, lstHeaderLoan.Count].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(77, 119, 204));
            ws.Cells[iPosHeader, 1, iPosHeader, lstHeaderLoan.Count].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(255, 255, 255));
            ws.Cells[iPosHeader, 1, iPosHeader, lstHeaderLoan.Count].Style.Font.Bold = true;
            ws.Row(iPosHeader).Height = 16;
            for (int i = 1; i <= lstHeaderLoan.Count; i++)
            {
                ws.Cells[iPosHeader, i].Value = lstHeaderLoan[i - 1];
            }
            // SET VALUE
            iPosHeader++;
            //for (int i = 0; i < entity.Count(); i++)
            //{

            //    iPosHeader++;
            //}

            fileContents = package.GetAsByteArray();
            string pathFile = currentPathFolder + "//FileExcel";
            if (!Directory.Exists(pathFile))
            {
                Directory.CreateDirectory(pathFile);
            }
            string fileName = $"DS_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm")}.xlsx";
            var path = $"{pathFile}\\{fileName}";
            File.WriteAllBytes(path, fileContents);
        }
    }
}
