using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WF.BASE.Helper;
using WF.BASE.Models;
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
            var page = txtPage.Text;
            if (string.IsNullOrEmpty(token))
            {
                MessageBox.Show("Vui lòng nhập TOKEN", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGetData.Enabled = true;
            }
            else
            {
                var lstDataBaseRequest = new List<Base.Responce.Request>();
                if (!string.IsNullOrEmpty(page))
                {
                    var lstPage = page?.Split(',')?.Select(Int32.Parse)?.ToList();
                    if (lstPage != null && lstPage.Count > 0)
                    {
                        foreach (var pageIndex in lstPage)
                        {
                            var result = new BaseService().GetAllRequest(new Base.Request.GetAllRequest() { Token = token, Page = pageIndex });
                            if (result != null && result.requests != null && result.requests.Count > 0)
                                lstDataBaseRequest.AddRange(result.requests);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Kiểm tra lại định dạng page", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btnGetData.Enabled = true;
                    }
                }
                else
                {
                    for (int i = 0; i < 50; i++)
                    {
                        var result = new BaseService().GetAllRequest(new Base.Request.GetAllRequest() { Token = token, Page = i });
                        if (result != null && result.requests != null && result.requests.Count > 0)
                            lstDataBaseRequest.AddRange(result.requests);
                    }
                }


                if (lstDataBaseRequest != null && lstDataBaseRequest.Count > 0)
                {
                    loading.Maximum = lstDataBaseRequest.Count;
                    loading.Minimum = 0;
                    loading.Step = 1;
                    loading.Style = ProgressBarStyle.Continuous;
                    var i = 0;
                    var lstDataExcel = new List<Excel.Responce.ExportDataRequest>();
                    foreach (var item in lstDataBaseRequest)
                    {
                        try
                        {
                            lstDataExcel.Add(new Excel.Responce.ExportDataRequest()
                            {
                                Id = item.id,
                                Title = ConvertExtensions.RemoveSpecialCharacters(item.name),
                                Content = ConvertExtensions.RemoveSpecialCharacters(item.content)
                            });
                            if (item.files != null && item.files.Count > 0)
                            {
                                var fileFolderImage = item.id;
                                foreach (var image in item.files)
                                {
                                    try
                                    {
                                        DownloadImage(image.url, fileFolderImage, image.name);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("StackTrace: " + ex.StackTrace + " | Message: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                            loading.Value = i++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("StackTrace: " + ex.StackTrace + " | Message: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    ExportExcel(lstDataExcel);
                }
                else
                    MessageBox.Show("Có lỗi không mong muốn sảy ra. Vui lòng thử lại sau", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);

                loading.Value = lstDataBaseRequest.Count;
                btnGetData.Enabled = true;
                MessageBox.Show("Chạy xong dữ liệu", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void StartLoading(bool start)
        {
            for (int i = 0; i < 1000; i++)
            {
                loading.Value = i;
            }
            if (start)
            {
                loading.Value = 1000;
            }
        }
        private static void DownloadImage(string url, string filePath, string fileName)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (WebClient client = new WebClient())
            {
                var pathFile = currentPathFolder + "//File//" + filePath;
                if (!Directory.Exists(pathFile))
                {
                    Directory.CreateDirectory(pathFile);
                }
                client.DownloadFile(new Uri(url), $"{pathFile}\\{fileName}");
            }
        }
        private void ExportExcel(List<Excel.Responce.ExportDataRequest> entity)
        {
            List<string> lstHeaderLoan = new List<string> { "Id", "Tiêu đề", "Content" };
            List<int> lstWidthLoan = new List<int> { 30, 50, 50 };
            List<ExcelHorizontalAlignment> lstAlign = new List<ExcelHorizontalAlignment> { ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           ExcelHorizontalAlignment.Left,
                                                                                           };
            List<System.Drawing.Color> lstColor = new List<System.Drawing.Color> { Color.Black,
                                                                                   Color.Black,
                                                                                   Color.Black,
                                                                                  };
            byte[] fileContents;
            ExcelPackage package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Dữ liệu");
            ws.Cells.Style.Font.Name = "Times New Roman";
            ws.Cells.Style.Font.Size = 12;
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
            for (int i = 0; i < entity.Count(); i++)
            {
                ws.Cells[iPosHeader, 1].Value = entity[i].Id;
                ws.Cells[iPosHeader, 2].Value = entity[i].Title;
                ws.Cells[iPosHeader, 3].Value = entity[i].Content;
                iPosHeader++;
            }

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
