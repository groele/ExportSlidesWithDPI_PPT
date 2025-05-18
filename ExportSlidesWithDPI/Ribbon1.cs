using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ExportSlidesWithDPIDoing
{
    public partial class Ribbon1
    {
        private PowerPoint.Application app;
        private int currentDPI = 300;
        private string exportFormat = "jpg";
        private string saveFolderPath = string.Empty;
        private List<int> selectedPages = new List<int>();

        #region Ribbon 初始化
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
            InitializeControls();
            BindEvents();

            editBox1.Label = "Page";
            editBox1.Text = "0";
        }

        void InitializeControls()
        {
            comboBox1.Items.Clear();
            int[] dpiOptions = { 96, 150, 200, 300, 600 };
            foreach (int dpi in dpiOptions)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = dpi.ToString();
                comboBox1.Items.Add(item);
            }
            comboBox1.Text = currentDPI.ToString();

            comboBox2.Items.Clear();
            string[] formats = { "png", "jpg", "bmp" };
            foreach (string fmt in formats)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = fmt;
                comboBox2.Items.Add(item);
            }
            comboBox2.Text = exportFormat;
        }

        void BindEvents()
        {
            this.Button1.Click += Button1_Click;       // 选择保存路径按钮
            this.button2.Click += Button2_Click;       // 导出按钮
            this.comboBox1.TextChanged += ComboBox1_TextChanged; // DPI修改
            this.comboBox2.TextChanged += ComboBox2_TextChanged; // 格式修改
            this.editBox1.TextChanged += EditBox_TextChanged;    // 页码输入
            this.button3.Click += button3_Click;       // 关于开发者
            this.button4.Click += button4_Click_1;     // 打开网址
        }
        #endregion

        #region 事件处理

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "选择保存路径";
                    folderDialog.ShowNewFolderButton = true;
                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        saveFolderPath = folderDialog.SelectedPath;
                        MessageBox.Show($"保存路径已设置为：{saveFolderPath}", "路径设置",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"路径选择错误：{ex.Message}", "错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ComboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (int.TryParse(comboBox1.Text, out int newDpi) && newDpi > 0)
            {
                currentDPI = newDpi;
            }
            else
            {
                MessageBox.Show("无效的DPI值，请输入正整数", "输入错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox1.Text = currentDPI.ToString();
            }
        }

        private void ComboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var format = comboBox2.Text.ToLower();
            if (new[] { "png", "jpg", "bmp" }.Contains(format))
            {
                exportFormat = format;
            }
            else
            {
                MessageBox.Show("仅支持 PNG/JPG/BMP 格式", "格式错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Text = exportFormat;
            }
        }

        private void EditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var input = editBox1.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(input))
            {
                editBox1.Text = "all";
                return;
            }

            if (input != "all" && input != "0" && !ValidatePageInput(input))
            {
                MessageBox.Show("请输入有效的页码范围格式，例如：all，0 或 1-3,5", "格式错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                editBox1.Text = "all";
            }
        }

        private bool ValidatePageInput(string input)
        {
            return Regex.IsMatch(input, @"^(\d+(-\d+)?)(,\s*\d+(-\d+)?)*$");
        }

        private List<int> ParsePageRange(string input, int maxPages)
        {
            var pages = new HashSet<int>();

            if (string.IsNullOrWhiteSpace(input))
                return new List<int>();

            input = input.Trim().ToLower();

            if (input == "all")
            {
                for (int i = 1; i <= maxPages; i++)
                    pages.Add(i);
                return pages.OrderBy(p => p).ToList();
            }

            if (input == "0")
            {
                try
                {
                    int currentSlideIndex = app.ActiveWindow.View.Slide.SlideIndex;
                    if (currentSlideIndex >= 1 && currentSlideIndex <= maxPages)
                        return new List<int> { currentSlideIndex };
                    else
                        return new List<int>();
                }
                catch
                {
                    return new List<int>();
                }
            }

            if (!ValidatePageInput(input))
                return new List<int>();

            foreach (var part in input.Split(','))
            {
                var range = part.Trim();
                if (string.IsNullOrEmpty(range))
                    continue;

                if (range.Contains("-"))
                {
                    var bounds = range.Split('-');
                    if (bounds.Length != 2)
                        continue;

                    if (int.TryParse(bounds[0], out int start) && int.TryParse(bounds[1], out int end))
                    {
                        start = Math.Max(1, start);
                        end = Math.Min(maxPages, end);
                        if (start > end)
                            (start, end) = (end, start);

                        for (int i = start; i <= end; i++)
                            pages.Add(i);
                    }
                }
                else
                {
                    if (int.TryParse(range, out int page) && page >= 1 && page <= maxPages)
                        pages.Add(page);
                }
            }

            return pages.OrderBy(p => p).ToList();
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrEmpty(saveFolderPath))
            {
                MessageBox.Show("请先选择保存路径！", "路径未设置",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var pres = app.ActivePresentation;
                int totalSlides = pres.Slides.Count;

                selectedPages = ParsePageRange(editBox1.Text, totalSlides);

                if (selectedPages.Count == 0)
                {
                    MessageBox.Show("没有有效的幻灯片被选择", "导出中止",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int exportedCount = 0;
                foreach (int slideNumber in selectedPages)
                {
                    if (ExportSingleSlide(pres, slideNumber))
                        exportedCount++;
                }

                MessageBox.Show($"成功导出 {exportedCount}/{selectedPages.Count} 张幻灯片", "导出完成",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "严重错误",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 核心功能

        private bool ExportSingleSlide(Presentation pres, int slideNumber)
        {
            try
            {
                if (slideNumber < 1 || slideNumber > pres.Slides.Count)
                    return false;

                var slide = pres.Slides[slideNumber];
                int baseWidth = (int)pres.PageSetup.SlideWidth;
                int baseHeight = (int)pres.PageSetup.SlideHeight;

                string fileName = $"{Path.GetFileNameWithoutExtension(pres.Name)}_Slide{slideNumber}_{DateTime.Now:yyyyMMddHHmmss}.{exportFormat}";
                string fullPath = Path.Combine(saveFolderPath, fileName);

                int outputWidth = (int)(baseWidth * currentDPI / 72.0);
                int outputHeight = (int)(baseHeight * currentDPI / 72.0);

                slide.Export(fullPath, exportFormat.ToUpper(), outputWidth, outputHeight);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"幻灯片 {slideNumber} 导出失败：{ex.Message}", "导出错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        #endregion

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            string message = "Developers: Shikun, CSU\nEmail: shikun.creative@gmail.com";
            MessageBox.Show(message, "关于开发者", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "https://github.com/ShikunSun/PowerPointVSTOPlugins",
                UseShellExecute = true
            });
        }
    }
}
