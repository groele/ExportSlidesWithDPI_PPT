using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
/// <summary>
/// Export the specified slide as an image, and support setting the format and DPI
/// </summary>
/// V 6.0  2026.07.18
/// 中文版本
namespace ExportSlidesWithDPIDoing
{
    public partial class Ribbon1
    {
        private PowerPoint.Application app;
        private int currentDPI = 300;
        private string exportFormat = "jpg";
        private string saveFolderPath = string.Empty;
        private List<int> selectedPages = new List<int>();
        private bool isExporting = false;
        private ProgressForm progressForm;
        private bool enableCropWhiteSpace = false;
        private int whiteSpaceMargin = 0; // 四周留白大小（像素）
        private string selectedExportFileName = string.Empty;
        private readonly Dictionary<string, string> imageFormatMap = new Dictionary<string, string>
        {
            { "jpg", "JPG" },
            { "png", "PNG" },
            { "bmp", "BMP" },
            { "tif", "TIF" }
        };

        #region Ribbon 初始化
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                app.PresentationBeforeClose += App_PresentationBeforeClose;
                InitializeControls();
                BindEvents();

                editBox1.Label = "页码";
                editBox1.Text = "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    text: $"初始化失败：{ex.Message}",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
            }
        }

        private void App_PresentationBeforeClose(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            try
            {
                if (isExporting)
                {
                    Cancel = true;
                    MessageBox.Show(
                        text: "正在导出幻灯片，请等待导出完成后再关闭。",
                        caption: "导出中",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Warning
                    );
                    return;
                }

                // 清理资源
                if (progressForm != null && !progressForm.IsDisposed)
                {
                    try
                    {
                        progressForm.Close();
                        progressForm.Dispose();
                    }
                    catch (Exception)
                    {
                        // 忽略清理进度窗体时的错误
                    }
                }

                // 重置状态
                isExporting = false;
                selectedPages.Clear();
                saveFolderPath = string.Empty;
            }
            catch (Exception ex)
            {
                // 记录错误但不阻止关闭
                System.Diagnostics.Debug.WriteLine($"关闭时发生错误：{ex.Message}");
            }
        }

        void InitializeControls()
        {
            try
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
                foreach (var format in imageFormatMap.Keys.Concat(new[] { "pdf" }))
                {
                    var item = Factory.CreateRibbonDropDownItem();
                    item.Label = format.ToUpper();
                    comboBox2.Items.Add(item);
                }
                comboBox2.Text = exportFormat.ToUpper();

                // 初始化裁剪白边相关控件
                if (checkBox1 != null)
                {
                    checkBox1.Label = "裁剪白边";
                    checkBox1.Checked = enableCropWhiteSpace;
                }

                if (editBox2 != null)
                {
                    editBox2.Label = "留白大小";
                    editBox2.Text = whiteSpaceMargin.ToString();
                    editBox2.Enabled = enableCropWhiteSpace;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    text: $"控件初始化失败：{ex.Message}",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
            }
        }

        void BindEvents()
        {
            this.button2.Click += Button2_Click;       // 导出按钮
            this.comboBox1.TextChanged += ComboBox1_TextChanged; // DPI修改
            this.comboBox2.TextChanged += ComboBox2_TextChanged; // 格式修改
            this.editBox1.TextChanged += EditBox_TextChanged;    // 页码输入
            this.button3.Click += button3_Click;       // 关于开发者
            this.button4.Click += button4_Click_1;     // 打开网址
            // 移除重复的事件绑定，因为这些事件已经在Designer文件中绑定了
            // this.button5.Click += button5_Click_1;     // 另存为按钮
            // this.button6.Click += button6_Click_1;     // 图片另存为按钮

            // 绑定裁剪白边相关事件
            if (checkBox1 != null)
            {
                this.checkBox1.Click += CheckBox1_Click;   // 裁剪白边复选框
            }
            if (editBox2 != null)
            {
                this.editBox2.TextChanged += EditBox2_TextChanged;   // 留白大小输入框
            }
        }
        #endregion
          
        #region 事件处理

        private DialogResult ShowMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            try
            {
                DialogResult result = DialogResult.None;
                using (var form = new Form())
                {
                    form.TopMost = true;  // 设置窗体始终显示在最顶层
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.None;
                    form.ShowInTaskbar = false;
                    form.Size = new System.Drawing.Size(0, 0);
                    form.Load += (s, e) =>
                    {
                        // 确保窗体在最顶层
                        form.Activate();
                        form.BringToFront();
                        
                        result = MessageBox.Show(
                            text: text,
                            caption: caption,
                            buttons: buttons,
                            icon: icon,
                            defaultButton: MessageBoxDefaultButton.Button1
                        );
                        form.Close();
                    };
                    form.ShowDialog();
                }
                return result;
            }
            catch (Exception)
            {
                // 如果自定义显示失败，回退到普通显示
                return MessageBox.Show(
                    text: text,
                    caption: caption,
                    buttons: buttons,
                    icon: icon,
                    defaultButton: MessageBoxDefaultButton.Button1
                );
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
                ShowMessageBox(
                    text: "无效的DPI值，请输入正整数",
                    caption: "输入错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Warning
                );
                comboBox1.Text = currentDPI.ToString();
            }
        }

        private void ComboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var format = comboBox2.Text.Trim().ToLowerInvariant();
                if (imageFormatMap.ContainsKey(format) || format == "pdf")
                {
                    exportFormat = format;
                    comboBox1.Enabled = true; // PDF 裁剪模式也需要 DPI
                }
                else
                {
                    ShowMessageBox(
                    text: "仅支持 PDF、PNG、JPG、BMP、TIF 格式",
                        caption: "格式错误",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Warning
                    );
                    comboBox2.Text = exportFormat.ToUpper();
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    text: $"格式设置失败：{ex.Message}",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
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
                ShowMessageBox(
                    text: "请输入有效的页码范围格式，例如：all，0 或 1-3,5",
                    caption: "格式错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Warning
                );
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
            if (isExporting)
            {
                ShowMessageBox("正在导出中，请等待当前任务完成", "导出中", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                EnsureOutputFolderIsWritable();
                var pres = app.ActivePresentation ?? throw new InvalidOperationException("无法访问当前演示文稿");
                selectedPages = ParsePageRange(editBox1.Text, pres.Slides.Count);
                if (selectedPages.Count == 0)
                    throw new InvalidOperationException("没有有效的幻灯片被选择");
                if (!string.IsNullOrEmpty(selectedExportFileName) && selectedPages.Count != 1 && exportFormat != "pdf")
                    throw new InvalidOperationException("“另存为”一次只能导出一张图片；请改用“文件夹”导出多张图片。");

                isExporting = true;
                SetExportControlsEnabled(false);
                if (exportFormat == "pdf")
                {
                    ExportPdf(pres, selectedPages);
                    return;
                }

                progressForm = new ProgressForm { TotalSlides = selectedPages.Count };
                progressForm.Show();
                var failedSlides = new List<int>();
                foreach (int slideNumber in selectedPages)
                {
                    try
                    {
                        ExportImageSlide(pres, slideNumber);
                        progressForm.UpdateProgress();
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch (Exception ex)
                    {
                        failedSlides.Add(slideNumber);
                        Debug.WriteLine($"幻灯片 {slideNumber} 导出失败：{ex}");
                    }
                }

                if (failedSlides.Count > 0)
                {
                    ShowMessageBox(
                        $"成功导出 {selectedPages.Count - failedSlides.Count}/{selectedPages.Count} 张图片。\n失败页码：{string.Join(", ", failedSlides)}",
                        "部分导出失败",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox($"导出失败：{ex.Message}", "导出错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                isExporting = false;
                SetExportControlsEnabled(true);
                if (progressForm != null && !progressForm.IsDisposed)
                    progressForm.Close();
                progressForm = null;
            }
        }

        private void SetExportControlsEnabled(bool enabled)
        {
            button2.Enabled = enabled;
            button5.Enabled = enabled;
            button6.Enabled = enabled;
            comboBox1.Enabled = enabled;
            comboBox2.Enabled = enabled;
            editBox1.Enabled = enabled;
            checkBox1.Enabled = enabled;
            editBox2.Enabled = enabled && enableCropWhiteSpace;
            button3.Enabled = enabled;
            button4.Enabled = enabled;
        }

        private void EnsureOutputFolderIsWritable()
        {
            if (string.IsNullOrWhiteSpace(saveFolderPath))
                throw new InvalidOperationException("请先选择保存路径。");
            Directory.CreateDirectory(saveFolderPath);
            string root = Path.GetPathRoot(saveFolderPath);
            if (!string.IsNullOrEmpty(root) && new DriveInfo(root).AvailableFreeSpace < 100L * 1024 * 1024)
                throw new IOException("保存磁盘可用空间不足 100 MB。");
            string testFile = Path.Combine(saveFolderPath, ".exportslides-write-test-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllText(testFile, "test");
            }
            finally
            {
                if (File.Exists(testFile)) File.Delete(testFile);
            }
        }

        private void ExportImageSlide(PowerPoint.Presentation pres, int slideNumber)
        {
            var slide = pres.Slides[slideNumber];
            string extension = exportFormat;
            string fileName = ConsumeSelectedFileName() ??
                $"{Path.GetFileNameWithoutExtension(pres.Name)}_Slide{slideNumber}_{DateTime.Now:yyyyMMddHHmmssfff}.{extension}";
            string fullPath = Path.Combine(saveFolderPath, fileName);
            int outputWidth = (int)(pres.PageSetup.SlideWidth * currentDPI / 72.0);
            int outputHeight = (int)(pres.PageSetup.SlideHeight * currentDPI / 72.0);
            string stagingPath = CreateStagingPath(fullPath);

            try
            {
                using (var tempFile = new TempFile(extension))
                {
                    // PowerPoint COM must remain on the Office UI thread. Do not wrap this in Task.Run.
                    slide.Export(tempFile.Path, imageFormatMap[extension], outputWidth, outputHeight);
                    if (!enableCropWhiteSpace)
                    {
                        File.Copy(tempFile.Path, stagingPath, false);
                    }
                    else
                    {
                        using (var image = System.Drawing.Image.FromFile(tempFile.Path))
                        using (var croppedImage = CropWhiteSpace(image, whiteSpaceMargin))
                        {
                            var encoder = GetEncoder(extension);
                            var encoderParams = GetEncoderParameters(extension);
                            try
                            {
                                if (encoder != null && encoderParams != null)
                                    croppedImage.Save(stagingPath, encoder, encoderParams);
                                else
                                    croppedImage.Save(stagingPath, GetImageFormat(extension));
                            }
                            finally
                            {
                                encoderParams?.Dispose();
                            }
                        }
                    }
                }

                VerifyOutputFile(stagingPath);
                CommitOutput(stagingPath, fullPath);
            }
            finally
            {
                if (File.Exists(stagingPath)) File.Delete(stagingPath);
            }
        }

        private void ExportPdf(PowerPoint.Presentation pres, IList<int> pages)
        {
            string fileName = ConsumeSelectedFileName() ??
                $"{Path.GetFileNameWithoutExtension(pres.Name)}_{DateTime.Now:yyyyMMddHHmmssfff}.pdf";
            fileName = Path.ChangeExtension(fileName, ".pdf");
            string fullPath = Path.Combine(saveFolderPath, fileName);
            string stagingPath = CreateStagingPath(fullPath);

            try
            {
                var contiguousRanges = ToContiguousRanges(pages).ToList();
                if (!enableCropWhiteSpace && contiguousRanges.Count == 1)
                {
                    ExportVectorPdf(pres, contiguousRanges[0], stagingPath);
                }
                else
                {
                    // PowerPoint only accepts one continuous PrintRange for a single PDF export.
                    // Keep the user's non-contiguous page selection by using the raster-PDF path.
                    progressForm = new ProgressForm { TotalSlides = pages.Count };
                    progressForm.Show();
                    int outputWidth = (int)(pres.PageSetup.SlideWidth * currentDPI / 72.0);
                    int outputHeight = (int)(pres.PageSetup.SlideHeight * currentDPI / 72.0);
                    using (var pdf = new RasterPdfDocument(stagingPath, pages.Count))
                    {
                        foreach (int slideNumber in pages)
                        {
                            using (var tempFile = new TempFile("png"))
                            {
                                pres.Slides[slideNumber].Export(tempFile.Path, "PNG", outputWidth, outputHeight);
                                using (var image = System.Drawing.Image.FromFile(tempFile.Path))
                                using (var croppedImage = enableCropWhiteSpace
                                    ? CropWhiteSpace(image, whiteSpaceMargin)
                                    : new System.Drawing.Bitmap(image))
                                {
                                    pdf.AddPage(croppedImage, currentDPI);
                                }
                            }
                            progressForm.UpdateProgress();
                            System.Windows.Forms.Application.DoEvents();
                        }
                        pdf.Complete();
                    }
                }

                VerifyOutputFile(stagingPath);
                CommitOutput(stagingPath, fullPath);
            }
            finally
            {
                if (File.Exists(stagingPath)) File.Delete(stagingPath);
            }
        }

        private void ExportVectorPdf(PowerPoint.Presentation pres, Tuple<int, int> pageRange, string outputPath)
        {
            PowerPoint.PrintRanges ranges = pres.PrintOptions.Ranges;
            var savedRanges = new List<Tuple<int, int>>();
            for (int i = 1; i <= ranges.Count; i++)
            {
                PowerPoint.PrintRange saved = ranges[i];
                savedRanges.Add(Tuple.Create(saved.Start, saved.End));
            }
            ranges.ClearAll();
            PowerPoint.PrintRange range = ranges.Add(pageRange.Item1, pageRange.Item2);

            try
            {
                pres.ExportAsFixedFormat(
                    outputPath,
                    PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                    PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    Office.MsoTriState.msoFalse,
                    PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                    Office.MsoTriState.msoFalse,
                    range,
                    PowerPoint.PpPrintRangeType.ppPrintSlideRange,
                    string.Empty,
                    true,
                    true,
                    true,
                    false,
                    false,
                    Type.Missing);
            }
            finally
            {
                ranges.ClearAll();
                foreach (var savedRange in savedRanges)
                    ranges.Add(savedRange.Item1, savedRange.Item2);
            }
        }

        private static IEnumerable<Tuple<int, int>> ToContiguousRanges(IList<int> pages)
        {
            int start = pages[0];
            int end = start;
            for (int i = 1; i < pages.Count; i++)
            {
                if (pages[i] == end + 1)
                {
                    end = pages[i];
                    continue;
                }
                yield return Tuple.Create(start, end);
                start = end = pages[i];
            }
            yield return Tuple.Create(start, end);
        }

        private string ConsumeSelectedFileName()
        {
            if (string.IsNullOrEmpty(selectedExportFileName)) return null;
            string fileName = selectedExportFileName;
            selectedExportFileName = string.Empty;
            return fileName;
        }

        private static void VerifyOutputFile(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
                throw new IOException("导出文件未创建或内容为空。");
        }

        private static string CreateStagingPath(string outputPath)
        {
            string directory = Path.GetDirectoryName(outputPath);
            string extension = Path.GetExtension(outputPath);
            string baseName = Path.GetFileNameWithoutExtension(outputPath);
            return Path.Combine(directory, baseName + "." + Guid.NewGuid().ToString("N") + ".part" + extension);
        }

        private static void CommitOutput(string stagingPath, string outputPath)
        {
            if (File.Exists(outputPath))
            {
                try
                {
                    File.Replace(stagingPath, outputPath, null, true);
                    return;
                }
                catch (PlatformNotSupportedException) { }
                catch (IOException) { }
            }
            File.Move(stagingPath, outputPath);
        }

        private System.Drawing.Image CropWhiteSpace(System.Drawing.Image image, int margin)
        {
            try
            {
                // Convert once to a predictable BGRA layout, then scan the raw buffer.
                // GetPixel incurs a GDI+ call per pixel and is prohibitively slow at 600 DPI.
                using (var bitmap = new System.Drawing.Bitmap(
                    image.Width,
                    image.Height,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    int width = bitmap.Width;
                    int height = bitmap.Height;

                    using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                    {
                        graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                        graphics.DrawImageUnscaled(image, 0, 0);
                    }

                    // 找到非白色区域的边界
                    int left = width;
                    int top = height;
                    int right = 0;
                    int bottom = 0;

                    var rectangle = new System.Drawing.Rectangle(0, 0, width, height);
                    var data = bitmap.LockBits(
                        rectangle,
                        System.Drawing.Imaging.ImageLockMode.ReadOnly,
                        System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    try
                    {
                        int stride = Math.Abs(data.Stride);
                        // A full 600-DPI slide can require more than 140 MB here. Reuse one row buffer.
                        var pixels = new byte[stride];

                        for (int y = 0; y < height; y++)
                        {
                            var rowPointer = new IntPtr(data.Scan0.ToInt64() + (long)y * data.Stride);
                            System.Runtime.InteropServices.Marshal.Copy(rowPointer, pixels, 0, stride);
                            for (int x = 0; x < width; x++)
                            {
                                int offset = x * 4;
                                if (!IsWhitePixel(pixels[offset], pixels[offset + 1], pixels[offset + 2], pixels[offset + 3]))
                                {
                                    left = Math.Min(left, x);
                                    top = Math.Min(top, y);
                                    right = Math.Max(right, x);
                                    bottom = Math.Max(bottom, y);
                                }
                            }
                        }
                    }
                    finally
                    {
                        bitmap.UnlockBits(data);
                    }

                    // 添加边距
                    left = Math.Max(0, left - margin);
                    top = Math.Max(0, top - margin);
                    right = Math.Min(width - 1, right + margin);
                    bottom = Math.Min(height - 1, bottom + margin);

                    // 整页为白色时保留原始尺寸；单像素内容也必须能被正确保留。
                    if (left == width)
                    {
                        var unchanged = new System.Drawing.Bitmap(bitmap);
                        unchanged.SetResolution(image.HorizontalResolution, image.VerticalResolution);
                        return unchanged;
                    }

                    // 创建裁剪后的图片
                    int newWidth = right - left + 1;
                    int newHeight = bottom - top + 1;
                    var croppedBitmap = new System.Drawing.Bitmap(newWidth, newHeight);

                    using (var graphics = System.Drawing.Graphics.FromImage(croppedBitmap))
                    {
                        graphics.DrawImage(bitmap, 
                            new System.Drawing.Rectangle(0, 0, newWidth, newHeight),
                            new System.Drawing.Rectangle(left, top, newWidth, newHeight),
                            System.Drawing.GraphicsUnit.Pixel);
                    }

                    // 设置DPI
                    croppedBitmap.SetResolution(image.HorizontalResolution, image.VerticalResolution);

                    return croppedBitmap;
                }
            }
            catch (Exception ex)
            {
                // 裁剪失败不应丢失已成功渲染的页面，降级为保留原图。
                Debug.WriteLine($"裁剪白边失败，保留原图：{ex}");
                return new System.Drawing.Bitmap(image);
            }
        }

        private static bool IsWhitePixel(byte blue, byte green, byte red, byte alpha)
        {
            // 透明像素与接近白色的像素都视为可裁切的背景。
            const int tolerance = 10;
            return alpha <= tolerance ||
                   (red >= 255 - tolerance &&
                    green >= 255 - tolerance &&
                    blue >= 255 - tolerance);
        }

        private System.Drawing.Imaging.ImageFormat GetImageFormat(string format)
        {
            switch (format.ToLower())
            {
                case "png":
                    return System.Drawing.Imaging.ImageFormat.Png;
                case "jpg":
                    return System.Drawing.Imaging.ImageFormat.Jpeg;
                case "bmp":
                    return System.Drawing.Imaging.ImageFormat.Bmp;
                case "tif":
                    return System.Drawing.Imaging.ImageFormat.Tiff;
                default:
                    return System.Drawing.Imaging.ImageFormat.Png;
            }
        }

        private System.Drawing.Imaging.EncoderParameters GetEncoderParameters(string format)
        {
            if (format.ToLower() == "jpg")
            {
                var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality,
                    100L
                );
                return encoderParams;
            }
            if (format.ToLower() == "tif")
            {
                var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Compression,
                    (long)System.Drawing.Imaging.EncoderValue.CompressionLZW
                );
                return encoderParams;
            }
            return null;
        }

        private System.Drawing.Imaging.ImageCodecInfo GetEncoder(string format)
        {
            var codecs = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
            string mimeType;
            switch (format.ToLower())
            {
                case "jpg":
                    mimeType = "image/jpeg";
                    break;
                case "png":
                    mimeType = "image/png";
                    break;
                case "bmp":
                    mimeType = "image/bmp";
                    break;
                case "tif":
                    mimeType = "image/tiff";
                    break;
                default:
                    mimeType = "image/png";
                    break;
            }
            return codecs.FirstOrDefault(codec => codec.MimeType == mimeType);
        }

        private class TempFile : IDisposable
        {
            public string Path { get; }
            public TempFile(string extension)
            {
                Path = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"temp_{Guid.NewGuid()}.{extension}"
                );
            }

            public void Dispose()
            {
                try
                {
                    if (File.Exists(Path))
                    {
                        File.Delete(Path);
                    }
                }
                catch
                {
                    // 忽略清理临时文件时的错误
                }
            }
        }
        #endregion

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DialogResult result = ShowMessageBox(
                    text: "PPT 导出工具\n\n" +
                    "开发者：Shikun\n" +
                    "版本：6.0.0\n" +
                    "联系方式：shikun.creative@gmail.com\n\n" +
                    "是否访问开发者主页？",
                    caption: "开发者信息",
                    buttons: MessageBoxButtons.YesNo,
                    icon: MessageBoxIcon.Information
                );

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start("https://groele.github.io/SK-Creative.github.io/");
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    text: $"打开开发者主页失败：{ex.Message}",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
            }
        }

        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {
            ShowMessageBox(
                text: "PPT 导出工具：\n\n" +
                "基本功能：\n" +
                "   - 支持 PDF、JPG、PNG、BMP、TIF\n" +
                "   - 灵活的 DPI 设置\n" +
                "   - 灵活的导出范围选择\n" +
                "   - 可裁剪图片与 PDF 的白边\n\n" +
                "PDF 说明：\n" +
                "   - 未勾选裁剪：使用 PowerPoint 原生矢量 PDF。\n" +
                "   - 勾选裁剪：按所选 DPI 裁切后生成 PDF。\n" +
                "   - 非连续页码导出时，为保留页码选择，会生成高分辨率图像 PDF。\n\n",
                caption: "关于",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Information
            );
        }

        private void CheckBox1_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBox1 != null)
            {
                enableCropWhiteSpace = checkBox1.Checked;
                if (editBox2 != null)
                {
                    editBox2.Enabled = enableCropWhiteSpace;
                }
            }
        }

        private void EditBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (editBox2 != null)
            {
                if (int.TryParse(editBox2.Text, out int margin) && margin >= 0)
                {
                    whiteSpaceMargin = margin;
                }
                else
                {
                    ShowMessageBox(
                        text: "请输入有效的留白大小（非负整数）",
                        caption: "输入错误",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Warning
                    );
                    editBox2.Text = whiteSpaceMargin.ToString();
                }
            }
        }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string initialFolder = !string.IsNullOrEmpty(saveFolderPath) && Directory.Exists(saveFolderPath)
                    ? saveFolderPath
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string selectedFolder = SelectFolderWithPowerPoint(initialFolder);
                if (!string.IsNullOrEmpty(selectedFolder))
                {
                    saveFolderPath = selectedFolder;
                    EnsureOutputFolderIsWritable();
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    text: $"选择路径时发生错误：{ex.Message}\n请重试",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
            }
        }

        private string SelectFolderWithPowerPoint(string initialFolder)
        {
            object dialogObject = null;
            try
            {
                // PowerPoint's FileDialog matches the Office-native file/folder picker UI.
                dynamic powerPoint = app;
                dialogObject = powerPoint.FileDialog(Office.MsoFileDialogType.msoFileDialogFolderPicker);
                dynamic dialog = dialogObject;
                dialog.Title = "选择导出文件夹";
                dialog.ButtonName = "选择文件夹";
                dialog.InitialFileName = initialFolder.EndsWith(Path.DirectorySeparatorChar.ToString())
                    ? initialFolder
                    : initialFolder + Path.DirectorySeparatorChar;
                if (dialog.Show() != -1) return null;

                dynamic selectedItems = dialog.SelectedItems;
                return Convert.ToString(selectedItems[1]);
            }
            finally
            {
                ReleaseComObject(dialogObject);
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
        }

        private void button6_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (isExporting)
            {
                ShowMessageBox(
                    text: "正在导出中，请等待当前任务完成",
                    caption: "导出中",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Warning
                );
                return;
            }

            try
            {
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Title = "选择导出文件保存位置";
                    saveDialog.Filter = "PDF 文件|*.pdf|JPEG 图片|*.jpg|PNG 图片|*.png|BMP 图片|*.bmp|TIFF 图片|*.tif|所有文件|*.*";
                    // 根据当前导出格式设置默认文件名和扩展名
                    string ext = exportFormat.ToLower();
                    string defaultName = "Fig." + ext;
                    saveDialog.FileName = defaultName;
                    // 设置FilterIndex与格式对应
                    switch (ext)
                    {
                        case "pdf": saveDialog.FilterIndex = 1; break;
                        case "jpg": saveDialog.FilterIndex = 2; break;
                        case "png": saveDialog.FilterIndex = 3; break;
                        case "bmp": saveDialog.FilterIndex = 4; break;
                        case "tif": saveDialog.FilterIndex = 5; break;
                        default: saveDialog.FilterIndex = 1; break;
                    }
                    saveDialog.InitialDirectory = !string.IsNullOrEmpty(saveFolderPath) && Directory.Exists(saveFolderPath) 
                        ? saveFolderPath 
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedPath = Path.GetDirectoryName(saveDialog.FileName);
                        string selectedFileName = Path.GetFileName(saveDialog.FileName);
                        
                        // 检查路径是否有效
                        if (string.IsNullOrEmpty(selectedPath))
                        {
                            ShowMessageBox(
                                text: "请选择有效的保存路径",
                                caption: "路径无效",
                                buttons: MessageBoxButtons.OK,
                                icon: MessageBoxIcon.Warning
                            );
                            return;
                        }

                        // 检查文件是否可写
                        try
                        {
                            string fullPath = Path.Combine(selectedPath, selectedFileName);
                            if (File.Exists(fullPath))
                            {
                                // 尝试打开文件以检查是否可写
                                using (FileStream fs = File.Open(fullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                                {
                                    // 文件未被占用，可以继续
                                }
                                // 确保文件没有被设置为只读
                                FileAttributes attributes = File.GetAttributes(fullPath);
                                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                                {
                                    ShowMessageBox(
                                        text: "所选文件是只读的，无法替换",
                                        caption: "文件只读",
                                        buttons: MessageBoxButtons.OK,
                                        icon: MessageBoxIcon.Error
                                    );
                                    return;
                                }
                            }
                            else
                            {
                                // 如果文件不存在，检查目录权限
                                try
                                {
                                    string testFile = Path.Combine(selectedPath, "test_write.tmp");
                                    File.WriteAllText(testFile, "test");
                                    File.Delete(testFile);
                                }
                                catch (UnauthorizedAccessException)
                                {
                                    ShowMessageBox(
                                        text: "没有权限在所选目录创建文件",
                                        caption: "权限不足",
                                        buttons: MessageBoxButtons.OK,
                                        icon: MessageBoxIcon.Error
                                    );
                                    return;
                                }
                            }
                        }
                        catch (IOException)
                        {
                            ShowMessageBox(
                                text: $"所选文件正在被其他程序使用，无法替换",
                                caption: "文件被占用",
                                buttons: MessageBoxButtons.OK,
                                icon: MessageBoxIcon.Error
                            );
                            return;
                        }
                        catch (Exception ex)
                        {
                            ShowMessageBox(
                                text: $"检查文件时发生错误：{ex.Message}\n请选择其他文件或路径",
                                caption: "错误",
                                buttons: MessageBoxButtons.OK,
                                icon: MessageBoxIcon.Error
                            );
                            return;
                        }

                        // 检查磁盘空间
                        try
                        {
                            var drive = new DriveInfo(Path.GetPathRoot(selectedPath));
                            if (drive.AvailableFreeSpace < 1024 * 1024 * 100) // 100MB
                            {
                                ShowMessageBox(
                                    text: "所选磁盘空间不足，请确保有至少100MB的可用空间",
                                    caption: "空间不足",
                                    buttons: MessageBoxButtons.OK,
                                    icon: MessageBoxIcon.Warning
                                );
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            ShowMessageBox(
                                text: $"检查磁盘空间时出错：{ex.Message}\n请选择其他路径",
                                caption: "磁盘错误",
                                buttons: MessageBoxButtons.OK,
                                icon: MessageBoxIcon.Error
                            );
                            return;
                        }

                        // 所有检查都通过，保存完整路径和文件名
                        try
                        {
                            saveFolderPath = selectedPath;
                            selectedExportFileName = selectedFileName;
                            
                            // 设置导出格式为选择的文件格式
                            string extension = Path.GetExtension(selectedFileName).TrimStart('.').ToLower();
                            if (imageFormatMap.ContainsKey(extension) || extension == "pdf")
                            {
                                exportFormat = extension;
                                comboBox2.Text = extension.ToUpper();
                            }
                            else
                            {
                                ShowMessageBox(
                                    text: "不支持的文件格式，请选择 PDF、JPG、PNG、BMP 或 TIF 格式",
                                    caption: "格式错误",
                                    buttons: MessageBoxButtons.OK,
                                    icon: MessageBoxIcon.Warning
                                );
                                return;
                            }
                            
                            // 调用导出功能
                            Button2_Click(sender, e);
                        }
                        catch (Exception ex)
                        {
                            ShowMessageBox(
                                text: $"设置导出参数时发生错误：{ex.Message}\n请重试",
                                caption: "错误",
                                buttons: MessageBoxButtons.OK,
                                icon: MessageBoxIcon.Error
                            );
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    text: $"选择路径时发生错误：{ex.Message}\n请重试",
                    caption: "错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
            }
        }
    }

    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label statusLabel;
        public int TotalSlides { get; set; }
        private int currentProgress = 0;
        public ProgressForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "导出进度";
            this.Size = new System.Drawing.Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;  // 设置窗口始终显示在最顶层

            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(350, 30),
                Maximum = 100
            };

            statusLabel = new Label
            {
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(350, 20),
                Text = "准备导出..."
            };

            this.Controls.Add(progressBar);
            this.Controls.Add(statusLabel);
        }

        public void UpdateProgress()
        {
            if (!IsDisposed && TotalSlides > 0)
            {
                currentProgress++;
                int percentage = Math.Min(100, (int)((float)currentProgress / TotalSlides * 100));
                progressBar.Value = percentage;
                statusLabel.Text = $"正在导出... {currentProgress}/{TotalSlides} ({percentage}%)";
            }
        }
    }
}
