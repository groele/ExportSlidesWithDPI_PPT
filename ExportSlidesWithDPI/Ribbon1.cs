using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
/// <summary>
/// Export the specified slide as an image, and support setting the format and DPI
/// </summary>
/// V 2.0  2025.05.28
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
        private readonly Dictionary<string, string> formatMap = new Dictionary<string, string>
        {
            { "jpg", "JPG" },
            { "png", "PNG" },
            { "bmp", "BMP" },
            { "tif", "TIF" }
        };

        // 性能优化相关参数
        private const int BATCH_SIZE = 10; // 增加批处理大小
        private const int MEMORY_THRESHOLD = 85; // 内存使用率阈值（百分比）
        private const int SPEED_LIMIT = 500; // 减少速度限制
        private const int MAX_RETRY_COUNT = 3; // 最大重试次数
        private const int RETRY_DELAY = 1000; // 减少重试延迟
        private const int MAX_PARALLEL_TASKS = 4; // 最大并行任务数
        private CancellationTokenSource cancellationTokenSource;
        private SemaphoreSlim semaphore;

        #region Ribbon 初始化
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                app.PresentationBeforeClose += App_PresentationBeforeClose;
                InitializeControls();
                BindEvents();

                editBox1.Label = "Page";
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

                // 安全地取消和释放 CancellationTokenSource
                if (cancellationTokenSource != null)
                {
                    try
                    {
                        if (!cancellationTokenSource.IsCancellationRequested)
                        {
                            cancellationTokenSource.Cancel();
                        }
                    }
                    catch (ObjectDisposedException)
                    {
                        // 忽略已释放的异常
                    }
                    finally
                    {
                        try
                        {
                            cancellationTokenSource.Dispose();
                        }
                        catch (ObjectDisposedException)
                        {
                            // 忽略已释放的异常
                        }
                    }
                }

                // 安全地释放信号量
                if (semaphore != null)
                {
                    try
                    {
                        semaphore.Dispose();
                    }
                    catch (ObjectDisposedException)
                    {
                        // 忽略已释放的异常
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
                foreach (var format in formatMap.Keys)
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
                var format = comboBox2.Text.ToLower();
                if (formatMap.ContainsKey(format))
                {
                    exportFormat = format;
                    comboBox1.Enabled = true;
                }
                else
                {
                    ShowMessageBox(
                        text: "仅支持 PNG/JPG/BMP/TIF 格式",
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

        private async void Button2_Click(object sender, RibbonControlEventArgs e)
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

            if (string.IsNullOrEmpty(saveFolderPath))
            {
                ShowMessageBox(
                    text: "请先选择保存路径！",
                    caption: "路径未设置",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Warning
                );
                return;
            }

            try
            {
                // 检查保存路径的权限
                try
                {
                    string testFile = Path.Combine(saveFolderPath, "test_write.tmp");
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);
                }
                catch (Exception ex)
                {
                    ShowMessageBox(
                        text: $"保存路径没有写入权限：{ex.Message}",
                        caption: "权限错误",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Error
                    );
                    return;
                }

                // 检查磁盘空间
                var drive = new DriveInfo(Path.GetPathRoot(saveFolderPath));
                if (drive.AvailableFreeSpace < 1024 * 1024 * 100) // 100MB
                {
                    ShowMessageBox(
                        text: "磁盘空间不足，请确保有至少100MB的可用空间",
                        caption: "空间不足",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Warning
                    );
                    return;
                }

                isExporting = true;
                cancellationTokenSource = new CancellationTokenSource();
                semaphore = new SemaphoreSlim(MAX_PARALLEL_TASKS); // 初始化信号量

                var pres = app.ActivePresentation;
                if (pres == null)
                {
                    throw new InvalidOperationException("无法访问当前演示文稿");
                }

                int totalSlides = pres.Slides.Count;
                if (totalSlides == 0)
                {
                    throw new InvalidOperationException("当前演示文稿没有幻灯片");
                }

                selectedPages = ParsePageRange(editBox1.Text, totalSlides);

                if (selectedPages.Count == 0)
                {
                    ShowMessageBox(
                        text: "没有有效的幻灯片被选择",
                        caption: "导出中止",
                        buttons: MessageBoxButtons.OK,
                        icon: MessageBoxIcon.Warning
                    );
                    return;
                }

                // 创建进度窗体
                progressForm = new ProgressForm();
                progressForm.TotalSlides = selectedPages.Count;
                progressForm.Show();

                int exportedCount = 0;
                var failedSlides = new List<int>();
                var errorMessages = new List<string>();

                // 分批处理幻灯片
                for (int i = 0; i < selectedPages.Count; i += BATCH_SIZE)
                {
                    if (cancellationTokenSource.Token.IsCancellationRequested)
                        break;

                    var batch = selectedPages.Skip(i).Take(BATCH_SIZE).ToList();
                    var tasks = new List<Task<bool>>();

                    foreach (int slideNumber in batch)
                    {
                        if (cancellationTokenSource.Token.IsCancellationRequested)
                            break;

                        // 检查内存使用情况
                        if (GetMemoryUsage() > MEMORY_THRESHOLD)
                        {
                            await Task.Delay(1000).ConfigureAwait(false);
                            if (GetMemoryUsage() > MEMORY_THRESHOLD)
                            {
                                if (!progressForm.IsDisposed)
                                {
                                    progressForm.Invoke((MethodInvoker)delegate
                                    {
                                        ShowMessageBox(
                                            text: "系统内存使用率过高，导出已暂停。请关闭其他程序后重试。",
                                            caption: "内存警告",
                                            buttons: MessageBoxButtons.OK,
                                            icon: MessageBoxIcon.Warning
                                        );
                                    });
                                }
                                cancellationTokenSource.Cancel();
                                break;
                            }
                        }

                        tasks.Add(ExportSlideAsync(pres, slideNumber));
                    }

                    try
                    {
                        var results = await Task.WhenAll(tasks).ConfigureAwait(false);
                        for (int j = 0; j < results.Length; j++)
                        {
                            if (results[j])
                                exportedCount++;
                            else
                                failedSlides.Add(batch[j]);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessages.Add($"批处理 {i / BATCH_SIZE + 1} 失败: {ex.Message}");
                    }

                    // 强制进行垃圾回收
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                if (!progressForm.IsDisposed)
                {
                    progressForm.Invoke((MethodInvoker)delegate
                    {
                        progressForm.Close();
                    });
                }

                // 显示导出结果
                string resultMessage = $"成功导出 {exportedCount}/{selectedPages.Count} 张幻灯片";
                if (failedSlides.Count > 0)
                {
                    resultMessage += $"\n\n导出失败的幻灯片：{string.Join(", ", failedSlides)}";
                }
                if (errorMessages.Count > 0)
                {
                    resultMessage += $"\n\n错误信息：\n{string.Join("\n", errorMessages)}";
                }

                ShowMessageBox(
                    text: resultMessage,
                    caption: "导出完成",
                    buttons: MessageBoxButtons.OK,
                    icon: exportedCount == selectedPages.Count ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                );
            }
            catch (Exception ex)
            {
                if (!progressForm.IsDisposed)
                {
                    progressForm.Invoke((MethodInvoker)delegate
                    {
                        ShowMessageBox(
                            text: $"导出失败：{ex.Message}",
                            caption: "严重错误",
                            buttons: MessageBoxButtons.OK,
                            icon: MessageBoxIcon.Error
                        );
                    });
                }
            }
            finally
            {
                isExporting = false;
                if (progressForm != null && !progressForm.IsDisposed)
                {
                    progressForm.Invoke((MethodInvoker)delegate
                    {
                        progressForm.Close();
                    });
                }
                cancellationTokenSource?.Dispose();
                semaphore?.Dispose();
            }
        }

        private double GetMemoryUsage()
        {
            try
            {
                var process = Process.GetCurrentProcess();
                var memoryUsage = process.WorkingSet64 / (1024.0 * 1024.0 * 1024.0); // 转换为GB
                var totalMemory = new PerformanceCounter("Memory", "Available MBytes").NextValue() / 1024.0; // 转换为GB
                return (memoryUsage / (memoryUsage + totalMemory)) * 100;
            }
            catch
            {
                return 0;
            }
        }

        private async Task<bool> ExportSlideAsync(PowerPoint.Presentation pres, int slideNumber)
        {
            int retryCount = 0;
            while (retryCount < MAX_RETRY_COUNT)
            {
                try
                {
                    await semaphore.WaitAsync().ConfigureAwait(false);
                    try
                    {
                        if (pres == null || pres.Slides == null)
                        {
                            throw new InvalidOperationException("演示文稿无效或已关闭");
                        }

                        if (slideNumber < 1 || slideNumber > pres.Slides.Count)
                        {
                            throw new ArgumentOutOfRangeException(nameof(slideNumber), "幻灯片编号超出范围");
                        }

                        var slide = pres.Slides[slideNumber];
                        if (slide == null)
                        {
                            throw new InvalidOperationException("无法访问指定的幻灯片");
                        }

                        string extension = exportFormat.ToLower() == "tif" ? "tif" : exportFormat;
                        string fileName = $"{Path.GetFileNameWithoutExtension(pres.Name)}_Slide{slideNumber}_{DateTime.Now:yyyyMMddHHmmss}.{extension}";
                        string fullPath = Path.Combine(saveFolderPath, fileName);

                        // 检查目标文件夹是否存在且可写
                        if (!Directory.Exists(saveFolderPath))
                        {
                            Directory.CreateDirectory(saveFolderPath);
                        }

                        // 检查文件是否被占用
                        if (File.Exists(fullPath))
                        {
                            try
                            {
                                using (FileStream fs = File.Open(fullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                                {
                                    // 文件未被占用，可以继续
                                }
                            }
                            catch (IOException)
                            {
                                throw new IOException($"文件 {fullPath} 正在被其他程序使用");
                            }
                        }

                        int baseWidth = (int)pres.PageSetup.SlideWidth;
                        int baseHeight = (int)pres.PageSetup.SlideHeight;
                        int outputWidth = (int)(baseWidth * currentDPI / 72.0);
                        int outputHeight = (int)(baseHeight * currentDPI / 72.0);

                        await Task.Run(() =>
                        {
                            using (var tempFile = new TempFile(extension))
                            {
                                try
                                {
                                    slide.Export(tempFile.Path, formatMap[extension], outputWidth, outputHeight);
                                    
                                    // 如果需要裁剪白边
                                    if (enableCropWhiteSpace && (extension == "png" || extension == "jpg" || extension == "bmp" || extension == "tif"))
                                    {
                                        using (var image = System.Drawing.Image.FromFile(tempFile.Path))
                                        {
                                            var croppedImage = CropWhiteSpace(image, whiteSpaceMargin);
                                            if (croppedImage != null)
                                            {
                                                try
                                                {
                                                    var encoder = GetEncoder(extension);
                                                    var encoderParams = GetEncoderParameters(extension);
                                                    if (encoder != null)
                                                    {
                                                        croppedImage.Save(fullPath, encoder, encoderParams);
                                                    }
                                                    else
                                                    {
                                                        croppedImage.Save(fullPath, GetImageFormat(extension));
                                                    }
                                                }
                                                finally
                                                {
                                                    croppedImage.Dispose();
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (File.Exists(fullPath))
                                        {
                                            File.Delete(fullPath);
                                        }
                                        File.Move(tempFile.Path, fullPath);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"导出图片格式失败: {ex.Message}");
                                }
                            }
                        }).ConfigureAwait(false);

                        // 验证文件是否成功创建
                        if (!File.Exists(fullPath))
                        {
                            throw new Exception("文件创建失败");
                        }

                        // 验证文件大小
                        var fileInfo = new FileInfo(fullPath);
                        if (fileInfo.Length == 0)
                        {
                            throw new Exception("导出的文件大小为0");
                        }

                        // 更新进度
                        if (progressForm != null && !progressForm.IsDisposed)
                        {
                            progressForm.UpdateProgress();
                        }
                        return true;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }
                catch (Exception ex)
                {
                    retryCount++;
                    if (retryCount >= MAX_RETRY_COUNT)
                    {
                        await Task.Run(() =>
                        {
                            if (!progressForm.IsDisposed)
                            {
                                progressForm.Invoke((MethodInvoker)delegate
                                {
                                    ShowMessageBox(
                                        text: $"幻灯片 {slideNumber} 导出失败（已重试{retryCount}次）：{ex.Message}",
                                        caption: "导出错误",
                                        buttons: MessageBoxButtons.OK,
                                        icon: MessageBoxIcon.Error
                                    );
                                });
                            }
                        }).ConfigureAwait(false);
                        return false;
                    }
                    await Task.Delay(RETRY_DELAY * retryCount).ConfigureAwait(false);
                }
            }
            return false;
        }

        private System.Drawing.Image CropWhiteSpace(System.Drawing.Image image, int margin)
        {
            try
            {
                using (var bitmap = new System.Drawing.Bitmap(image))
                {
                    int width = bitmap.Width;
                    int height = bitmap.Height;

                    // 找到非白色区域的边界
                    int left = width;
                    int top = height;
                    int right = 0;
                    int bottom = 0;

                    for (int y = 0; y < height; y++)
                    {
                        for (int x = 0; x < width; x++)
                        {
                            var pixel = bitmap.GetPixel(x, y);
                            if (!IsWhitePixel(pixel))
                            {
                                left = Math.Min(left, x);
                                top = Math.Min(top, y);
                                right = Math.Max(right, x);
                                bottom = Math.Max(bottom, y);
                            }
                        }
                    }

                    // 添加边距
                    left = Math.Max(0, left - margin);
                    top = Math.Max(0, top - margin);
                    right = Math.Min(width - 1, right + margin);
                    bottom = Math.Min(height - 1, bottom + margin);

                    // 如果找不到非白色区域，返回原图
                    if (left >= right || top >= bottom)
                    {
                        return null;
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
                ShowMessageBox(
                    text: $"裁剪白边时发生错误：{ex.Message}",
                    caption: "裁剪错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
                return null;
            }
        }

        private bool IsWhitePixel(System.Drawing.Color pixel)
        {
            // 判断像素是否为白色（允许一定的容差）
            const int tolerance = 10;
            return Math.Abs(pixel.R - 255) <= tolerance &&
                   Math.Abs(pixel.G - 255) <= tolerance &&
                   Math.Abs(pixel.B - 255) <= tolerance;
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
            if (format.ToLower() == "tif")
            {
                var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Compression,
                    (long)System.Drawing.Imaging.EncoderValue.CompressionNone
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

        #region 核心功能

        private bool ExportSlide(PowerPoint.Presentation pres, int slideNumber)
        {
            try
            {
                if (slideNumber < 1 || slideNumber > pres.Slides.Count)
                    return false;

                var slide = pres.Slides[slideNumber];
                string extension = exportFormat.ToLower() == "tif" ? "tif" : exportFormat;
                string fileName = $"{Path.GetFileNameWithoutExtension(pres.Name)}_Slide{slideNumber}_{DateTime.Now:yyyyMMddHHmmss}.{extension}";
                string fullPath = Path.Combine(saveFolderPath, fileName);

                int baseWidth = (int)pres.PageSetup.SlideWidth;
                int baseHeight = (int)pres.PageSetup.SlideHeight;
                int outputWidth = (int)(baseWidth * currentDPI / 72.0);
                int outputHeight = (int)(baseHeight * currentDPI / 72.0);

                // 对于高DPI导出，使用临时文件进行优化
                if (currentDPI >= 600)
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.{extension}");
                    try
                    {
                        slide.Export(tempPath, formatMap[extension], outputWidth, outputHeight);
                        if (File.Exists(fullPath))
                        {
                            File.Delete(fullPath);
                        }
                        File.Move(tempPath, fullPath);
                    }
                    finally
                    {
                        if (File.Exists(tempPath))
                        {
                            File.Delete(tempPath);
                        }
                    }
                }
                else
                {
                    slide.Export(fullPath, formatMap[extension], outputWidth, outputHeight);
                }
                return true;
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    text: $"幻灯片 {slideNumber} 导出失败：{ex.Message}",
                    caption: "导出错误",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error
                );
                return false;
            }
        }
        #endregion

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DialogResult result = ShowMessageBox(
                    text: "PPT输出图片\n\n" +
                    "开发者：Shikun\n" +
                    "版本：3.0.0\n" +
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
                text: "PPT输出图片：\n\n" +
                "基本功能：\n" +
                "   - 支持多种导出格式\n" +
                "   - 灵活的 DPI 设置\n" +
                "   - 灵活的导出范围选择\n" +
                "   - 用户友好的界面\n" +
                "   - 裁剪图片白边\n\n",
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
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Title = "选择保存位置";
                    saveDialog.Filter = "所有文件|*.*";
                    saveDialog.FileName = "导出图片";
                    saveDialog.InitialDirectory = !string.IsNullOrEmpty(saveFolderPath) && Directory.Exists(saveFolderPath) 
                        ? saveFolderPath 
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedPath = Path.GetDirectoryName(saveDialog.FileName);
                        
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

                        // 检查路径是否可写
                        try
                        {
                            string testFile = Path.Combine(selectedPath, "test_write.tmp");
                            File.WriteAllText(testFile, "test");
                            File.Delete(testFile);
                        }
                        catch (Exception ex)
                        {
                            ShowMessageBox(
                                text: $"所选路径没有写入权限：{ex.Message}\n请选择其他路径",
                                caption: "权限错误",
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

                        // 所有检查都通过，直接保存路径
                        saveFolderPath = selectedPath;
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
                    saveDialog.Title = "选择图片保存位置";
                    saveDialog.Filter = "JPEG图片|*.jpg|PNG图片|*.png|BMP图片|*.bmp|TIFF图片|*.tif|所有文件|*.*";
                    saveDialog.FilterIndex = 1;  // 设置为JPG格式
                    saveDialog.FileName = "导出图片.jpg";  // 默认文件名添加.jpg后缀
                    saveDialog.InitialDirectory = !string.IsNullOrEmpty(saveFolderPath) && Directory.Exists(saveFolderPath) 
                        ? saveFolderPath 
                        : Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedPath = Path.GetDirectoryName(saveDialog.FileName);
                        
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

                        // 检查路径是否可写
                        try
                        {
                            string testFile = Path.Combine(selectedPath, "test_write.tmp");
                            File.WriteAllText(testFile, "test");
                            File.Delete(testFile);
                        }
                        catch (Exception ex)
                        {
                            ShowMessageBox(
                                text: $"所选路径没有写入权限：{ex.Message}\n请选择其他路径",
                                caption: "权限错误",
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

                        // 所有检查都通过，保存路径并执行导出
                        saveFolderPath = selectedPath;
                        
                        // 调用导出功能
                        Button2_Click(sender, e);
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
        private SynchronizationContext uiContext;

        public ProgressForm()
        {
            InitializeComponents();
            uiContext = SynchronizationContext.Current;
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
            if (uiContext != null)
            {
                uiContext.Post(_ =>
                {
                    if (!this.IsDisposed)
                    {
                        currentProgress++;
                        int percentage = (int)((float)currentProgress / TotalSlides * 100);
                        progressBar.Value = percentage;
                        statusLabel.Text = $"正在导出... {currentProgress}/{TotalSlides} ({percentage}%)";
                    }
                }, null);
            }
        }
    }
}


