using HelixToolkit.Wpf.SharpDX;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using SharpDX;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Text.Json;

namespace SharpDX_PCV
{
    /// <summary>
    /// 点云查看器主窗口 - 负责文件加载和点云可视化
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 私有字段
        private readonly List<Vector3> fullPointCloudData = new List<Vector3>(); // 完整数据
        private readonly List<Vector3> renderPointCloudData = new List<Vector3>(); // 渲染数据
        private readonly List<int> renderPointFileIndices = new List<int>(); // 渲染数据对应的文件索引
        private readonly List<string> loadedFilePaths = new List<string>(); // 已加载文件路径
        private readonly List<PointCloudFileInfo> fileInfoList = new List<PointCloudFileInfo>(); // 文件信息列表
        private PointGeometryModel3D? currentPointCloud;
        private MeshGeometryModel3D? currentSTLMesh;
        private readonly PointCloudFileLoader fileLoader;
        private readonly PointCloudVisualizer visualizer;
        private readonly STLConverter stlConverter;

        // 降采样控制
        private int downsampleLevel = 3;

        // 输出目录选择
        private string? selectedOutputDir;

        // 流式渲染控制
        private CancellationTokenSource? streamingCancellation;
        private const int STREAMING_CHUNK_SIZE = 50000;
        private const int STREAMING_THRESHOLD = 100000;
        private const int FIRST_RENDER_CHUNK_SIZE = 20000;

        // 全局变换参数（解决大坐标值渲染问题）
        private Vector3 globalCenter;
        private double globalScale = 1.0;
        private bool globalTransformCalculated = false;
        #endregion

        #region 构造函数
        public MainWindow()
        {
            InitializeComponent();

            // 初始化组件
            fileLoader = new PointCloudFileLoader();
            visualizer = new PointCloudVisualizer(helixViewport);
            stlConverter = new STLConverter();

            // 确保所有UI控件都已加载后再初始化应用程序
            this.Loaded += MainWindow_Loaded;
        }

        /// <summary>
        /// 窗口加载完成事件
        /// </summary>
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            InitializeApplication();
        }
        #endregion

        #region 初始化方法
        /// <summary>
        /// 初始化应用程序状态
        /// </summary>
        private void InitializeApplication()
        {
            fileLoader.LoadCompleted += OnLoadCompleted;

            // 设置初始状态
            txtStatus.Text = "应用程序已就绪";

            // 设置降采样滑块的初始值
            DownsampleSlider.Value = downsampleLevel;
            DownsampleValueText.Text = downsampleLevel.ToString();

            // 设置默认颜色模式 - 先设置visualizer，再设置UI控件避免触发事件
            visualizer.CurrentColorMode = PointCloudVisualizer.ColorMappingMode.ByFile;

            // 临时移除事件处理器，设置选中项，然后重新添加
            ColorModeComboBox.SelectionChanged -= ColorModeComboBox_SelectionChanged;
            ColorModeComboBox.SelectedIndex = 0; // 按文件区分
            ColorModeComboBox.SelectionChanged += ColorModeComboBox_SelectionChanged;

            // 设置默认输出目录
            try
            {
                var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                selectedOutputDir = docs;
                if (txtOutputDir != null) txtOutputDir.Text = selectedOutputDir;
            }
            catch { /* 忽略初始化错误 */ }
        }

        /// <summary>
        /// 颜色模式切换事件处理
        /// </summary>
        private void ColorModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 检查必要的对象是否已初始化
            if (visualizer == null || txtStatus == null || ColorModeComboBox == null)
                return;

            if (ColorModeComboBox.SelectedItem is ComboBoxItem selectedItem &&
                selectedItem.Tag is string modeTag)
            {
                // 解析颜色模式
                if (Enum.TryParse<PointCloudVisualizer.ColorMappingMode>(modeTag, out var colorMode))
                {
                    visualizer.CurrentColorMode = colorMode;

                    // 如果有数据，重新渲染以应用新的颜色方案
                    if (renderPointCloudData != null && renderPointCloudData.Count > 0)
                    {
                        RefreshVisualizationWithNewColors();
                    }

                    txtStatus.Text = $"颜色模式已切换为: {selectedItem.Content}";
                }
            }
        }

        /// <summary>
        /// 输出目录浏览按钮点击事件
        /// </summary>
        private void BtnBrowseOutputDir_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "选择STL输出目录",
                    CheckFileExists = false,
                    CheckPathExists = true,
                    FileName = "选择文件夹",
                    Filter = "文件夹|*.folder"
                };

                if (!string.IsNullOrWhiteSpace(selectedOutputDir) && Directory.Exists(selectedOutputDir))
                {
                    dlg.InitialDirectory = selectedOutputDir;
                }

                var result = dlg.ShowDialog();
                if (result == true && !string.IsNullOrWhiteSpace(dlg.FileName))
                {
                    var selectedPath = Path.GetDirectoryName(dlg.FileName);
                    if (!string.IsNullOrWhiteSpace(selectedPath))
                    {
                        selectedOutputDir = selectedPath;
                        txtOutputDir.Text = selectedOutputDir;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择目录时发生错误:\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 使用新颜色方案刷新可视化
        /// </summary>
        private async void RefreshVisualizationWithNewColors()
        {
            // 检查必要的对象是否已初始化
            if (visualizer == null || renderPointCloudData == null || txtStatus == null)
                return;

            try
            {
                txtStatus.Text = "正在应用新颜色方案...";

                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        // 重新创建可视化以应用新颜色
                        if (visualizer != null && renderPointCloudData != null)
                        {
                            // 更新文件信息
                            visualizer.SetFileInfoList(fileInfoList);
                            visualizer.SetPointFileIndices(renderPointFileIndices);
                            visualizer.CreatePointCloudVisualization(renderPointCloudData);

                            var colorModeName = ColorModeComboBox?.SelectedItem is ComboBoxItem item ?
                                              item.Content?.ToString() : "未知";
                            if (txtStatus != null)
                            {
                                txtStatus.Text = $"颜色方案已更新: {colorModeName} | 点数: {renderPointCloudData.Count:N0}";
                            }
                        }
                    });
                });
            }
            catch (Exception ex)
            {
                if (txtStatus != null)
                {
                    txtStatus.Text = $"颜色更新失败: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// 点云转STL按钮点击事件
        /// </summary>
        private async void BtnConvertToSTL_Click(object sender, RoutedEventArgs e)
        {
            if (renderPointCloudData.Count == 0)
            {
                MessageBox.Show("没有可转换的点云数据。请先加载并渲染点云。", "转换错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnConvertToSTL.IsEnabled = false;
                await ConvertRenderDataToSTLAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"点云转STL时发生错误:\n{ex.Message}", "转换错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnConvertToSTL.IsEnabled = renderPointCloudData.Count > 0;
            }
        }

        /// <summary>
        /// STL导出按钮点击事件 - 优化后的工作流
        /// </summary>
        private async void BtnExportSTL_Click(object sender, RoutedEventArgs e)
        {
            if (currentSTLMesh == null)
            {
                MessageBox.Show("没有可导出的STL模型。请先转换点云为STL。", "导出错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(selectedOutputDir) || !Directory.Exists(selectedOutputDir))
            {
                MessageBox.Show("请先选择有效的输出目录。", "输出目录错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnExportSTL.IsEnabled = false;
                UpdateRightProgress(10, "正在导出STL文件...");

                // 使用时间戳生成文件名
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var fileName = $"PointCloud_STL_{timestamp}.stl";
                var fullPath = Path.Combine(selectedOutputDir, fileName);

                await Task.Run(() =>
                {
                    stlConverter.ExportSTL(currentSTLMesh, fullPath);
                });

                UpdateRightProgress(100, $"STL文件已导出: {fileName}");
                txtStatus.Text = $"STL文件已成功导出: {fileName}";
                MessageBox.Show($"STL文件已成功保存到:\n{fullPath}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateRightProgress(0, "STL导出失败");
                txtStatus.Text = "STL导出失败";
                MessageBox.Show($"导出STL文件时发生错误:\n{ex.Message}", "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnExportSTL.IsEnabled = true;
            }
        }
        #endregion

        #region 事件处理器
        /// <summary>
        /// 加载文件按钮点击事件 - 支持多文件选择
        /// </summary>
        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择点云文件或STL文件",
                Filter = "支持的文件 (*.txt;*.stl)|*.txt;*.stl|点云文件 (*.txt)|*.txt|STL文件 (*.stl)|*.stl|所有文件 (*.*)|*.*",
                InitialDirectory = Environment.CurrentDirectory,
                Multiselect = true // 启用多文件选择
            };

            if (openFileDialog.ShowDialog() == true)
            {
                HandleMultipleFileSelection(openFileDialog.FileNames);
            }
        }

        /// <summary>
        /// 处理多文件选择
        /// </summary>
        private async void HandleMultipleFileSelection(string[] selectedFiles)
        {
            if (selectedFiles.Length == 0) return;

            // 分类文件类型
            var txtFiles = selectedFiles.Where(f => Path.GetExtension(f).ToLower() == ".txt").ToArray();
            var stlFiles = selectedFiles.Where(f => Path.GetExtension(f).ToLower() == ".stl").ToArray();

            // 检查文件类型组合
            if (txtFiles.Length > 0 && stlFiles.Length > 0)
            {
                MessageBox.Show("不能同时选择TXT和STL文件。请分别选择一种类型的文件。", "文件类型错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (txtFiles.Length > 0)
            {
                await ProcessTxtFilesAsync(txtFiles);
            }
            else if (stlFiles.Length > 0)
            {
                await ProcessStlFilesAsync(stlFiles);
            }
            else
            {
                MessageBox.Show("请选择TXT或STL格式的文件。", "不支持的文件格式", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// 处理TXT文件
        /// </summary>
        private async Task ProcessTxtFilesAsync(string[] txtFiles)
        {
            // 如果已有数据，询问用户操作
            if (loadedFilePaths.Count > 0)
            {
                var result = MessageBox.Show(
                    "当前已有加载的点云数据。\n\n点击\"是\"继续添加更多文件\n点击\"否\"重新选择（清空当前文件列表）\n点击\"取消\"取消操作",
                    "多文件加载选项",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // 继续添加
                        break;
                    case MessageBoxResult.No:
                        // 清空重新开始
                        ClearAllData();
                        break;
                    case MessageBoxResult.Cancel:
                        return;
                }
            }

            // 加载TXT文件
            await LoadMultipleFilesAsync(txtFiles);

            // 只有在渲染数据就绪时才启用转换按钮
            btnConvertToSTL.IsEnabled = renderPointCloudData.Count > 0;
        }

        /// <summary>
        /// 处理STL文件
        /// </summary>
        private async Task ProcessStlFilesAsync(string[] stlFiles)
        {
            if (stlFiles.Length > 1)
            {
                MessageBox.Show("一次只能加载一个STL文件。", "文件数量限制", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 清空右侧STL显示（不要清理左侧点云，以免禁用“点云→STL”按钮）
            ClearSTLDisplay();

            // 加载STL文件
            await LoadStlFileAsync(stlFiles[0]);
        }

        /// <summary>
        /// 清空所有数据
        /// </summary>
        private void ClearAllData()
        {
            fullPointCloudData.Clear();
            renderPointCloudData.Clear();
            renderPointFileIndices.Clear();
            loadedFilePaths.Clear();
            fileInfoList.Clear();

            // 重置全局变换
            ResetGlobalTransform();

            // 清空3D视图
            ClearPointCloudDisplay();
            ClearSTLDisplay();

            // 重置进度显示
            ResetProgressDisplays();

            txtStatus.Text = "已清空数据";
        }

        /// <summary>
        /// 清空点云显示
        /// </summary>
        private void ClearPointCloudDisplay()
        {
            if (currentPointCloud != null)
            {
                helixViewport.Items.Remove(currentPointCloud);
                currentPointCloud = null;
            }
            btnConvertToSTL.IsEnabled = false;
        }

        /// <summary>
        /// 清空STL显示
        /// </summary>
        private void ClearSTLDisplay()
        {
            if (currentSTLMesh != null)
            {
                stlViewport.Items.Remove(currentSTLMesh);
                currentSTLMesh = null;
            }
            btnExportSTL.IsEnabled = false;

            // 重置右侧进度显示
            UpdateRightProgress(0, "就绪");
        }

        /// <summary>
        /// 重置进度显示
        /// </summary>
        private void ResetProgressDisplays()
        {
            UpdateRightProgress(0, "就绪");
        }

        /// <summary>
        /// 加载STL文件
        /// </summary>
        private async Task LoadStlFileAsync(string filePath)
        {
            try
            {
                UpdateRightProgress(10, "正在读取STL文件...");

                var stlLoader = new STLLoader();
                var mesh = await Task.Run(() => stlLoader.LoadSTL(filePath));

                UpdateRightProgress(80, "正在显示STL模型...");

                DisplaySTLMesh(mesh);
                btnExportSTL.IsEnabled = true;

                // 关键：加载外部STL后，按左侧点云数据决定“点云→STL”按钮状态
                btnConvertToSTL.IsEnabled = renderPointCloudData.Count > 0;

                UpdateRightProgress(100, $"STL文件加载完成: {Path.GetFileName(filePath)}");
                txtStatus.Text = $"STL文件已加载: {Path.GetFileName(filePath)}";
            }
            catch (Exception ex)
            {
                UpdateRightProgress(0, $"STL加载失败: {ex.Message}");
                MessageBox.Show($"加载STL文件时发生错误:\n{ex.Message}", "加载错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        /// <summary>
        /// 更新右侧进度显示 - 使用圆形进度指示器
        /// </summary>
        private void UpdateRightProgress(double value, string text)
        {
            Dispatcher.Invoke(() =>
            {
                // 更新圆形进度指示器
                var circumference = Math.PI * 2 * 10.5; // 半径约为10.5
                var dashLength = (value / 100.0) * circumference;
                var gapLength = circumference - dashLength;

                rightProgressRing.StrokeDashArray = new System.Windows.Media.DoubleCollection { dashLength, gapLength };

                // 更新百分比文本
                rightProgressPercent.Text = $"{value:F0}%";

                // 更新状态文本
                rightProgressText.Text = text;

                // 根据进度值改变颜色
                var progressColor = value >= 100 ? "#4CAF50" : value >= 50 ? "#FF9800" : "#2196F3";
                rightProgressRing.Stroke = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(progressColor));
                rightProgressPercent.Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(progressColor));
            });
        }

        /// <summary>
        /// 文件加载完成事件处理器
        /// </summary>
        private void OnLoadCompleted(object? sender, PointCloudLoadResult result)
        {
            Dispatcher.Invoke(() =>
            {
                if (!result.IsSuccess)
                {
                    HandleLoadError(result.ErrorMessage);
                }

                btnOpenFile.IsEnabled = true;
            });
        }
        #endregion

        #region 文件加载方法
        /// <summary>
        /// 异步加载多个点云文件
        /// </summary>
        private async Task LoadMultipleFilesAsync(string[] filePaths)
        {
            btnOpenFile.IsEnabled = false;
            txtStatus.Text = "加载中...";

            try
            {
                int totalFiles = filePaths.Length;
                int loadedFiles = 0;
                var allNewPoints = new List<Vector3>();

                foreach (string filePath in filePaths)
                {
                    try
                    {
                        var fileName = Path.GetFileName(filePath);
                        var progressPercent = (double)loadedFiles / totalFiles * 100;

                        txtStatus.Text = $"加载文件 {loadedFiles + 1}/{totalFiles}: {fileName}";

                        // 检查文件是否存在
                        if (!File.Exists(filePath))
                        {
                            throw new FileNotFoundException($"文件不存在: {filePath}");
                        }

                        // 检查文件大小
                        var fileInfo = new FileInfo(filePath);
                        if (fileInfo.Length > 500 * 1024 * 1024) // 500MB限制
                        {
                            var result = MessageBox.Show(
                                $"文件 {Path.GetFileName(filePath)} 较大 ({fileInfo.Length / (1024 * 1024):F1} MB)，加载可能需要较长时间。\n\n是否继续？",
                                "大文件警告",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning);

                            if (result == MessageBoxResult.No)
                                continue;
                        }

                        var loadResult = await fileLoader.LoadPointCloudAsync(filePath);

                        if (loadResult.IsSuccess)
                        {
                            // 创建文件信息记录
                            var pointCloudFileInfo = new PointCloudFileInfo
                            {
                                FilePath = filePath,
                                FileName = Path.GetFileName(filePath),
                                StartIndex = allNewPoints.Count,
                                PointCount = loadResult.Points.Count,
                                FileIndex = loadedFiles
                            };

                            allNewPoints.AddRange(loadResult.Points);
                            loadedFilePaths.Add(filePath);
                            fileInfoList.Add(pointCloudFileInfo);
                            loadedFiles++;
                        }
                        else
                        {
                            MessageBox.Show($"加载文件失败: {Path.GetFileName(filePath)}\n错误: {loadResult.ErrorMessage}",
                                          "文件加载错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    catch (FileNotFoundException ex)
                    {
                        MessageBox.Show($"文件不存在: {Path.GetFileName(filePath)}\n{ex.Message}",
                                      "文件不存在", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        MessageBox.Show($"无权限访问文件: {Path.GetFileName(filePath)}\n{ex.Message}",
                                      "权限错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"文件读取错误: {Path.GetFileName(filePath)}\n{ex.Message}",
                                      "IO错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"加载文件时发生未知错误: {Path.GetFileName(filePath)}\n{ex.Message}",
                                      "未知错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }

                if (allNewPoints.Count > 0)
                {
                    // 检查内存使用情况
                    await CheckMemoryUsage(allNewPoints.Count);

                    // 添加到完整数据集
                    fullPointCloudData.AddRange(allNewPoints);

                    // 决定渲染策略
                    if (fullPointCloudData.Count > STREAMING_THRESHOLD)
                    {
                        await HandleStreamingRender();
                    }
                    else
                    {
                        await HandleStandardRender();
                    }

                    txtStatus.Text = $"成功加载 {loadedFiles}/{totalFiles} 个文件，共 {fullPointCloudData.Count:N0} 个点";
                }
                else
                {
                    txtStatus.Text = "没有加载到有效数据";
                }
            }
            catch (OutOfMemoryException ex)
            {
                HandleLoadError("内存不足，无法加载所有文件", ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                HandleLoadError("访问文件被拒绝", ex);
            }
            catch (IOException ex)
            {
                HandleLoadError("文件读取错误", ex);
            }
            catch (Exception ex)
            {
                HandleLoadError($"批量加载发生未知错误", ex);
            }
            finally
            {
                btnOpenFile.IsEnabled = true;
            }
        }



        /// <summary>
        /// 处理加载错误 - 增强版
        /// </summary>
        private void HandleLoadError(string errorMessage, Exception? ex = null)
        {
            string detailedMessage = errorMessage;
            string title = "加载错误";
            MessageBoxImage icon = MessageBoxImage.Error;

            // 根据异常类型提供具体的错误处理
            if (ex != null)
            {
                switch (ex)
                {
                    case OutOfMemoryException:
                        title = "内存不足";
                        detailedMessage = "系统内存不足，无法加载如此大的点云文件。\n\n建议：\n1. 关闭其他应用程序释放内存\n2. 尝试加载较小的文件\n3. 使用更高的降采样级别";
                        icon = MessageBoxImage.Warning;
                        break;

                    case FileNotFoundException:
                        title = "文件未找到";
                        detailedMessage = $"无法找到指定的文件。\n\n请检查：\n1. 文件路径是否正确\n2. 文件是否已被移动或删除\n3. 是否有足够的访问权限\n\n错误详情：{ex.Message}";
                        icon = MessageBoxImage.Warning;
                        break;

                    case UnauthorizedAccessException:
                        title = "访问被拒绝";
                        detailedMessage = $"没有权限访问该文件。\n\n请检查：\n1. 文件是否被其他程序占用\n2. 是否有足够的读取权限\n3. 文件是否为只读状态\n\n错误详情：{ex.Message}";
                        icon = MessageBoxImage.Warning;
                        break;

                    case IOException:
                        title = "文件读取错误";
                        detailedMessage = $"读取文件时发生错误。\n\n可能原因：\n1. 文件已损坏\n2. 磁盘空间不足\n3. 网络连接中断（网络文件）\n\n错误详情：{ex.Message}";
                        icon = MessageBoxImage.Error;
                        break;

                    case FormatException:
                        title = "文件格式错误";
                        detailedMessage = $"文件格式不正确或数据格式有误。\n\n请检查：\n1. 文件是否为有效的点云数据文件\n2. 数据格式是否符合要求\n3. 文件编码是否正确\n\n错误详情：{ex.Message}";
                        icon = MessageBoxImage.Warning;
                        break;

                    default:
                        detailedMessage = $"{errorMessage}\n\n技术详情：{ex.Message}\n\n如果问题持续存在，请联系技术支持。";
                        break;
                }
            }

            MessageBox.Show(detailedMessage, title, MessageBoxButton.OK, icon);
            txtStatus.Text = $"错误: {title}";
        }

        /// <summary>
        /// 检查内存使用情况
        /// </summary>
        private async Task CheckMemoryUsage(int pointCount)
        {
            await Task.Run(() =>
            {
                try
                {
                    // 估算点云数据内存使用量 (每个Vector3约12字节)
                    long estimatedMemory = pointCount * 12L;

                    // 获取当前可用内存
                    var currentMemory = GC.GetTotalMemory(false);

                    // 如果估算内存使用超过500MB，给出警告
                    if (estimatedMemory > 500 * 1024 * 1024)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            var result = MessageBox.Show(
                                $"即将加载 {pointCount:N0} 个点，估算内存使用: {estimatedMemory / (1024 * 1024):F1} MB\n\n" +
                                $"当前内存使用: {currentMemory / (1024 * 1024):F1} MB\n\n" +
                                "大量数据可能导致性能下降或内存不足。是否继续？",
                                "内存使用警告",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning);

                            if (result == MessageBoxResult.No)
                            {
                                throw new OperationCanceledException("用户取消了大数据加载操作");
                            }
                        });
                    }
                }
                catch (OperationCanceledException)
                {
                    throw; // 重新抛出用户取消异常
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        txtStatus.Text = $"内存检查失败: {ex.Message}";
                    });
                }
            });
        }

        /// <summary>
        /// 流式渲染处理 - 增强版
        /// </summary>
        private async Task HandleStreamingRender()
        {
            streamingCancellation = new CancellationTokenSource();
            var token = streamingCancellation.Token;

            try
            {
                txtStatus.Text = "准备流式渲染...";

                // 1. 计算全局变换参数
                await CalculateGlobalTransformAsync(fullPointCloudData, token);

                // 2. 应用全局变换
                var transformedData = new List<Vector3>(fullPointCloudData);
                ApplyGlobalTransform(transformedData);

                // 3. 应用降采样到变换后的数据
                var (downsampledData, fileIndices) = await Task.Run(() => VoxelGridDownsampleWithFileInfo(transformedData, downsampleLevel), token);
                renderPointCloudData.Clear();
                renderPointCloudData.AddRange(downsampledData);
                renderPointFileIndices.Clear();
                renderPointFileIndices.AddRange(fileIndices);

                // 3.5. 更新可视化器的文件信息
                visualizer.SetFileInfoList(fileInfoList);
                visualizer.SetPointFileIndices(renderPointFileIndices);

                // 4. 首帧快速渲染
                await RenderFirstChunk(renderPointCloudData, token);

                // 5. 流式渲染剩余数据
                await StreamRenderRemainingPoints(renderPointCloudData, token);
            }
            catch (OperationCanceledException)
            {
                txtStatus.Text = "渲染已取消";
            }
            catch (OutOfMemoryException ex)
            {
                HandleLoadError("流式渲染时内存不足", ex);
            }
            catch (Exception ex)
            {
                HandleLoadError("流式渲染发生错误", ex);
            }
            finally
            {
                CleanupStreamingResources();
            }
        }

        /// <summary>
        /// 首帧快速渲染优化
        /// </summary>
        private async Task RenderFirstChunk(List<Vector3> points, CancellationToken token)
        {
            if (points.Count == 0) return;

            txtStatus.Text = "首帧渲染中...";

            // 取前30000个点进行快速渲染
            var firstChunk = points.Take(FIRST_RENDER_CHUNK_SIZE).ToList();

            // 在UI线程上创建初始可视化
            visualizer.CreatePointCloudVisualization(firstChunk);

            txtStatus.Text = $"首帧完成: {firstChunk.Count:N0}/{points.Count:N0} 点";

            // 短暂延迟让UI更新
            await Task.Delay(100, token);
        }

        /// <summary>
        /// 流式渲染剩余点
        /// </summary>
        private async Task StreamRenderRemainingPoints(List<Vector3> points, CancellationToken token)
        {
            if (points.Count <= FIRST_RENDER_CHUNK_SIZE) return;

            txtStatus.Text = "流式渲染剩余数据...";

            var remainingPoints = points.Skip(FIRST_RENDER_CHUNK_SIZE).ToList();
            int renderedCount = FIRST_RENDER_CHUNK_SIZE;

            for (int i = 0; i < remainingPoints.Count; i += STREAMING_CHUNK_SIZE)
            {
                if (token.IsCancellationRequested) break;

                var chunk = remainingPoints.Skip(i).Take(STREAMING_CHUNK_SIZE).ToList();
                visualizer.AddPointsToVisualization(chunk);

                renderedCount += chunk.Count;
                txtStatus.Text = $"流式渲染: {renderedCount:N0}/{points.Count:N0} 点";

                // 短暂延迟避免UI阻塞
                await Task.Delay(50, token);
            }

            var reductionRatio = fullPointCloudData.Count > 0 ? (double)renderPointCloudData.Count / fullPointCloudData.Count : 1.0;
            txtStatus.Text = $"流式渲染完成: {renderPointCloudData.Count:N0}/{fullPointCloudData.Count:N0} 点 ({reductionRatio:P1})";

            // 渲染完成后启用转换按钮
            btnConvertToSTL.IsEnabled = renderPointCloudData.Count > 0;
        }

        /// <summary>
        /// 标准渲染处理 - 增强版
        /// </summary>
        private async Task HandleStandardRender()
        {
            try
            {
                txtStatus.Text = "渲染中...";

                // 1. 计算全局变换参数
                CalculateGlobalTransform(fullPointCloudData);

                // 2. 应用全局变换
                var transformedData = new List<Vector3>(fullPointCloudData);
                ApplyGlobalTransform(transformedData);

                // 3. 应用降采样
                var (downsampledData, fileIndices) = await Task.Run(() => VoxelGridDownsampleWithFileInfo(transformedData, downsampleLevel));
                renderPointCloudData.Clear();
                renderPointCloudData.AddRange(downsampledData);
                renderPointFileIndices.Clear();
                renderPointFileIndices.AddRange(fileIndices);

                // 3.5. 更新可视化器的文件信息
                visualizer.SetFileInfoList(fileInfoList);
                visualizer.SetPointFileIndices(renderPointFileIndices);

                // 4. 直接渲染
                visualizer.CreatePointCloudVisualization(renderPointCloudData);

                var reductionRatio = fullPointCloudData.Count > 0 ? (double)renderPointCloudData.Count / fullPointCloudData.Count : 1.0;
                var transformInfo = globalTransformCalculated ? $" | 已应用坐标变换 (缩放: {globalScale:F3})" : "";
                txtStatus.Text = $"渲染完成: {renderPointCloudData.Count:N0}/{fullPointCloudData.Count:N0} 点 ({reductionRatio:P1}){transformInfo}";

                // 标准渲染完成后启用转换按钮
                btnConvertToSTL.IsEnabled = renderPointCloudData.Count > 0;
            }
            catch (OutOfMemoryException ex)
            {
                HandleLoadError("渲染时内存不足", ex);
                btnConvertToSTL.IsEnabled = false; // 渲染失败时禁用转换按钮
            }
            catch (Exception ex)
            {
                HandleLoadError("渲染发生错误", ex);
                btnConvertToSTL.IsEnabled = false; // 渲染失败时禁用转换按钮
            }
        }



        /// <summary>
        /// 清理流式渲染资源
        /// </summary>
        private void CleanupStreamingResources()
        {
            streamingCancellation?.Dispose();
            streamingCancellation = null;
        }

        /// <summary>
        /// 降采样滑块值改变事件
        /// </summary>
        private void DownsampleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (DownsampleValueText != null)
            {
                int newLevel = (int)e.NewValue;
                DownsampleValueText.Text = newLevel.ToString();

                if (newLevel != downsampleLevel)
                {
                    downsampleLevel = newLevel;

                    // 如果有数据，重新应用降采样
                    if (fullPointCloudData.Count > 0)
                    {
                        ApplyDownsamplingAndRefresh();
                    }
                }
            }
        }

        /// <summary>
        /// 应用降采样并刷新显示
        /// </summary>
        private async void ApplyDownsamplingAndRefresh()
        {
            if (fullPointCloudData.Count == 0) return;

            try
            {
                txtStatus.Text = "正在应用降采样...";
                btnOpenFile.IsEnabled = false;

                await Task.Run(() =>
                {
                    // 1. 确保全局变换已计算
                    if (!globalTransformCalculated)
                    {
                        CalculateGlobalTransform(fullPointCloudData);
                    }

                    // 2. 应用全局变换
                    var transformedData = new List<Vector3>(fullPointCloudData);
                    ApplyGlobalTransform(transformedData);

                    // 3. 应用降采样
                    var (downsampledData, fileIndices) = VoxelGridDownsampleWithFileInfo(transformedData, downsampleLevel);

                    Dispatcher.Invoke(() =>
                    {
                        renderPointCloudData.Clear();
                        renderPointCloudData.AddRange(downsampledData);
                        renderPointFileIndices.Clear();
                        renderPointFileIndices.AddRange(fileIndices);

                        // 更新可视化器的文件信息
                        visualizer.SetFileInfoList(fileInfoList);
                        visualizer.SetPointFileIndices(renderPointFileIndices);

                        // 更新3D显示
                        visualizer.CreatePointCloudVisualization(renderPointCloudData);

                        // 更新状态信息
                        var reductionRatio = fullPointCloudData.Count > 0 ? (double)renderPointCloudData.Count / fullPointCloudData.Count : 1.0;
                        var transformInfo = globalTransformCalculated ? $" | 变换缩放: {globalScale:F3}" : "";
                        txtStatus.Text = $"降采样完成: {renderPointCloudData.Count:N0}/{fullPointCloudData.Count:N0} 点 ({reductionRatio:P1}) | 级别: {downsampleLevel}{transformInfo}";
                    });
                });
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"降采样失败: {ex.Message}";
            }
            finally
            {
                btnOpenFile.IsEnabled = true;
            }
        }

        #region === 全局变换处理 ===

        /// <summary>
        /// 计算全局变换参数（解决大坐标值渲染问题）
        /// </summary>
        private void CalculateGlobalTransform(List<Vector3> allPoints)
        {
            if (allPoints.Count == 0)
            {
                globalCenter = new Vector3(0, 0, 0);
                globalScale = 1.0;
                globalTransformCalculated = true;
                return;
            }

            try
            {
                // 计算边界框
                var minX = allPoints.Min(p => p.X);
                var maxX = allPoints.Max(p => p.X);
                var minY = allPoints.Min(p => p.Y);
                var maxY = allPoints.Max(p => p.Y);
                var minZ = allPoints.Min(p => p.Z);
                var maxZ = allPoints.Max(p => p.Z);

                // 计算中心点和缩放因子
                globalCenter = new Vector3((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
                var maxRange = Math.Max(Math.Max(maxX - minX, maxY - minY), maxZ - minZ);
                globalScale = maxRange > 0 ? 100.0 / maxRange : 1.0;
                globalTransformCalculated = true;

                // 调试信息
                Dispatcher.Invoke(() =>
                {
                    txtStatus.Text = $"全局变换计算完成 | 中心: ({globalCenter.X:F2}, {globalCenter.Y:F2}, {globalCenter.Z:F2}) | 缩放: {globalScale:F4}";
                });
            }
            catch (Exception ex)
            {
                // 全局变换计算失败时使用默认值
                globalCenter = new Vector3(0, 0, 0);
                globalScale = 1.0;
                globalTransformCalculated = true;

                Dispatcher.Invoke(() =>
                {
                    txtStatus.Text = $"全局变换计算失败，使用默认值: {ex.Message}";
                });
            }
        }

        /// <summary>
        /// 异步计算全局变换参数
        /// </summary>
        private async Task CalculateGlobalTransformAsync(List<Vector3> points, CancellationToken token)
        {
            await Task.Run(() =>
            {
                CalculateGlobalTransform(points);
            }, token);
        }

        /// <summary>
        /// 应用全局变换到点云数据
        /// </summary>
        private void ApplyGlobalTransform(List<Vector3> points)
        {
            if (!globalTransformCalculated || points.Count == 0) return;

            try
            {
                for (int i = 0; i < points.Count; i++)
                {
                    var p = points[i];
                    points[i] = new Vector3(
                        (p.X - globalCenter.X) * (float)globalScale,
                        (p.Y - globalCenter.Y) * (float)globalScale,
                        (p.Z - globalCenter.Z) * (float)globalScale
                    );
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    txtStatus.Text = $"全局变换应用失败: {ex.Message}";
                });
            }
        }

        /// <summary>
        /// 重置全局变换参数
        /// </summary>
        private void ResetGlobalTransform()
        {
            globalTransformCalculated = false;
            globalCenter = new Vector3(0, 0, 0);
            globalScale = 1.0;
        }

        #endregion

        #region === 降采样算法 ===

        /// <summary>
        /// 体素网格降采样算法（带文件索引）
        /// </summary>
        public (List<Vector3> points, List<int> fileIndices) VoxelGridDownsampleWithFileInfo(List<Vector3> points, int downsampleLevel = 1)
        {
            if (points?.Count <= 2 || downsampleLevel < 1)
            {
                var fileIndices = CreateFileIndicesForPoints(points?.Count ?? 0);
                return (points ?? new List<Vector3>(), fileIndices);
            }

            // 根据降采样级别计算目标点数和体素大小
            int targetPoints = CalculateTargetPoints(points.Count, downsampleLevel);
            if (points.Count <= targetPoints)
            {
                var fileIndices = CreateFileIndicesForPoints(points.Count);
                return (new List<Vector3>(points), fileIndices);
            }

            var bounds = CalculateBounds(points);
            double voxelSize = CalculateOptimalVoxelSize(bounds, targetPoints, downsampleLevel);
            var voxelGrid = BuildVoxelGridWithFileInfo(points, bounds, voxelSize);
            var (resultPoints, resultFileIndices) = ExtractRepresentativePointsWithFileInfo(voxelGrid);

            // 在UI线程上更新调试信息
            Dispatcher.Invoke(() =>
            {
                var reductionRatio = points.Count > 0 ? (double)resultPoints.Count / points.Count : 1.0;
                var xRange = bounds.maxX - bounds.minX;
                var yRange = bounds.maxY - bounds.minY;
                var zRange = bounds.maxZ - bounds.minZ;

                txtStatus.Text = $"降采样完成: {resultPoints.Count:N0}/{points.Count:N0} 点 ({reductionRatio:P1}) | " +
                               $"级别: {downsampleLevel} | 体素大小: {voxelSize:F4} | " +
                               $"点云范围: X={xRange:F2}, Y={yRange:F2}, Z={zRange:F2}";
            });

            return (resultPoints, resultFileIndices);
        }

        /// <summary>
        /// 体素网格降采样算法（原有方法，保持兼容性）
        /// </summary>
        public List<Vector3> VoxelGridDownsample(List<Vector3> points, int downsampleLevel = 1)
        {
            if (points?.Count <= 2 || downsampleLevel < 1)
                return points ?? new List<Vector3>();

            // 根据降采样级别计算目标点数和体素大小
            int targetPoints = CalculateTargetPoints(points.Count, downsampleLevel);
            if (points.Count <= targetPoints)
                return new List<Vector3>(points);

            var bounds = CalculateBounds(points);
            double voxelSize = CalculateOptimalVoxelSize(bounds, targetPoints, downsampleLevel);
            var voxelGrid = BuildVoxelGrid(points, bounds, voxelSize);
            var result = ExtractRepresentativePoints(voxelGrid);

            // 在UI线程上更新调试信息
            Dispatcher.Invoke(() =>
            {
                var reductionRatio = points.Count > 0 ? (double)result.Count / points.Count : 1.0;
                var xRange = bounds.maxX - bounds.minX;
                var yRange = bounds.maxY - bounds.minY;
                var zRange = bounds.maxZ - bounds.minZ;

                txtStatus.Text = $"降采样完成: {result.Count:N0}/{points.Count:N0} 点 ({reductionRatio:P1}) | " +
                               $"级别: {downsampleLevel} | 体素大小: {voxelSize:F4} | " +
                               $"点云范围: X={xRange:F2}, Y={yRange:F2}, Z={zRange:F2}";
            });

            return result;
        }

        /// <summary>
        /// 根据降采样级别计算目标点数
        /// </summary>
        private int CalculateTargetPoints(int originalCount, int downsampleLevel)
        {
            // 调整降采样策略：级别1-10，保留更多点，降采样更温和
            double retentionRatio = downsampleLevel switch
            {
                1 => 0.95,  // 级别1：保留95%的点
                2 => 0.85,  // 级别2：保留85%的点
                3 => 0.75,  // 级别3：保留75%的点
                4 => 0.65,  // 级别4：保留65%的点
                5 => 0.55,  // 级别5：保留55%的点
                6 => 0.45,  // 级别6：保留45%的点
                7 => 0.35,  // 级别7：保留35%的点
                8 => 0.25,  // 级别8：保留25%的点
                9 => 0.15,  // 级别9：保留15%的点
                10 => 0.10, // 级别10：保留10%的点
                _ => 1.0
            };

            return Math.Max(1000, (int)(originalCount * retentionRatio));
        }

        /// <summary>
        /// 计算点云边界
        /// </summary>
        private (float minX, float maxX, float minY, float maxY, float minZ, float maxZ) CalculateBounds(List<Vector3> points)
        {
            return (
                points.Min(p => p.X), points.Max(p => p.X),
                points.Min(p => p.Y), points.Max(p => p.Y),
                points.Min(p => p.Z), points.Max(p => p.Z)
            );
        }

        /// <summary>
        /// 计算最优体素大小 - 优化为更小的体素单位
        /// </summary>
        private double CalculateOptimalVoxelSize((float minX, float maxX, float minY, float maxY, float minZ, float maxZ) bounds, int targetPoints, int downsampleLevel)
        {
            // 计算点云的尺寸
            var xRange = bounds.maxX - bounds.minX;
            var yRange = bounds.maxY - bounds.minY;
            var zRange = bounds.maxZ - bounds.minZ;

            // 使用更保守的体素大小计算方法
            var avgDimension = (xRange + yRange + zRange) / 3.0;

            // 基础体素大小：根据平均尺寸和目标点数计算
            var baseVoxelSize = avgDimension / Math.Pow(targetPoints, 1.0 / 3.0);

            // 应用缩放因子，使体素更小更精细
            var scaleFactor = downsampleLevel switch
            {
                1 => 0.1,   // 级别1：非常小的体素
                2 => 0.15,  // 级别2：小体素
                3 => 0.2,   // 级别3：较小体素
                4 => 0.25,  // 级别4：中小体素
                5 => 0.3,   // 级别5：中等体素
                6 => 0.4,   // 级别6：中大体素
                7 => 0.5,   // 级别7：较大体素
                8 => 0.6,   // 级别8：大体素
                9 => 0.8,   // 级别9：很大体素
                10 => 1.0,  // 级别10：最大体素
                _ => 0.3
            };

            var finalVoxelSize = baseVoxelSize * scaleFactor;

            // 确保体素大小不会太小（避免性能问题）或太大（避免过度降采样）
            var minVoxelSize = Math.Min(xRange, Math.Min(yRange, zRange)) * 0.001; // 最小尺寸的0.1%
            var maxVoxelSize = Math.Max(xRange, Math.Max(yRange, zRange)) * 0.1;   // 最大尺寸的10%

            return Math.Max(minVoxelSize, Math.Min(maxVoxelSize, finalVoxelSize));
        }

        /// <summary>
        /// 构建体素网格
        /// </summary>
        private Dictionary<(int, int, int), List<Vector3>> BuildVoxelGrid(List<Vector3> points,
            (float minX, float maxX, float minY, float maxY, float minZ, float maxZ) bounds, double voxelSize)
        {
            var voxelGrid = new Dictionary<(int, int, int), List<Vector3>>();

            foreach (var point in points)
            {
                var key = (
                    (int)((point.X - bounds.minX) / voxelSize),
                    (int)((point.Y - bounds.minY) / voxelSize),
                    (int)((point.Z - bounds.minZ) / voxelSize)
                );

                if (!voxelGrid.ContainsKey(key))
                    voxelGrid[key] = new List<Vector3>();

                voxelGrid[key].Add(point);
            }

            return voxelGrid;
        }

        /// <summary>
        /// 提取代表点
        /// </summary>
        private List<Vector3> ExtractRepresentativePoints(Dictionary<(int, int, int), List<Vector3>> voxelGrid)
        {
            return voxelGrid.Values.Select(voxel => voxel.First()).ToList();
        }

        /// <summary>
        /// 为点创建文件索引数组
        /// </summary>
        private List<int> CreateFileIndicesForPoints(int pointCount)
        {
            var fileIndices = new List<int>();
            int currentIndex = 0;

            for (int i = 0; i < pointCount; i++)
            {
                // 找到当前点属于哪个文件
                int fileIndex = 0;
                int accumulatedCount = 0;

                for (int j = 0; j < fileInfoList.Count; j++)
                {
                    if (currentIndex >= accumulatedCount && currentIndex < accumulatedCount + fileInfoList[j].PointCount)
                    {
                        fileIndex = j;
                        break;
                    }
                    accumulatedCount += fileInfoList[j].PointCount;
                }

                fileIndices.Add(fileIndex);
                currentIndex++;
            }

            return fileIndices;
        }

        /// <summary>
        /// 构建带文件信息的体素网格
        /// </summary>
        private Dictionary<(int, int, int), List<(Vector3 point, int fileIndex)>> BuildVoxelGridWithFileInfo(
            List<Vector3> points,
            (float minX, float maxX, float minY, float maxY, float minZ, float maxZ) bounds,
            double voxelSize)
        {
            var voxelGrid = new Dictionary<(int, int, int), List<(Vector3, int)>>();

            for (int i = 0; i < points.Count; i++)
            {
                var point = points[i];
                var fileIndex = GetOriginalFileIndexForPoint(i);

                var key = (
                    (int)((point.X - bounds.minX) / voxelSize),
                    (int)((point.Y - bounds.minY) / voxelSize),
                    (int)((point.Z - bounds.minZ) / voxelSize)
                );

                if (!voxelGrid.ContainsKey(key))
                    voxelGrid[key] = new List<(Vector3, int)>();

                voxelGrid[key].Add((point, fileIndex));
            }

            return voxelGrid;
        }

        /// <summary>
        /// 提取带文件信息的代表点
        /// </summary>
        private (List<Vector3> points, List<int> fileIndices) ExtractRepresentativePointsWithFileInfo(
            Dictionary<(int, int, int), List<(Vector3 point, int fileIndex)>> voxelGrid)
        {
            var points = new List<Vector3>();
            var fileIndices = new List<int>();

            foreach (var voxel in voxelGrid.Values)
            {
                var representative = voxel.First();
                points.Add(representative.point);
                fileIndices.Add(representative.fileIndex);
            }

            return (points, fileIndices);
        }

        /// <summary>
        /// 获取原始数据中点的文件索引
        /// </summary>
        private int GetOriginalFileIndexForPoint(int pointIndex)
        {
            if (fileInfoList.Count == 0) return 0;

            int currentIndex = 0;
            for (int i = 0; i < fileInfoList.Count; i++)
            {
                if (pointIndex >= currentIndex && pointIndex < currentIndex + fileInfoList[i].PointCount)
                {
                    return i;
                }
                currentIndex += fileInfoList[i].PointCount;
            }

            return fileInfoList.Count - 1;
        }

        #endregion

        #region === STL转换功能 ===

        /// <summary>
        /// 查找项目根目录
        /// </summary>
        private string? FindProjectRoot()
        {
            try
            {
                // 从当前执行目录开始向上查找
                var currentDir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);

                while (currentDir != null)
                {
                    // 查找包含pyFunc目录的位置
                    var pyFuncPath = Path.Combine(currentDir.FullName, "pyFunc");
                    var sharpDxPyFuncPath = Path.Combine(currentDir.FullName, "SharpDX_PCV", "pyFunc");

                    if (Directory.Exists(pyFuncPath) && File.Exists(Path.Combine(pyFuncPath, "pointcloud_to_stl.py")))
                    {
                        return currentDir.FullName;
                    }

                    if (Directory.Exists(sharpDxPyFuncPath) && File.Exists(Path.Combine(sharpDxPyFuncPath, "pointcloud_to_stl.py")))
                    {
                        return currentDir.FullName;
                    }

                    // 也可以通过解决方案文件来识别
                    if (File.Exists(Path.Combine(currentDir.FullName, "SharpDX_PCV.sln")))
                    {
                        return currentDir.FullName;
                    }

                    currentDir = currentDir.Parent;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 异步将渲染数据转换为STL模型（使用Python）
        /// </summary>
        private async Task ConvertRenderDataToSTLAsync()
        {
            if (renderPointCloudData.Count == 0) return;

            try
            {
                UpdateRightProgress(5, "准备Python转换环境...");

                // 将当前渲染数据写入临时TXT文件（使用快照，避免并发修改）
                var tempTxtPath = Path.Combine(Path.GetTempPath(), $"render_points_{Guid.NewGuid():N}.txt");
                var pointsSnapshot = renderPointCloudData.ToArray();
                using (var sw = new StreamWriter(tempTxtPath))
                {
                    foreach (var p in pointsSnapshot)
                    {
                        await sw.WriteLineAsync(string.Format(CultureInfo.InvariantCulture, "{0} {1} {2}", p.X, p.Y, p.Z));
                    }
                }

                // 确定输出目录和文件名
                var outputDir = string.IsNullOrWhiteSpace(selectedOutputDir)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : selectedOutputDir;
                var outputName = $"pointcloud_{DateTime.Now:yyyyMMdd_HHmmss}.stl";
                var tempJsonPath = Path.Combine(Path.GetTempPath(), $"result_{Guid.NewGuid():N}.json");

                // 调用Python转换脚本（输入使用预处理后的渲染点云文件）
                var result = await RunPythonConversionAsync(new[] { tempTxtPath }, outputDir, tempJsonPath, outputName);

                if (result.Success && !string.IsNullOrWhiteSpace(result.OutputPath))
                {
                    UpdateRightProgress(90, "加载转换后的STL模型...");

                    // 加载生成的STL文件
                    var stlLoader = new STLLoader();
                    var mesh = await Task.Run(() => stlLoader.LoadSTL(result.OutputPath));

                    UpdateRightProgress(95, "正在显示STL模型...");

                    // 在UI线程中更新显示
                    Dispatcher.Invoke(() =>
                    {
                        DisplaySTLMesh(mesh);
                        btnExportSTL.IsEnabled = true;
                        UpdateRightProgress(100, $"STL转换完成 | 输入点数: {result.OriginalPoints:N0} | 三角形数: {result.FinalTriangles:N0}");
                        txtStatus.Text = $"STL转换完成 | 输入点数: {result.OriginalPoints:N0} | 三角形数: {result.FinalTriangles:N0}";
                    });

                    // 清理临时文件
                    try
                    {
                        if (File.Exists(tempJsonPath)) File.Delete(tempJsonPath);
                        if (File.Exists(tempTxtPath)) File.Delete(tempTxtPath);
                    }
                    catch { /* 忽略清理错误 */ }
                }
                else
                {
                    throw new Exception(result.ErrorMessage ?? "Python转换失败");
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    UpdateRightProgress(0, $"STL转换失败: {ex.Message}");
                    txtStatus.Text = $"STL转换失败: {ex.Message}";
                    btnExportSTL.IsEnabled = false;
                });
            }
        }

        /// <summary>
        /// Python转换结果
        /// </summary>
        public class PythonConversionResult
        {
            public bool Success { get; set; }
            public string? ErrorMessage { get; set; }
            public string? OutputPath { get; set; }
            public int OriginalPoints { get; set; }
            public int FinalTriangles { get; set; }
            public int FinalVertices { get; set; }
        }

        /// <summary>
        /// 运行Python转换脚本
        /// </summary>
        private async Task<PythonConversionResult> RunPythonConversionAsync(string[] inputFiles, string outputDir, string jsonOutputPath, string? outputName = null)
        {
            try
            {
                // 获取项目根目录（更可靠的路径解析）
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string exeDir = AppContext.BaseDirectory;
                string currentDir = Environment.CurrentDirectory;

                // 尝试找到项目根目录（包含SharpDX_PCV.sln或pyFunc目录的位置）
                string? projectRoot = FindProjectRoot();

                string[] candidateScriptPaths = new[]
                {
                    // 优先使用项目根目录
                    projectRoot != null ? Path.Combine(projectRoot, "SharpDX_PCV", "pyFunc", "pointcloud_to_stl.py") : null,
                    projectRoot != null ? Path.Combine(projectRoot, "pyFunc", "pointcloud_to_stl.py") : null,
                    // 备用路径
                    Path.Combine(baseDir, "pyFunc", "pointcloud_to_stl.py"),
                    Path.Combine(exeDir, "pyFunc", "pointcloud_to_stl.py"),
                    Path.Combine(currentDir, "pyFunc", "pointcloud_to_stl.py"),
                    Path.Combine(baseDir, "..", "..", "pyFunc", "pointcloud_to_stl.py"),
                    Path.Combine(baseDir, "..", "..", "SharpDX_PCV", "pyFunc", "pointcloud_to_stl.py")
                }.Where(p => p != null).Cast<string>().ToArray();

                string? scriptPath = candidateScriptPaths.FirstOrDefault(File.Exists);
                if (scriptPath == null)
                {
                    throw new FileNotFoundException("Python脚本不存在: " + string.Join(" | ", candidateScriptPaths));
                }

                // 获取脚本所在的pyFunc目录
                string pyFuncDir = Path.GetDirectoryName(scriptPath)!;

                string[] candidatePythonExePaths = new[]
                {
                    // 优先使用脚本同目录下的虚拟环境
                    Path.Combine(pyFuncDir, ".venv", "Scripts", "python.exe"),
                    // 备用路径
                    projectRoot != null ? Path.Combine(projectRoot, "SharpDX_PCV", "pyFunc", ".venv", "Scripts", "python.exe") : null,
                    projectRoot != null ? Path.Combine(projectRoot, "pyFunc", ".venv", "Scripts", "python.exe") : null,
                    Path.Combine(baseDir, "pyFunc", ".venv", "Scripts", "python.exe"),
                    Path.Combine(exeDir, "pyFunc", ".venv", "Scripts", "python.exe"),
                    Path.Combine(currentDir, "pyFunc", ".venv", "Scripts", "python.exe")
                }.Where(p => p != null).Cast<string>().ToArray();

                string? pythonExePath = candidatePythonExePaths.FirstOrDefault(File.Exists);
                if (pythonExePath == null)
                {
                    throw new FileNotFoundException("Python解释器不存在: " + string.Join(" | ", candidatePythonExePaths));
                }

                // 调试信息
                UpdateRightProgress(8, $"找到Python: {Path.GetFileName(pythonExePath)}");
                UpdateRightProgress(9, $"脚本路径: {Path.GetFileName(scriptPath)}");

                // 构建命令行参数
                var args = new List<string>
                {
                    $"\"{scriptPath}\"",
                    "--output-dir", $"\"{outputDir}\"",
                    "--json-output", $"\"{jsonOutputPath}\""
                };

                if (!string.IsNullOrWhiteSpace(outputName))
                {
                    args.Add("--output-name");
                    args.Add($"\"{outputName}\"");
                }

                // 添加输入文件
                foreach (var file in inputFiles)
                {
                    args.Add($"\"{file}\"");
                }

                var argumentString = string.Join(" ", args);

                UpdateRightProgress(10, "启动Python转换进程...");

                // 创建进程（使用正确的工作目录和环境变量）
                var processInfo = new ProcessStartInfo
                {
                    FileName = pythonExePath,
                    Arguments = argumentString,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = pyFuncDir  // 使用脚本所在的pyFunc目录
                };

                // 设置环境变量以确保Python能找到虚拟环境
                var venvPath = Path.GetDirectoryName(pythonExePath);
                if (venvPath != null)
                {
                    var venvRoot = Path.GetDirectoryName(venvPath); // .venv目录
                    var scriptsPath = Path.Combine(venvRoot!, "Scripts");
                    var libPath = Path.Combine(venvRoot!, "Lib", "site-packages");

                    // 设置VIRTUAL_ENV环境变量
                    processInfo.EnvironmentVariables["VIRTUAL_ENV"] = venvRoot;

                    // 更新PATH环境变量，确保虚拟环境的Scripts目录在最前面
                    var currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                    processInfo.EnvironmentVariables["PATH"] = $"{scriptsPath};{currentPath}";

                    // 设置PYTHONPATH
                    processInfo.EnvironmentVariables["PYTHONPATH"] = libPath;
                }

                using (var process = new Process { StartInfo = processInfo })
                {
                    var outputBuilder = new System.Text.StringBuilder();
                    var errorBuilder = new System.Text.StringBuilder();

                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            outputBuilder.AppendLine(e.Data);

                            // 解析进度信息
                            if (e.Data.Contains("[") && e.Data.Contains("%]"))
                            {
                                try
                                {
                                    var start = e.Data.IndexOf('[') + 1;
                                    var end = e.Data.IndexOf('%');
                                    if (start > 0 && end > start)
                                    {
                                        var progressStr = e.Data.Substring(start, end - start).Trim();
                                        if (double.TryParse(progressStr, out double progress))
                                        {
                                            var message = e.Data.Substring(e.Data.IndexOf(']') + 1).Trim();
                                            UpdateRightProgress(10 + progress * 0.75, message); // 10-85%的进度范围
                                        }
                                    }
                                }
                                catch { /* 忽略进度解析错误 */ }
                            }
                        }
                    };

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            errorBuilder.AppendLine(e.Data);
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    await process.WaitForExitAsync();

                    if (process.ExitCode == 0)
                    {
                        // 读取JSON结果
                        if (File.Exists(jsonOutputPath))
                        {
                            var jsonContent = await File.ReadAllTextAsync(jsonOutputPath);
                            using var jsonDoc = JsonDocument.Parse(jsonContent);
                            var root = jsonDoc.RootElement;

                            return new PythonConversionResult
                            {
                                Success = root.TryGetProperty("success", out var succEl) && succEl.ValueKind == JsonValueKind.True,
                                OutputPath = root.TryGetProperty("output_path", out var outEl) && outEl.ValueKind == JsonValueKind.String ? outEl.GetString() : null,
                                OriginalPoints = root.TryGetProperty("original_points", out var pEl) && pEl.TryGetInt32(out var pVal) ? pVal : 0,
                                FinalTriangles = root.TryGetProperty("final_triangles", out var tEl) && tEl.TryGetInt32(out var tVal) ? tVal : 0,
                                FinalVertices = root.TryGetProperty("final_vertices", out var vEl) && vEl.TryGetInt32(out var vVal) ? vVal : 0,
                                ErrorMessage = root.TryGetProperty("error", out var eEl) && eEl.ValueKind == JsonValueKind.String ? eEl.GetString() : null
                            };
                        }
                        else
                        {
                            // 如果没有JSON文件，构建预期的输出路径
                            var expectedPath = outputName != null ? Path.Combine(outputDir, outputName) : null;
                            return new PythonConversionResult
                            {
                                Success = true,
                                OutputPath = expectedPath
                            };
                        }
                    }
                    else
                    {
                        var errorMessage = errorBuilder.ToString();
                        if (string.IsNullOrEmpty(errorMessage))
                        {
                            errorMessage = outputBuilder.ToString();
                        }

                        return new PythonConversionResult
                        {
                            Success = false,
                            ErrorMessage = $"Python进程退出码: {process.ExitCode}\n{errorMessage}"
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                return new PythonConversionResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// 在STL视口中显示网格模型
        /// </summary>
        private void DisplaySTLMesh(HelixToolkit.Wpf.SharpDX.MeshGeometry3D mesh)
        {
            // 移除现有的STL模型
            if (currentSTLMesh != null)
            {
                stlViewport.Items.Remove(currentSTLMesh);
            }

            // 创建新的STL模型
            currentSTLMesh = new MeshGeometryModel3D
            {
                Geometry = mesh,
                Material = new PhongMaterial
                {
                    DiffuseColor = new SharpDX.Color4(0.7f, 0.7f, 0.9f, 1.0f), // 浅蓝色
                    SpecularColor = new SharpDX.Color4(0.3f, 0.3f, 0.3f, 1.0f),
                    SpecularShininess = 30f
                }
            };

            // 添加到STL视口
            stlViewport.Items.Add(currentSTLMesh);

            // 调整相机以适应模型
            AdjustSTLCameraToFitMesh(mesh);
        }

        /// <summary>
        /// 调整STL相机以适应网格模型
        /// </summary>
        private void AdjustSTLCameraToFitMesh(HelixToolkit.Wpf.SharpDX.MeshGeometry3D mesh)
        {
            if (mesh.Positions.Count == 0) return;

            // 计算边界框
            var positions = mesh.Positions;
            var minX = positions.Min(p => p.X);
            var maxX = positions.Max(p => p.X);
            var minY = positions.Min(p => p.Y);
            var maxY = positions.Max(p => p.Y);
            var minZ = positions.Min(p => p.Z);
            var maxZ = positions.Max(p => p.Z);

            var center = new Vector3((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
            var size = Math.Max(Math.Max(maxX - minX, maxY - minY), maxZ - minZ);
            var distance = size * 2.0;

            if (stlViewport.Camera is HelixToolkit.Wpf.SharpDX.PerspectiveCamera perspectiveCamera)
            {
                perspectiveCamera.Position = new Point3D(center.X + distance, center.Y + distance, center.Z + distance);
                perspectiveCamera.LookDirection = new Vector3D(-distance, -distance, -distance);
            }
        }

        #endregion

        #endregion

    }

    #region 辅助类定义
    /// <summary>
    /// 点云文件信息
    /// </summary>
    public class PointCloudFileInfo
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public int StartIndex { get; set; }
        public int PointCount { get; set; }
        public int FileIndex { get; set; }
    }

    /// <summary>
    /// 点云加载结果
    /// </summary>
    public class PointCloudLoadResult
    {
        public bool IsSuccess { get; set; }
        public List<Vector3> Points { get; set; } = new();
        public string FilePath { get; set; } = string.Empty;
        public string DataInfo { get; set; } = string.Empty;
        public long LoadTimeMs { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;

        /// <summary>
        /// 创建成功的加载结果
        /// </summary>
        public static PointCloudLoadResult Success(string filePath, List<Vector3> points, long loadTimeMs, string dataInfo)
        {
            return new PointCloudLoadResult
            {
                IsSuccess = true,
                FilePath = filePath,
                Points = points,
                LoadTimeMs = loadTimeMs,
                DataInfo = dataInfo
            };
        }

        /// <summary>
        /// 创建失败的加载结果
        /// </summary>
        public static PointCloudLoadResult Failure(string filePath, string errorMessage)
        {
            return new PointCloudLoadResult
            {
                IsSuccess = false,
                FilePath = filePath,
                ErrorMessage = errorMessage
            };
        }
    }

    /// <summary>
    /// 点云文件加载器
    /// </summary>
    public class PointCloudFileLoader
    {
        #region 常量定义
        private const long MaxFileSizeBytes = 100 * 1024 * 1024; // 100MB
        private const double BytesToMB = 1024.0 * 1024.0;
        #endregion

        #region 事件定义
        public event EventHandler<string>? LoadProgressChanged;
        public event EventHandler<PointCloudLoadResult>? LoadCompleted;
        #endregion

        /// <summary>
        /// 异步加载点云文件
        /// </summary>
        public async Task<PointCloudLoadResult> LoadPointCloudAsync(string filePath)
        {
            var stopwatch = Stopwatch.StartNew();

            try
            {
                // 验证文件存在
                if (!File.Exists(filePath))
                {
                    return PointCloudLoadResult.Failure(filePath, $"File not found: {filePath}");
                }

                // 检查大文件
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length > MaxFileSizeBytes)
                {
                    var dialogResult = MessageBox.Show(
                        $"所选文件过大，有：({fileInfo.Length / BytesToMB:F1} MB)。加载需要等待。继续？",
                        "大文件警告",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (dialogResult != MessageBoxResult.Yes)
                    {
                        return PointCloudLoadResult.Failure(filePath, "文件加载由用户取消！");
                    }
                }

                LoadProgressChanged?.Invoke(this, "加载中...");

                // 异步加载数据
                var points = new List<Vector3>();
                await Task.Run(() => LoadPointCloudData(filePath, points));

                stopwatch.Stop();

                if (points.Count == 0)
                {
                    return PointCloudLoadResult.Failure(filePath, "所选文件没有有效数据点。请检查文件格式。");
                }

                var dataInfo = AnalyzeLoadedData(points, filePath);
                var result = PointCloudLoadResult.Success(filePath, points, stopwatch.ElapsedMilliseconds, dataInfo);

                LoadCompleted?.Invoke(this, result);
                return result;
            }
            catch (Exception ex)
            {
                var result = PointCloudLoadResult.Failure(filePath, ex.Message);
                LoadCompleted?.Invoke(this, result);
                return result;
            }
        }

        /// <summary>
        /// 加载点云数据
        /// </summary>
        private void LoadPointCloudData(string filePath, List<Vector3> points)
        {
            int validPoints = 0;
            int totalLines = 0;
            int skippedLines = 0;

            using (var reader = new StreamReader(filePath))
            {
                string? line;
                int lineNumber = 0;

                while ((line = reader.ReadLine()) != null)
                {
                    lineNumber++;
                    totalLines++;
                    line = line.Trim();

                    // 跳过空行和注释行
                    if (string.IsNullOrEmpty(line) || line.StartsWith("#") || line.StartsWith("//"))
                    {
                        skippedLines++;
                        continue;
                    }

                    try
                    {
                        // 按多种分隔符分割：逗号、空格、制表符
                        var parts = line.Split(new char[] { ',', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                        // 至少需要2列（X, Y），Z是可选的
                        if (parts.Length >= 2)
                        {
                            // 解析X和Y坐标（必需）
                            if (float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out float x) &&
                                float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out float y))
                            {
                                // 解析Z坐标（可选，默认为0）
                                float z = 0;
                                if (parts.Length >= 3 && float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out float zValue))
                                {
                                    z = zValue;
                                }

                                points.Add(new Vector3(x, y, z));
                                validPoints++;
                            }
                            else
                            {
                                skippedLines++;
                            }
                        }
                        else
                        {
                            skippedLines++;
                        }
                    }
                    catch
                    {
                        skippedLines++;
                    }
                }
            }

            Debug.WriteLine($"Parsing complete: {validPoints} valid points from {totalLines} total lines ({skippedLines} skipped)");
        }

        /// <summary>
        /// 分析加载的数据
        /// </summary>
        private string AnalyzeLoadedData(List<Vector3> points, string filePath)
        {
            if (points.Count == 0)
                return "No data";

            // 分析Z坐标分布
            var zValues = points.Select(p => p.Z).ToList();
            var uniqueZCount = zValues.Distinct().Count();
            var minZ = zValues.Min();
            var maxZ = zValues.Max();

            // 判断是2D还是3D数据
            bool is2D = uniqueZCount == 1 && Math.Abs(minZ - maxZ) < 1e-6;

            // 分析坐标范围
            var xRange = points.Max(p => p.X) - points.Min(p => p.X);
            var yRange = points.Max(p => p.Y) - points.Min(p => p.Y);
            var zRange = maxZ - minZ;

            string dataType = is2D ? "2D" : "3D";
            string rangeInfo = is2D ?
                $"Range: X={xRange:F2}, Y={yRange:F2}" :
                $"Range: X={xRange:F2}, Y={yRange:F2}, Z={zRange:F2}";

            // 根据扩展名检测文件格式
            string extension = Path.GetExtension(filePath).ToLower();
            string formatInfo = extension switch
            {
                ".csv" => "CSV format",
                ".txt" => "Text format",
                ".dat" => "Data format",
                _ => "Unknown format"
            };

            return $"{dataType} data • {formatInfo} • {rangeInfo}";
        }
    }

    /// <summary>
    /// STL转换器 - 将点云数据转换为三角网格
    /// </summary>
    public class STLConverter
    {
        /// <summary>
        /// 进度回调委托
        /// </summary>
        public Action<double, string>? ProgressCallback { get; set; }

        /// <summary>
        /// 将点云转换为三角网格
        /// </summary>
        public HelixToolkit.Wpf.SharpDX.MeshGeometry3D ConvertPointCloudToMesh(List<Vector3> points)
        {
            if (points.Count < 3)
            {
                throw new ArgumentException("点云数据不足，无法生成三角网格");
            }

            ProgressCallback?.Invoke(0, "开始点云三角剖分...");

            // 使用优化的网格生成算法
            return CreateOptimizedMesh(points);
        }

        /// <summary>
        /// 创建优化的网格（性能改进版本）
        /// </summary>
        private HelixToolkit.Wpf.SharpDX.MeshGeometry3D CreateOptimizedMesh(List<Vector3> points)
        {
            var mesh = new HelixToolkit.Wpf.SharpDX.MeshGeometry3D();

            ProgressCallback?.Invoke(20, "创建网格结构...");

            // 使用更高效的网格生成算法
            var gridPoints = CreateOptimizedGrid(points);

            ProgressCallback?.Invoke(60, "生成三角形...");

            CreateTrianglesFromGrid(gridPoints, mesh);

            ProgressCallback?.Invoke(90, "计算法向量...");

            return mesh;
        }

        /// <summary>
        /// 创建优化的网格（性能改进）
        /// </summary>
        private Vector3[,] CreateOptimizedGrid(List<Vector3> points)
        {
            // 计算边界
            var minX = points.Min(p => p.X);
            var maxX = points.Max(p => p.X);
            var minY = points.Min(p => p.Y);
            var maxY = points.Max(p => p.Y);

            // 优化网格分辨率计算
            int gridSize = CalculateOptimalGridSize(points.Count);

            ProgressCallback?.Invoke(30, $"创建 {gridSize}x{gridSize} 网格...");

            var grid = new Vector3[gridSize, gridSize];
            var stepX = (maxX - minX) / (gridSize - 1);
            var stepY = (maxY - minY) / (gridSize - 1);

            // 为每个网格点找到最近的点云点
            for (int i = 0; i < gridSize; i++)
            {
                for (int j = 0; j < gridSize; j++)
                {
                    var gridX = minX + i * stepX;
                    var gridY = minY + j * stepY;

                    // 找到最近的点
                    var nearestPoint = FindNearestPoint(points, gridX, gridY);
                    grid[i, j] = nearestPoint;
                }
            }

            return grid;
        }

        /// <summary>
        /// 计算最优网格大小
        /// </summary>
        private int CalculateOptimalGridSize(int pointCount)
        {
            // 根据点数量动态调整网格大小，平衡质量和性能
            if (pointCount < 1000) return 15;
            if (pointCount < 10000) return 25;
            if (pointCount < 50000) return 35;
            if (pointCount < 100000) return 45;
            return Math.Min(60, (int)Math.Sqrt(pointCount / 50)); // 最大60x60网格
        }

        /// <summary>
        /// 找到最近的点
        /// </summary>
        private Vector3 FindNearestPoint(List<Vector3> points, float targetX, float targetY)
        {
            var minDistance = float.MaxValue;
            var nearestPoint = points[0];

            foreach (var point in points)
            {
                var distance = (point.X - targetX) * (point.X - targetX) +
                              (point.Y - targetY) * (point.Y - targetY);

                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestPoint = point;
                }
            }

            return nearestPoint;
        }

        /// <summary>
        /// 从网格创建三角形
        /// </summary>
        private void CreateTrianglesFromGrid(Vector3[,] grid, HelixToolkit.Wpf.SharpDX.MeshGeometry3D mesh)
        {
            if (grid == null || mesh == null)
            {
                throw new ArgumentNullException("网格或mesh对象为空");
            }

            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);

            // 确保集合已初始化
            if (mesh.Positions == null)
                mesh.Positions = new Vector3Collection();
            if (mesh.TriangleIndices == null)
                mesh.TriangleIndices = new IntCollection();

            // 添加顶点
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    mesh.Positions.Add(grid[i, j]);
                }
            }

            // 创建三角形索引
            for (int i = 0; i < rows - 1; i++)
            {
                for (int j = 0; j < cols - 1; j++)
                {
                    int topLeft = i * cols + j;
                    int topRight = i * cols + (j + 1);
                    int bottomLeft = (i + 1) * cols + j;
                    int bottomRight = (i + 1) * cols + (j + 1);

                    // 第一个三角形
                    mesh.TriangleIndices.Add(topLeft);
                    mesh.TriangleIndices.Add(bottomLeft);
                    mesh.TriangleIndices.Add(topRight);

                    // 第二个三角形
                    mesh.TriangleIndices.Add(topRight);
                    mesh.TriangleIndices.Add(bottomLeft);
                    mesh.TriangleIndices.Add(bottomRight);
                }
            }

            // 计算法向量（手动计算）
            CalculateNormals(mesh);
        }

        /// <summary>
        /// 导出STL文件
        /// </summary>
        public void ExportSTL(MeshGeometryModel3D meshModel, string filePath)
        {
            if (meshModel?.Geometry == null)
            {
                throw new ArgumentException("无效的网格模型");
            }

            var mesh = meshModel.Geometry as HelixToolkit.Wpf.SharpDX.MeshGeometry3D;
            if (mesh == null)
            {
                throw new ArgumentException("网格几何体类型不正确");
            }

            ExportSTLAscii(mesh, filePath);
        }

        /// <summary>
        /// 手动计算法向量
        /// </summary>
        private void CalculateNormals(HelixToolkit.Wpf.SharpDX.MeshGeometry3D mesh)
        {
            var normals = new Vector3Collection();
            var positions = mesh.Positions;
            var indices = mesh.TriangleIndices;

            // 初始化法向量数组
            for (int i = 0; i < positions.Count; i++)
            {
                normals.Add(Vector3.Zero);
            }

            // 计算每个三角形的法向量并累加到顶点
            for (int i = 0; i < indices.Count; i += 3)
            {
                var i1 = indices[i];
                var i2 = indices[i + 1];
                var i3 = indices[i + 2];

                var v1 = positions[i1];
                var v2 = positions[i2];
                var v3 = positions[i3];

                var edge1 = v2 - v1;
                var edge2 = v3 - v1;
                var normal = Vector3.Cross(edge1, edge2);

                normals[i1] += normal;
                normals[i2] += normal;
                normals[i3] += normal;
            }

            // 归一化法向量
            for (int i = 0; i < normals.Count; i++)
            {
                normals[i] = Vector3.Normalize(normals[i]);
            }

            mesh.Normals = normals;
        }

        /// <summary>
        /// 导出ASCII格式的STL文件
        /// </summary>
        private void ExportSTLAscii(HelixToolkit.Wpf.SharpDX.MeshGeometry3D mesh, string filePath)
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine("solid PointCloudMesh");

                var positions = mesh.Positions;
                var indices = mesh.TriangleIndices;
                var normals = mesh.Normals;

                for (int i = 0; i < indices.Count; i += 3)
                {
                    var v1 = positions[indices[i]];
                    var v2 = positions[indices[i + 1]];
                    var v3 = positions[indices[i + 2]];

                    // 计算法向量（如果没有预计算的话）
                    Vector3 normal;
                    if (normals != null && normals.Count > indices[i])
                    {
                        normal = normals[indices[i]];
                    }
                    else
                    {
                        var edge1 = v2 - v1;
                        var edge2 = v3 - v1;
                        normal = Vector3.Cross(edge1, edge2);
                        normal = Vector3.Normalize(normal);
                    }

                    writer.WriteLine($"  facet normal {normal.X:F6} {normal.Y:F6} {normal.Z:F6}");
                    writer.WriteLine("    outer loop");
                    writer.WriteLine($"      vertex {v1.X:F6} {v1.Y:F6} {v1.Z:F6}");
                    writer.WriteLine($"      vertex {v2.X:F6} {v2.Y:F6} {v2.Z:F6}");
                    writer.WriteLine($"      vertex {v3.X:F6} {v3.Y:F6} {v3.Z:F6}");
                    writer.WriteLine("    endloop");
                    writer.WriteLine("  endfacet");
                }

                writer.WriteLine("endsolid PointCloudMesh");
            }
        }
    }

    /// <summary>
    /// STL文件加载器
    /// </summary>
    public class STLLoader
    {
        /// <summary>
        /// 加载STL文件
        /// </summary>
        public HelixToolkit.Wpf.SharpDX.MeshGeometry3D LoadSTL(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"STL文件不存在: {filePath}");
            }

            // 检查是ASCII还是二进制格式
            if (IsAsciiSTL(filePath))
            {
                return LoadAsciiSTL(filePath);
            }
            else
            {
                return LoadBinarySTL(filePath);
            }
        }

        /// <summary>
        /// 检查是否为ASCII格式STL
        /// </summary>
        private bool IsAsciiSTL(string filePath)
        {
            using (var reader = new StreamReader(filePath))
            {
                var firstLine = reader.ReadLine()?.Trim().ToLower();
                return firstLine?.StartsWith("solid") == true;
            }
        }

        /// <summary>
        /// 加载ASCII格式STL
        /// </summary>
        private HelixToolkit.Wpf.SharpDX.MeshGeometry3D LoadAsciiSTL(string filePath)
        {
            var mesh = new HelixToolkit.Wpf.SharpDX.MeshGeometry3D();
            if (mesh.Positions == null) mesh.Positions = new Vector3Collection();
            if (mesh.TriangleIndices == null) mesh.TriangleIndices = new IntCollection();
            if (mesh.Normals == null) mesh.Normals = new Vector3Collection();

            using (var reader = new StreamReader(filePath))
            {
                string? line;
                Vector3 normal = Vector3.Zero;
                var vertices = new List<Vector3>();

                while ((line = reader.ReadLine()) != null)
                {
                    line = line.Trim().ToLower();

                    if (line.StartsWith("facet normal"))
                    {
                        var parts = line.Split(' ');
                        if (parts.Length >= 5)
                        {
                            float.TryParse(parts[2], out normal.X);
                            float.TryParse(parts[3], out normal.Y);
                            float.TryParse(parts[4], out normal.Z);
                        }
                        vertices.Clear();
                    }
                    else if (line.StartsWith("vertex"))
                    {
                        var parts = line.Split(' ');
                        if (parts.Length >= 4)
                        {
                            var vertex = new Vector3();
                            float.TryParse(parts[1], out vertex.X);
                            float.TryParse(parts[2], out vertex.Y);
                            float.TryParse(parts[3], out vertex.Z);
                            vertices.Add(vertex);
                        }
                    }
                    else if (line.StartsWith("endfacet") && vertices.Count == 3)
                    {
                        // 添加三角形
                        int startIndex = mesh.Positions.Count;
                        foreach (var vertex in vertices)
                        {
                            mesh.Positions.Add(vertex);
                            mesh.Normals.Add(normal);
                        }

                        mesh.TriangleIndices.Add(startIndex);
                        mesh.TriangleIndices.Add(startIndex + 1);
                        mesh.TriangleIndices.Add(startIndex + 2);
                    }
                }
            }

            return mesh;
        }

        /// <summary>
        /// 加载二进制格式STL
        /// </summary>
        private HelixToolkit.Wpf.SharpDX.MeshGeometry3D LoadBinarySTL(string filePath)
        {
            var mesh = new HelixToolkit.Wpf.SharpDX.MeshGeometry3D();
            if (mesh.Positions == null) mesh.Positions = new Vector3Collection();
            if (mesh.TriangleIndices == null) mesh.TriangleIndices = new IntCollection();
            if (mesh.Normals == null) mesh.Normals = new Vector3Collection();

            using (var reader = new BinaryReader(File.OpenRead(filePath)))
            {
                // 跳过80字节头部
                reader.ReadBytes(80);

                // 读取三角形数量
                uint triangleCount = reader.ReadUInt32();

                for (int i = 0; i < triangleCount; i++)
                {
                    // 读取法向量
                    var normal = new Vector3(reader.ReadSingle(), reader.ReadSingle(), reader.ReadSingle());

                    // 读取三个顶点
                    int startIndex = mesh.Positions.Count;
                    for (int j = 0; j < 3; j++)
                    {
                        var vertex = new Vector3(reader.ReadSingle(), reader.ReadSingle(), reader.ReadSingle());
                        mesh.Positions.Add(vertex);
                        mesh.Normals.Add(normal);
                    }

                    // 添加三角形索引
                    mesh.TriangleIndices.Add(startIndex);
                    mesh.TriangleIndices.Add(startIndex + 1);
                    mesh.TriangleIndices.Add(startIndex + 2);

                    // 跳过属性字节计数
                    reader.ReadUInt16();
                }
            }

            return mesh;
        }
    }

    /// <summary>
    /// 点云可视化器 - 支持流式渲染和高级颜色映射
    /// </summary>
    public class PointCloudVisualizer
    {
        private readonly Viewport3DX viewport;
        private PointGeometryModel3D? currentPointCloud;
        private List<Vector3> accumulatedPoints = new List<Vector3>();

        // 颜色映射模式
        public enum ColorMappingMode
        {
            ByFile,         // 按文件区分颜色
            HeightBased,    // 基于高度(Z坐标) - 科学实用
            Uniform,        // 统一颜色 - 简洁清晰
            Grayscale       // 灰度渐变 - 专业显示
        }

        public ColorMappingMode CurrentColorMode { get; set; } = ColorMappingMode.ByFile;
        private List<PointCloudFileInfo> fileInfoList = new List<PointCloudFileInfo>();
        private List<int> pointFileIndices = new List<int>();

        public PointCloudVisualizer(Viewport3DX viewport)
        {
            this.viewport = viewport;
        }

        /// <summary>
        /// 设置文件信息列表，用于按文件着色
        /// </summary>
        public void SetFileInfoList(List<PointCloudFileInfo> fileInfos)
        {
            this.fileInfoList = fileInfos;
        }

        /// <summary>
        /// 设置点的文件索引信息
        /// </summary>
        public void SetPointFileIndices(List<int> fileIndices)
        {
            this.pointFileIndices = fileIndices;
        }

        /// <summary>
        /// 创建点云可视化
        /// </summary>
        public void CreatePointCloudVisualization(List<Vector3> points)
        {
            // 重置累积点
            accumulatedPoints.Clear();
            accumulatedPoints.AddRange(points);

            // 移除现有的点云
            if (currentPointCloud != null)
            {
                viewport.Items.Remove(currentPointCloud);
            }

            if (points.Count == 0)
                return;

            // 创建点几何体
            var pointGeometry = new PointGeometry3D();
            pointGeometry.Positions = new Vector3Collection(points);

            // 为点创建颜色
            var colors = CreateColorMapping(points);
            pointGeometry.Colors = colors;

            // 调试：验证颜色数据
            DebugColorMapping(colors, points.Count);

            // 创建点云模型 - 基于WPF_PCV的工作实现
            currentPointCloud = new PointGeometryModel3D
            {
                Geometry = pointGeometry,
                // 根据WPF_PCV的实现，需要设置Color属性作为基础色
                // Colors集合会在此基础上进行调制
                Color = System.Windows.Media.Colors.White,
                Size = new Size(CalculatePointSize(points.Count), CalculatePointSize(points.Count))
            };

            // 添加到视口
            viewport.Items.Add(currentPointCloud);

            // 调整相机以适应点云
            AdjustCameraToFitPointCloud(points);
        }

        /// <summary>
        /// 增量添加点到现有可视化
        /// </summary>
        public void AddPointsToVisualization(List<Vector3> newPoints)
        {
            if (newPoints.Count == 0) return;

            // 添加到累积点
            accumulatedPoints.AddRange(newPoints);

            // 更新现有点云几何体
            if (currentPointCloud?.Geometry is PointGeometry3D geometry)
            {
                geometry.Positions = new Vector3Collection(accumulatedPoints);
                geometry.Colors = CreateColorMapping(accumulatedPoints);

                // 更新点大小
                currentPointCloud.Size = new Size(CalculatePointSize(accumulatedPoints.Count), CalculatePointSize(accumulatedPoints.Count));
            }
        }

        /// <summary>
        /// 创建颜色映射 - 支持多种颜色方案
        /// </summary>
        private Color4Collection CreateColorMapping(List<Vector3> points)
        {
            var colors = new Color4Collection();
            if (points.Count == 0) return colors;

            System.Diagnostics.Debug.WriteLine($"[颜色映射] 开始创建颜色映射，模式: {CurrentColorMode}, 点数: {points.Count}");

            switch (CurrentColorMode)
            {
                case ColorMappingMode.ByFile:
                    return CreateFileBasedColors(points);
                case ColorMappingMode.HeightBased:
                    return CreateHeightBasedColors(points);
                case ColorMappingMode.Uniform:
                    return CreateUniformColors(points);
                case ColorMappingMode.Grayscale:
                    return CreateGrayscaleColors(points);
                default:
                    return CreateFileBasedColors(points); // 默认按文件着色
            }
        }

        /// <summary>
        /// 按文件创建颜色 - 不同文件使用不同颜色便于区分
        /// </summary>
        private Color4Collection CreateFileBasedColors(List<Vector3> points)
        {
            var colors = new Color4Collection();

            // 定义一组易于区分的颜色
            var fileColors = new Color4[]
            {
                new Color4(0.2f, 0.6f, 1.0f, 1.0f),  // 蓝色
                new Color4(1.0f, 0.4f, 0.2f, 1.0f),  // 橙色
                new Color4(0.2f, 0.8f, 0.3f, 1.0f),  // 绿色
                new Color4(0.9f, 0.2f, 0.6f, 1.0f),  // 粉红色
                new Color4(0.7f, 0.5f, 0.9f, 1.0f),  // 紫色
                new Color4(1.0f, 0.8f, 0.2f, 1.0f),  // 黄色
                new Color4(0.3f, 0.9f, 0.9f, 1.0f),  // 青色
                new Color4(0.9f, 0.6f, 0.3f, 1.0f),  // 棕色
                new Color4(0.5f, 0.5f, 0.5f, 1.0f),  // 灰色
                new Color4(0.8f, 0.2f, 0.2f, 1.0f)   // 红色
            };

            // 为每个点分配对应文件的颜色
            for (int i = 0; i < points.Count; i++)
            {
                int fileIndex = GetFileIndexForPoint(i);
                var color = fileColors[fileIndex % fileColors.Length];
                colors.Add(color);
            }

            // 调试信息：统计每个文件的点数
            var filePointCounts = new Dictionary<int, int>();
            for (int i = 0; i < points.Count; i++)
            {
                int fileIndex = GetFileIndexForPoint(i);
                filePointCounts[fileIndex] = filePointCounts.GetValueOrDefault(fileIndex, 0) + 1;
            }

            System.Diagnostics.Debug.WriteLine($"[按文件着色] 生成了 {colors.Count} 个颜色，涉及 {fileInfoList.Count} 个文件");
            System.Diagnostics.Debug.WriteLine($"[按文件着色] 点文件索引数组长度: {pointFileIndices.Count}");
            foreach (var kvp in filePointCounts)
            {
                var fileName = kvp.Key < fileInfoList.Count ? fileInfoList[kvp.Key].FileName : $"文件{kvp.Key}";
                System.Diagnostics.Debug.WriteLine($"[按文件着色] 文件 {kvp.Key} ({fileName}): {kvp.Value} 个点");
            }

            return colors;
        }

        /// <summary>
        /// 基于高度(Z坐标)的颜色映射 - 科学实用的高度可视化
        /// </summary>
        private Color4Collection CreateHeightBasedColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            if (points.Count == 0) return colors;

            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);
            var zRange = maxZ - minZ;

            foreach (var point in points)
            {
                var normalizedZ = zRange > 0 ? (point.Z - minZ) / zRange : 0.5f;

                // 使用更科学的蓝-绿-黄-红渐变，类似地形图
                float r, g, b;
                if (normalizedZ < 0.25f)
                {
                    // 深蓝到浅蓝
                    r = 0.0f;
                    g = 0.2f + normalizedZ * 1.2f;
                    b = 0.8f + normalizedZ * 0.8f;
                }
                else if (normalizedZ < 0.5f)
                {
                    // 浅蓝到绿色
                    var t = (normalizedZ - 0.25f) * 4.0f;
                    r = 0.0f;
                    g = 0.5f + t * 0.5f;
                    b = 1.0f - t * 0.7f;
                }
                else if (normalizedZ < 0.75f)
                {
                    // 绿色到黄色
                    var t = (normalizedZ - 0.5f) * 4.0f;
                    r = t * 0.8f;
                    g = 1.0f;
                    b = 0.3f - t * 0.3f;
                }
                else
                {
                    // 黄色到红色
                    var t = (normalizedZ - 0.75f) * 4.0f;
                    r = 0.8f + t * 0.2f;
                    g = 1.0f - t * 0.6f;
                    b = 0.0f;
                }

                colors.Add(new Color4(r, g, b, 1.0f));
            }

            System.Diagnostics.Debug.WriteLine($"[高度着色] Z范围: {minZ:F3} - {maxZ:F3}, 生成 {colors.Count} 个颜色");
            return colors;
        }

        /// <summary>
        /// 统一颜色 - 简洁清晰的单色显示
        /// </summary>
        private Color4Collection CreateUniformColors(List<Vector3> points)
        {
            var colors = new Color4Collection();

            // 使用专业的中性蓝色，对眼睛友好且易于观察细节
            var uniformColor = new Color4(0.3f, 0.6f, 0.9f, 1.0f);

            for (int i = 0; i < points.Count; i++)
            {
                colors.Add(uniformColor);
            }

            System.Diagnostics.Debug.WriteLine($"[统一颜色] 生成 {colors.Count} 个统一颜色");
            return colors;
        }

        /// <summary>
        /// 灰度渐变 - 专业的黑白显示，突出形状和结构
        /// </summary>
        private Color4Collection CreateGrayscaleColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            if (points.Count == 0) return colors;

            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);
            var zRange = maxZ - minZ;

            foreach (var point in points)
            {
                var normalizedZ = zRange > 0 ? (point.Z - minZ) / zRange : 0.5f;

                // 从深灰到浅灰的渐变，保持良好的对比度
                var intensity = 0.2f + normalizedZ * 0.6f; // 0.2到0.8的范围
                colors.Add(new Color4(intensity, intensity, intensity, 1.0f));
            }

            System.Diagnostics.Debug.WriteLine($"[灰度渐变] Z范围: {minZ:F3} - {maxZ:F3}, 生成 {colors.Count} 个颜色");
            return colors;
        }

        /// <summary>
        /// 获取指定点索引对应的文件索引
        /// </summary>
        private int GetFileIndexForPoint(int pointIndex)
        {
            // 如果有预计算的文件索引，直接使用
            if (pointFileIndices.Count > pointIndex && pointIndex >= 0)
            {
                return pointFileIndices[pointIndex];
            }

            // 回退到原始方法（用于兼容性）
            if (fileInfoList.Count == 0) return 0;

            int currentIndex = 0;
            for (int i = 0; i < fileInfoList.Count; i++)
            {
                if (pointIndex >= currentIndex && pointIndex < currentIndex + fileInfoList[i].PointCount)
                {
                    return i;
                }
                currentIndex += fileInfoList[i].PointCount;
            }

            return fileInfoList.Count - 1; // 默认返回最后一个文件的索引
        }

        /// <summary>
        /// 调试颜色映射数据
        /// </summary>
        private void DebugColorMapping(Color4Collection colors, int pointCount)
        {
            if (colors == null || colors.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine($"[颜色调试] 警告: 颜色集合为空或null");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[颜色调试] 点数: {pointCount}, 颜色数: {colors.Count}");
            System.Diagnostics.Debug.WriteLine($"[颜色调试] 颜色模式: {CurrentColorMode}");

            // 输出前几个颜色值
            for (int i = 0; i < Math.Min(5, colors.Count); i++)
            {
                var color = colors[i];
                System.Diagnostics.Debug.WriteLine($"[颜色调试] 颜色[{i}]: R={color.Red:F3}, G={color.Green:F3}, B={color.Blue:F3}, A={color.Alpha:F3}");
            }

            // 检查是否所有颜色都是黑色
            bool allBlack = colors.All(c => c.Red == 0 && c.Green == 0 && c.Blue == 0);
            if (allBlack)
            {
                System.Diagnostics.Debug.WriteLine($"[颜色调试] 警告: 所有颜色都是黑色！");
            }

            // 检查颜色范围
            var minR = colors.Min(c => c.Red);
            var maxR = colors.Max(c => c.Red);
            var minG = colors.Min(c => c.Green);
            var maxG = colors.Max(c => c.Green);
            var minB = colors.Min(c => c.Blue);
            var maxB = colors.Max(c => c.Blue);

            System.Diagnostics.Debug.WriteLine($"[颜色调试] 颜色范围 - R:[{minR:F3}-{maxR:F3}], G:[{minG:F3}-{maxG:F3}], B:[{minB:F3}-{maxB:F3}]");
        }

        /// <summary>
        /// 根据点数计算点大小
        /// </summary>
        private double CalculatePointSize(int pointCount)
        {
            if (pointCount > 100000) return 1.0;
            if (pointCount > 50000) return 1.5;
            if (pointCount > 10000) return 2.0;
            return 2.5;
        }

        /// <summary>
        /// 调整相机以适应点云
        /// </summary>
        private void AdjustCameraToFitPointCloud(List<Vector3> points)
        {
            if (points.Count == 0)
                return;

            // 计算边界框
            var minX = points.Min(p => p.X);
            var maxX = points.Max(p => p.X);
            var minY = points.Min(p => p.Y);
            var maxY = points.Max(p => p.Y);
            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);

            var center = new Vector3(
                (minX + maxX) / 2,
                (minY + maxY) / 2,
                (minZ + maxZ) / 2
            );

            var size = Math.Max(Math.Max(maxX - minX, maxY - minY), maxZ - minZ);
            var distance = size * 2.0;
            
            if (viewport.Camera is HelixToolkit.Wpf.SharpDX.OrthographicCamera orthographicCamera)
            {
                orthographicCamera.Position = new Point3D(center.X + distance, center.Y + distance, center.Z + distance);
                orthographicCamera.LookDirection = new Vector3D(-distance, -distance, -distance);
                orthographicCamera.Width = size * 1.2; // 设置正交相机的视野宽度
            }
        }
    }
    #endregion
}