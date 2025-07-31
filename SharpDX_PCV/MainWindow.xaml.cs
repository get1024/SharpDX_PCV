using HelixToolkit.Wpf.SharpDX;
using System.Windows;
using System.Windows.Media.Media3D;
using System.Windows.Controls;
using SharpDX;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;

namespace SharpDX_PCV
{
    /// <summary>
    /// 点云查看器主窗口 - 负责文件加载和点云可视化
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 私有字段
        private readonly List<Vector3> pointCloudData = new List<Vector3>();
        private readonly List<Vector3> fullPointCloudData = new List<Vector3>(); // 完整数据
        private readonly List<Vector3> renderPointCloudData = new List<Vector3>(); // 渲染数据
        private readonly List<string> loadedFilePaths = new List<string>(); // 已加载文件路径
        private PointGeometryModel3D? currentPointCloud;
        private readonly PointCloudFileLoader fileLoader;
        private readonly PointCloudVisualizer visualizer;

        // 降采样控制
        private int downsampleLevel = 1;

        // 流式渲染控制
        private bool isStreamingMode = false;
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
            visualizer.CurrentColorMode = PointCloudVisualizer.ColorMappingMode.Rainbow;

            // 临时移除事件处理器，设置选中项，然后重新添加
            ColorModeComboBox.SelectionChanged -= ColorModeComboBox_SelectionChanged;
            ColorModeComboBox.SelectedIndex = 0; // 彩虹色谱
            ColorModeComboBox.SelectionChanged += ColorModeComboBox_SelectionChanged;
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
        #endregion

        #region 事件处理器
        /// <summary>
        /// 加载文件按钮点击事件 - 支持多文件选择
        /// </summary>
        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select Point Cloud File(s)",
                Filter = "Point Cloud files (*.txt;*.csv;*.dat)|*.txt;*.csv;*.dat|Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*",
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
        private void HandleMultipleFileSelection(string[] selectedFiles)
        {
            if (selectedFiles.Length == 0) return;

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

            // 加载选中的文件
            LoadMultipleFilesAsync(selectedFiles);
        }

        /// <summary>
        /// 清空所有数据
        /// </summary>
        private void ClearAllData()
        {
            pointCloudData.Clear();
            fullPointCloudData.Clear();
            renderPointCloudData.Clear();
            loadedFilePaths.Clear();

            // 重置全局变换
            ResetGlobalTransform();

            // 清空3D视图
            if (currentPointCloud != null)
            {
                helixViewport.Items.Remove(currentPointCloud);
                currentPointCloud = null;
            }

            txtStatus.Text = "已清空数据";
        }

        /// <summary>
        /// 文件加载完成事件处理器
        /// </summary>
        private void OnLoadCompleted(object? sender, PointCloudLoadResult result)
        {
            Dispatcher.Invoke(() =>
            {
                if (result.IsSuccess)
                {
                    HandleSuccessfulLoad(result);
                }
                else
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
        private async void LoadMultipleFilesAsync(string[] filePaths)
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
                        txtStatus.Text = $"加载文件 {loadedFiles + 1}/{totalFiles}: {Path.GetFileName(filePath)}";

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
                            allNewPoints.AddRange(loadResult.Points);
                            loadedFilePaths.Add(filePath);
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
                    pointCloudData.AddRange(allNewPoints);

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
        /// 处理成功加载的结果
        /// </summary>
        private void HandleSuccessfulLoad(PointCloudLoadResult result)
        {
            // 更新点云数据
            pointCloudData.Clear();
            pointCloudData.AddRange(result.Points);

            // 创建可视化
            visualizer.CreatePointCloudVisualization(result.Points);

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
            isStreamingMode = true;
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
                var downsampledData = await Task.Run(() => VoxelGridDownsample(transformedData, downsampleLevel), token);
                renderPointCloudData.Clear();
                renderPointCloudData.AddRange(downsampledData);

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
                var downsampledData = await Task.Run(() => VoxelGridDownsample(transformedData, downsampleLevel));
                renderPointCloudData.Clear();
                renderPointCloudData.AddRange(downsampledData);

                // 4. 直接渲染
                visualizer.CreatePointCloudVisualization(renderPointCloudData);

                var reductionRatio = fullPointCloudData.Count > 0 ? (double)renderPointCloudData.Count / fullPointCloudData.Count : 1.0;
                var transformInfo = globalTransformCalculated ? $" | 已应用坐标变换 (缩放: {globalScale:F3})" : "";
                txtStatus.Text = $"渲染完成: {renderPointCloudData.Count:N0}/{fullPointCloudData.Count:N0} 点 ({reductionRatio:P1}){transformInfo}";
            }
            catch (OutOfMemoryException ex)
            {
                HandleLoadError("渲染时内存不足", ex);
            }
            catch (Exception ex)
            {
                HandleLoadError("渲染发生错误", ex);
            }
        }



        /// <summary>
        /// 清理流式渲染资源
        /// </summary>
        private void CleanupStreamingResources()
        {
            isStreamingMode = false;
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
                    var downsampledData = VoxelGridDownsample(transformedData, downsampleLevel);

                    Dispatcher.Invoke(() =>
                    {
                        renderPointCloudData.Clear();
                        renderPointCloudData.AddRange(downsampledData);

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
        /// 体素网格降采样算法
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

        #endregion

        #endregion

    }

    #region 辅助类定义
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
            HeightBased,    // 基于高度(Z坐标)
            DepthBased,     // 基于深度(距离相机)
            DensityBased,   // 基于点云密度
            Rainbow,        // 彩虹色谱
            Thermal         // 热力图色谱
        }

        public ColorMappingMode CurrentColorMode { get; set; } = ColorMappingMode.Rainbow;

        public PointCloudVisualizer(Viewport3DX viewport)
        {
            this.viewport = viewport;
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

            // 创建点云模型
            currentPointCloud = new PointGeometryModel3D
            {
                Geometry = pointGeometry,
                // 不设置Color属性，让HelixToolkit使用Colors集合
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

            // 临时测试：使用固定的鲜艳颜色来验证颜色系统
            System.Diagnostics.Debug.WriteLine($"[颜色映射] 开始创建颜色映射，模式: {CurrentColorMode}, 点数: {points.Count}");

            switch (CurrentColorMode)
            {
                case ColorMappingMode.HeightBased:
                    return CreateHeightBasedColors(points);
                case ColorMappingMode.DepthBased:
                    return CreateDepthBasedColors(points);
                case ColorMappingMode.DensityBased:
                    return CreateDensityBasedColors(points);
                case ColorMappingMode.Rainbow:
                    return CreateRainbowColors(points);
                case ColorMappingMode.Thermal:
                    return CreateThermalColors(points);
                default:
                    return CreateRainbowColors(points);
            }
        }

        /// <summary>
        /// 创建测试颜色 - 用于调试
        /// </summary>
        private Color4Collection CreateTestColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            var colorList = new Color4[]
            {
                new Color4(1.0f, 0.0f, 0.0f, 1.0f), // 红色
                new Color4(0.0f, 1.0f, 0.0f, 1.0f), // 绿色
                new Color4(0.0f, 0.0f, 1.0f, 1.0f), // 蓝色
                new Color4(1.0f, 1.0f, 0.0f, 1.0f), // 黄色
                new Color4(1.0f, 0.0f, 1.0f, 1.0f), // 洋红
                new Color4(0.0f, 1.0f, 1.0f, 1.0f)  // 青色
            };

            for (int i = 0; i < points.Count; i++)
            {
                colors.Add(colorList[i % colorList.Length]);
            }

            System.Diagnostics.Debug.WriteLine($"[测试颜色] 生成了 {colors.Count} 个测试颜色");
            return colors;
        }

        /// <summary>
        /// 基于高度(Z坐标)的颜色映射
        /// </summary>
        private Color4Collection CreateHeightBasedColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);
            var zRange = maxZ - minZ;

            foreach (var point in points)
            {
                var normalizedZ = zRange > 0 ? (point.Z - minZ) / zRange : 0.5f;

                // 蓝色(低) -> 绿色(中) -> 红色(高)
                float r, g, b;
                if (normalizedZ < 0.5f)
                {
                    // 蓝色到绿色
                    r = 0.0f;
                    g = normalizedZ * 2.0f;
                    b = 1.0f - normalizedZ * 2.0f;
                }
                else
                {
                    // 绿色到红色
                    r = (normalizedZ - 0.5f) * 2.0f;
                    g = 1.0f - (normalizedZ - 0.5f) * 2.0f;
                    b = 0.0f;
                }

                colors.Add(new Color4(r, g, b, 1.0f));
            }

            return colors;
        }

        /// <summary>
        /// 基于深度(距离相机)的颜色映射
        /// </summary>
        private Color4Collection CreateDepthBasedColors(List<Vector3> points)
        {
            var colors = new Color4Collection();

            // 计算每个点到原点的距离作为深度
            var distances = points.Select(p => (float)Math.Sqrt(p.X * p.X + p.Y * p.Y + p.Z * p.Z)).ToList();
            var minDist = distances.Min();
            var maxDist = distances.Max();
            var distRange = maxDist - minDist;

            for (int i = 0; i < points.Count; i++)
            {
                var normalizedDist = distRange > 0 ? (distances[i] - minDist) / distRange : 0.5f;

                // 近处亮色，远处暗色，增强深度感
                var intensity = 1.0f - normalizedDist * 0.7f; // 保持一定亮度
                colors.Add(new Color4(0.2f + intensity * 0.8f, 0.4f + intensity * 0.6f, 0.8f + intensity * 0.2f, 1.0f));
            }

            return colors;
        }

        /// <summary>
        /// 基于点云密度的颜色映射 - 优化版本
        /// </summary>
        private Color4Collection CreateDensityBasedColors(List<Vector3> points)
        {
            var colors = new Color4Collection();

            // 对于大数据集，使用简化的密度估算
            if (points.Count > 10000)
            {
                return CreateSimplifiedDensityColors(points);
            }

            const float searchRadius = 2.0f;
            var densities = new List<float>();

            // 使用空间分割优化密度计算
            for (int i = 0; i < points.Count; i++)
            {
                var currentPoint = points[i];
                int neighborCount = 0;

                // 只检查附近的点，减少计算量
                for (int j = Math.Max(0, i - 100); j < Math.Min(points.Count, i + 100); j++)
                {
                    if (i == j) continue;

                    var distance = Vector3.Distance(currentPoint, points[j]);
                    if (distance <= searchRadius)
                    {
                        neighborCount++;
                    }
                }

                densities.Add(neighborCount);
            }

            var minDensity = densities.Min();
            var maxDensity = densities.Max();
            var densityRange = maxDensity - minDensity;

            for (int i = 0; i < points.Count; i++)
            {
                var normalizedDensity = densityRange > 0 ? (densities[i] - minDensity) / densityRange : 0.5f;

                // 低密度紫色，高密度黄色
                colors.Add(new Color4(
                    0.5f + normalizedDensity * 0.5f,  // R: 紫到黄
                    normalizedDensity * 0.8f,         // G: 增强黄色
                    1.0f - normalizedDensity,         // B: 紫到黄
                    1.0f));
            }

            return colors;
        }

        /// <summary>
        /// 简化的密度颜色映射（用于大数据集）
        /// </summary>
        private Color4Collection CreateSimplifiedDensityColors(List<Vector3> points)
        {
            var colors = new Color4Collection();

            // 基于点的索引位置创建伪密度效果
            for (int i = 0; i < points.Count; i++)
            {
                var normalizedIndex = (float)i / points.Count;

                // 创建波浪状密度效果
                var density = (float)(Math.Sin(normalizedIndex * Math.PI * 8) * 0.5 + 0.5);

                colors.Add(new Color4(
                    0.5f + density * 0.5f,
                    density * 0.8f,
                    1.0f - density,
                    1.0f));
            }

            return colors;
        }

        /// <summary>
        /// 彩虹色谱颜色映射
        /// </summary>
        private Color4Collection CreateRainbowColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            if (points.Count == 0) return colors;

            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);
            var zRange = maxZ - minZ;

            System.Diagnostics.Debug.WriteLine($"[彩虹颜色] Z范围: {minZ:F3} - {maxZ:F3}, 范围: {zRange:F3}");

            foreach (var point in points)
            {
                var normalizedZ = zRange > 0 ? (point.Z - minZ) / zRange : 0.5f;

                // HSV到RGB的转换，创建彩虹效果
                var hue = normalizedZ * 300.0f; // 0-300度，避免回到红色
                var color = HsvToRgb(hue, 1.0f, 1.0f);
                colors.Add(new Color4(color.R, color.G, color.B, 1.0f));
            }

            System.Diagnostics.Debug.WriteLine($"[彩虹颜色] 生成了 {colors.Count} 个颜色");
            return colors;
        }

        /// <summary>
        /// 热力图色谱颜色映射
        /// </summary>
        private Color4Collection CreateThermalColors(List<Vector3> points)
        {
            var colors = new Color4Collection();
            var minZ = points.Min(p => p.Z);
            var maxZ = points.Max(p => p.Z);
            var zRange = maxZ - minZ;

            foreach (var point in points)
            {
                var normalizedZ = zRange > 0 ? (point.Z - minZ) / zRange : 0.5f;

                // 热力图：黑色 -> 红色 -> 黄色 -> 白色
                float r, g, b;
                if (normalizedZ < 0.33f)
                {
                    // 黑色到红色
                    r = normalizedZ * 3.0f;
                    g = 0.0f;
                    b = 0.0f;
                }
                else if (normalizedZ < 0.66f)
                {
                    // 红色到黄色
                    r = 1.0f;
                    g = (normalizedZ - 0.33f) * 3.0f;
                    b = 0.0f;
                }
                else
                {
                    // 黄色到白色
                    r = 1.0f;
                    g = 1.0f;
                    b = (normalizedZ - 0.66f) * 3.0f;
                }

                colors.Add(new Color4(r, g, b, 1.0f));
            }

            return colors;
        }

        /// <summary>
        /// HSV到RGB颜色转换
        /// </summary>
        private (float R, float G, float B) HsvToRgb(float h, float s, float v)
        {
            h = h % 360.0f;
            var c = v * s;
            var x = c * (1 - Math.Abs((h / 60.0f) % 2 - 1));
            var m = v - c;

            float r, g, b;
            if (h < 60) { r = c; g = x; b = 0; }
            else if (h < 120) { r = x; g = c; b = 0; }
            else if (h < 180) { r = 0; g = c; b = x; }
            else if (h < 240) { r = 0; g = x; b = c; }
            else if (h < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }

            return ((float)(r + m), (float)(g + m), (float)(b + m));
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