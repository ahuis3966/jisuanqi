using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Newtonsoft.Json;
using IniParser;
using IniParser.Model;
using NModbus;
using System.Windows.Media;
using Newtonsoft.Json.Linq;
using System.Net.Sockets;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Input;
using System.Text.RegularExpressions;
using System.Windows.Controls.Primitives;

namespace totalizer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private TabControl tabControl1;
        private System.Windows.Threading.DispatcherTimer timer; // 用于显示时间的定时器
        private System.Windows.Threading.DispatcherTimer loadTimer; // 从ini配置读取json文件加载到界面
        private System.Windows.Threading.DispatcherTimer readTimer; // 从ini配置读取json文件将从机的内容更新到json
        private System.Windows.Threading.DispatcherTimer connectionTimer;
        private string formTitle ; // 从.ini文件读取的窗体标题
        private ToolTip toolTip; // 声明 ToolTip 控件
        private int debug_mode; // 调试模式:0否,1是
        private List<IModbusMaster> modbusMasters = new List<IModbusMaster>(); // 存储 Modbus 连接
        private static bool isLogging = false; // 静态变量指示是否正在写日志
        private static readonly object logLock = new object(); // 锁对象，防止多线程并发写日志
        private bool isCheckingConnection = false;
        private IModbusMaster modbusMaster = null;
        private bool isClosing = false; // 标志变量

        public MainWindow()
        {
            InitializeComponent();

            toolTip = new ToolTip();

            var parser = new FileIniDataParser();
            IniData data = null; // 确保 data 被声明为 null
            try
            {
                using (var reader = new StreamReader("config.ini", Encoding.UTF8))
                {
                    data = parser.ReadData(reader);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"加载ini文件错误: {ex.Message}");
            }

            formTitle = data["Settings"]["title"];
            int loadInterval = int.Parse(data["Settings"]["load_interval"]);
            int readInterval = int.Parse(data["Settings"]["read_interval"]);
            debug_mode = int.Parse(data["Settings"]["debug_mode"]);

            var hostInfo = new Dictionary<string, string>();

            // 获取host信息
            for (int i = 0; ; i++)
            {
                var hostKey = $"host[{i}]";
                if (!data["Settings"].ContainsKey(hostKey))
                    break;

                var h = data["Settings"][hostKey].Split(',');
                if (h.Length == 4)
                {
                    string ip = h[0];
                    string port = h[1];
                    string slaveId = h[2];
                    string name = h[3];
                    hostInfo[ip] = name; // 将IP和名称存入字典
                }
            }

            CreateTabControl(hostInfo); // 将host信息传递给CreateTabControl

            LoadJsonAndBuildUI(data);

            // 创建并设置定时器
            timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1); // 每秒更新一次
            timer.Tick += Timer_Tick;
            timer.Start(); // 启动定时器

            // 设置定时器更新 JSON 数据
            loadTimer = new System.Windows.Threading.DispatcherTimer();
            loadTimer.Interval = TimeSpan.FromSeconds(loadInterval);
            loadTimer.Tick += (s, e) => LoadJsonAndBuildUI(data);
            loadTimer.Start();

            // 设置定时器读取从机数据
            readTimer = new System.Windows.Threading.DispatcherTimer();
            readTimer.Interval = TimeSpan.FromSeconds(readInterval);
            readTimer.Tick += (s, e) => ReadSlaveToJson(data); // 需要实现该方法
            readTimer.Start();

            this.Closing += MainWindow_Closing;

            WriteLog("启动完成!");

            // 设置初始窗体标题
            UpdateWindowTitle();
        }

        // 定时器事件处理
        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateWindowTitle();
        }

        // 更新窗体标题的方法
        private void UpdateWindowTitle()
        {
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // 格式化当前日期时间
            this.Title = $"{formTitle} - {currentTime}"; // 设置窗体标题
        }

        //创建选项卡
        private void CreateTabControl(Dictionary<string, string> hostInfo)
        {
            // 清空现有的 TabItem
            tabControl.Items.Clear();

            // 动态添加选项卡
            foreach (var kvp in hostInfo)
            {
                string ip = kvp.Key; // IP地址
                string name = kvp.Value; // 名称或其他信息

                TabItem tabItem = new TabItem
                {
                    Header = ip, // 设置选项卡的标题为IP地址
                    Content = new TextBlock { Text = $"这是来自 {name} ({ip}) 的内容。" } // 选项卡的内容
                };

                tabControl.Items.Add(tabItem);
            }
        }

        //选项卡改变事件
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 遍历所有 TabItem
            foreach (TabItem tabItem in tabControl.Items)
            {
                // 重置所有选项卡的背景颜色
                tabItem.Background = Brushes.Gainsboro;
            }

            // 确保选中的 TabItem 背景颜色不同
            if (tabControl.SelectedItem is TabItem selectedItem)
            {
                selectedItem.Background = Brushes.LightGray; // 选中选项卡背景颜色
            }
        }

        //允许拖动窗口
        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                this.DragMove(); // 允许拖动窗口
            }
        }

        // 将从json得到的内容处理加载到选项卡上
        private void LoadJsonAndBuildUI(IniData data)
        {
            // 记录当前选中项的索引
            int previousSelectedIndex = tabControl.SelectedIndex;

            ClearTabControl();
            tabControl.Items.Clear();

            for (int i = 0; ; i++)
            {
                var hostKey = $"host[{i}]";
                if (!data["Settings"].ContainsKey(hostKey))
                    break;

                var hostInfo = data["Settings"][hostKey].Split(',');
                if (hostInfo.Length == 4)
                {
                    string ip = hostInfo[0];
                    string port = hostInfo[1];
                    string slaveId = hostInfo[2];
                    string name = hostInfo[3];

                    // 创建选项卡
                    TabItem tabItem = new TabItem
                    {
                        Header = ip // 选项卡的标题为IP地址
                    };

                    // 创建主 StackPanel 用于包含整个内容
                    StackPanel mainStackPanel = new StackPanel
                    {
                        Margin = new Thickness(5) // 设置一些边距
                    };

                    // 创建水平 StackPanel 用于放置 Label 和状态 Label
                    StackPanel horizontalStackPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal
                    };

                    // 创建 label 显示从机信息
                    Label label = new Label
                    {
                        Content = $"从机IP: {ip}  端口号: {port}  从机ID: {slaveId}",
                        FontSize = 14
                    };

                    // 创建 statusLabel 显示连接状态
                    Label statusLabel = new Label
                    {
                        Content = "连接状态: 连接中", // 初始文本
                        FontSize = 14
                    };

                    // 将两个 Label 添加到水平 StackPanel 中
                    horizontalStackPanel.Children.Add(label);
                    horizontalStackPanel.Children.Add(statusLabel);

                    // 将水平 StackPanel 添加到主 StackPanel
                    mainStackPanel.Children.Add(horizontalStackPanel);

                    // 将主 StackPanel 添加到选项卡
                    tabItem.Content = mainStackPanel;

                    // 读取 JSON 数据
                    string jsonFileName = $"{ip}.json";
                    if (!File.Exists(jsonFileName))
                    {
                        File.Copy("original.json", jsonFileName, true);
                    }

                    // 读取并解析 JSON 文件
                    List<ModbusCategory> listdata = null;
                    try
                    {
                        var jsonData = File.ReadAllText(jsonFileName);
                        listdata = JsonConvert.DeserializeObject<List<ModbusCategory>>(jsonData);
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"IP:{ip}, JSON 加载失败: {ex.Message}");
                        MessageBox.Show($"无法加载文件 {jsonFileName}，请检查文件是否存在并格式正确。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        continue; // 如果加载失败，跳过此选项卡
                    }

                    // 动态生成控件并添加到界面
                    GenerateControls(ip, port, slaveId, listdata, tabItem);

                    // 将选项卡添加到选项卡控件中
                    tabControl.Items.Add(tabItem);

                    // 启动定时器检测连接状态
                    InitializeConnectionTimer(ip, port, slaveId, statusLabel);
                }
            }

            // 重新设置为先前选中的选项卡
            if (previousSelectedIndex >= 0 && previousSelectedIndex < tabControl.Items.Count)
            {
                tabControl.SelectedIndex = previousSelectedIndex;
            }

            // 创建并添加时间标签
            if (timeLabel == null)
            {
                timeLabel = new Label
                {
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    FontSize = 10,
                    Height = 30 // 时间标签高度
                };
                mainGrid.Children.Add(timeLabel); // 假设你的主窗口有一个名为 mainGrid 的 Grid
            }

            // 完成加载后设置默认第1个选项卡的背景色
            if (tabControl.Items.Count > 0)
            {
                var firstTabItem = tabControl.Items[0] as TabItem;
                if (firstTabItem != null)
                {
                    firstTabItem.Background = Brushes.LightGray; // 设置第一个选项卡背景色
                }
            }
        }

        private void GenerateControls(string ip, string port, string slaveId, List<ModbusCategory> data, TabItem tabItem)
        {
            // 假设你已经在 LoadJsonAndBuildUI 中创建了 stackPanel
            var stackPanel = tabItem.Content as StackPanel;
            if (stackPanel == null)
            {
                // 如果 stackPanel 为 null，则重新创建
                stackPanel = new StackPanel();
                tabItem.Content = stackPanel;
            }

            foreach (var category in data)
            {
                if (category.IsDisplay == 0)
                    continue;

                // 创建一个 Border 作为类别面板
                var categoryBorder = new Border
                {
                    Margin = new Thickness(0, 10, 0, 0),
                    Padding = new Thickness(10),
                    BorderBrush = Brushes.DarkGray,
                    BorderThickness = new Thickness(1),
                    Background = Brushes.Transparent
                };

                var categoryStackPanel = new StackPanel();
                categoryBorder.Child = categoryStackPanel;

                // 创建顶级名称标签
                var topLabel = new TextBlock
                {
                    Text = category.Name,
                    FontSize = 16,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                categoryStackPanel.Children.Add(topLabel);

                foreach (var group in category.Group)
                {
                    // 创建分组标签
                    if (!string.IsNullOrEmpty(group.Name))
                    {
                        var groupLabel = new TextBlock
                        {
                            Text = group.Name + ":",
                            FontSize = 14,
                            Foreground = new SolidColorBrush(Color.FromRgb(98, 98, 98)),
                            Margin = new Thickness(0, 5, 0, 5)
                        };
                        categoryStackPanel.Children.Add(groupLabel);
                    }

                    // 创建一个 WrapPanel 用于横向显示项目
                    var wrapPanel = new WrapPanel
                    {
                        Margin = new Thickness(0, 5, 0, 5),
                        HorizontalAlignment = HorizontalAlignment.Left
                    };
                    categoryStackPanel.Children.Add(wrapPanel);

                    foreach (var item in group.Items)
                    {
                        if (item.IsDisplay == 0)
                            continue;

                        // 创建一个 StackPanel 用于放置项目名和输入控件
                        var itemPanel = new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            Width = 250,  // 固定宽度，保证对齐
                            Margin = new Thickness(0, 0, 20, 10)  // 设置每个组合之间的间距
                        };

                        // 创建项目标签
                        var itemLabel = new TextBlock
                        {
                            Text = item.Name + ":",
                            FontSize = 14,
                            Width = 100, // 固定项目名宽度
                            VerticalAlignment = VerticalAlignment.Center
                        };

                        // ToolTip
                        var labelToolTip = new ToolTip
                        {
                            Content = "[" + item.Name + "]数值范围: " + item.Range
                        };
                        ToolTipService.SetToolTip(itemLabel, labelToolTip);

                        // 创建输入控件
                        FrameworkElement inputControl;
                        if (item.Input == "输入框")
                        {
                            var textBox = new TextBox
                            {
                                Name = "textBox_" + item.Name,
                                Width = 100,
                                Text = item.Value.ToString(),
                                FontSize = 12
                            };

                            // ToolTip
                            var textBoxToolTip = new ToolTip
                            {
                                Content = "[" + item.Name + "]数值范围: " + item.Range
                            };
                            ToolTipService.SetToolTip(textBox, textBoxToolTip);

                            // TextBox 事件处理
                            textBox.LostFocus += (s, e) =>
                            {
                                if (!string.IsNullOrWhiteSpace(textBox.Text))
                                {
                                    SendDataToSlave(ip, port, slaveId, item.Name, item.Address, item.Length, textBox.Text, item.Range, "失去焦点时提交");
                                }
                            };

                            textBox.KeyDown += (s, e) =>
                            {
                                if (e.Key == Key.Enter)
                                {
                                    if (!string.IsNullOrWhiteSpace(textBox.Text))
                                    {
                                        SendDataToSlave(ip, port, slaveId, item.Name, item.Address, item.Length, textBox.Text, item.Range, "回车时提交");
                                    }
                                    e.Handled = true;
                                }
                            };

                            inputControl = textBox;
                        }
                        else
                        {
                            var comboBox = new ComboBox
                            {
                                Name = "comboBox_" + item.Name,
                                Width = 100
                            };

                            var rangeOptions = item.Range.Split('~')[0].Split('、');
                            foreach (var option in rangeOptions)
                            {
                                comboBox.Items.Add(option);
                            }

                            int valueIndex;
                            if (int.TryParse(item.Value?.ToString(), out valueIndex) && valueIndex >= 0 && valueIndex < comboBox.Items.Count)
                            {
                                comboBox.SelectedIndex = valueIndex;
                            }
                            else
                            {
                                comboBox.SelectedIndex = 0; // 默认选择第一个选项
                            }

                            // ToolTip for ComboBox
                            var comboBoxToolTip = new ToolTip
                            {
                                Content = "[" + item.Name + "]数值范围: " + item.Range
                            };
                            ToolTipService.SetToolTip(comboBox, comboBoxToolTip);

                            comboBox.SelectionChanged += (s, e) =>
                            {
                                SendDataToSlave(ip, port, slaveId, item.Name, item.Address, item.Length, comboBox.SelectedIndex.ToString(), "", "改变下拉时提交");
                            };

                            inputControl = comboBox;
                        }

                        // 将项目标签和输入控件添加到 StackPanel 中
                        itemPanel.Children.Add(itemLabel);
                        itemPanel.Children.Add(inputControl);

                        // 将当前项目的 StackPanel 添加到 WrapPanel
                        wrapPanel.Children.Add(itemPanel);
                    }
                }

                stackPanel.Children.Add(categoryBorder);
            }
        }


        //发送数据到从机的寄存器
        private async void SendDataToSlave(string ip, string port, string slaveId, string name, int address, int length, string input, string range = "", string remark = "")
        {
            if (address < 40000 || address >= 50000)
            {
                MessageBox.Show("IP:" + ip + "[" + address + "]数据不正确,请检查json文件中的数据地址!");
                return;
            }

            if (length != 1 && length != 2)
            {
                MessageBox.Show("IP:" + ip + "数据不正确,请检查json文件中的数据长度!");
                return;
            }

            if (string.IsNullOrEmpty(input))
            {
                MessageBox.Show("IP:" + ip + "[" + name + "]输入不能为空!");
                return;
            }

            if (!Regex.IsMatch(input, @"^-?\d{1,9}(\.\d{0,1})?$|^-?\d{10}$"))
            {
                MessageBox.Show("IP:" + ip + "[" + name + "]输入数值不符合要求!");
                return;
            }

            if (length == 1)
            {
                string pattern = @"^-?\d+$";
                Regex regex = new Regex(pattern);
                if (!regex.IsMatch(input))
                {
                    MessageBox.Show("IP:" + ip + "[" + name + "]输入数值不正确,不能有小数!");
                    return;
                }

                int number = int.Parse(input);
                if (number <= -32768 || number >= 65535)
                {
                    MessageBox.Show("IP:" + ip + "[" + name + "]输入数值不正确,范围应在: -32768到32767 或0到65535 之间!");
                    return;
                }
            }

            if (length == 2)
            {
                string pattern = @"^-?\d{0,6}(\.\d{1,4})?$";
                Regex regex = new Regex(pattern);
                if (!regex.IsMatch(input))
                {
                    MessageBox.Show("IP:" + ip + "[" + name + "]输入数值不正确,整数不超过6位,小数4位!");
                    return;
                }
            }

            if (!string.IsNullOrEmpty(range) && range.Contains("~"))
            {
                var match = Regex.Match(range, @"(-?\d+\.?\d*)~(-?\d+\.?\d*)");
                if (match.Success)
                {
                    double minValue = double.Parse(match.Groups[1].Value);
                    double maxValue = double.Parse(match.Groups[2].Value);

                    WriteLog(minValue + "|" + maxValue);

                    // 检查提取的范围是否有效
                    if (minValue != double.MinValue && maxValue != double.MaxValue)
                    {
                        if (double.TryParse(input, out double inputValue))
                        {
                            if (inputValue < minValue || inputValue > maxValue)
                            {
                                MessageBox.Show("IP:" + ip + "[" + name + "]输入值不正确,超过范围值!");
                                return;
                            }
                        }
                    }
                }
            }

            try
            {
                ushort modbusAddress = Convert.ToUInt16(ConvertAddressToHex(address), 16); // 转换为 ushort 类型

                if (decimal.TryParse(input, out decimal valueToWrite))
                {
                    using (CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(2))) // 2秒超时
                    {
                        TcpClient tcpClient = null; // 在try块外部声明
                        try
                        {
                            await Task.Run(() =>
                            {
                                using (tcpClient = new TcpClient())
                                {
                                    tcpClient.NoDelay = true;
                                    tcpClient.LingerState = new LingerOption(true, 0);
                                    tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                                    tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.KeepAlive, true);

                                    var result = tcpClient.BeginConnect(ip, int.Parse(port), null, null);
                                    bool success = result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(2));

                                    if (!success)
                                    {
                                        // 如果连接超时，捕获异常但不抛出，继续执行后续操作
                                        WriteLog($"IP:{ip}连接超时，未能成功连接。");
                                        return;
                                    }

                                    tcpClient.EndConnect(result);

                                    var factory = new ModbusFactory();
                                    IModbusMaster master = factory.CreateMaster(tcpClient);

                                    byte SlaveId = (byte)int.Parse(slaveId);

                                    if (valueToWrite % 1 == 0 && valueToWrite <= ushort.MaxValue)
                                    {
                                        WriteLog($"IP:{ip},写入从机地址:{modbusAddress},十进制地址:{address},写入数值:{valueToWrite}");
                                        master.WriteSingleRegister(SlaveId, modbusAddress, (ushort)valueToWrite);
                                        this.InputToJson(ip, address, input);
                                    }
                                    else
                                    {
                                        byte[] floatBytes = BitConverter.GetBytes((float)valueToWrite);
                                        ushort[] registers = new ushort[2];
                                        registers[0] = BitConverter.ToUInt16(floatBytes, 2);
                                        registers[1] = BitConverter.ToUInt16(floatBytes, 0);

                                        master.WriteMultipleRegisters(SlaveId, modbusAddress, registers);
                                        WriteLog($"IP:{ip} Register 0: {registers[0]:X4}, Register 1: {registers[1]:X4}");
                                        WriteLog($"IP:{ip},写入从机地址:{modbusAddress},十进制地址:{address},写入数值:{valueToWrite}");
                                        this.InputToJson(ip, address, input);
                                    }
                                }
                            }, cts.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            WriteLog($"IP:{ip}通讯超时！");
                        }
                        catch (SocketException ex)
                        {
                            // 捕获SocketException并记录日志，不抛出异常
                            WriteLog($"IP:{ip}通讯失败: {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            // 捕获其他异常并记录日志
                            WriteLog($"IP:{ip}发生错误: {ex.Message}");
                        }
                        finally
                        {
                            // 确保连接被释放
                            tcpClient?.Close();
                        }

                    }
                }
                else
                {
                    MessageBox.Show($"输入的值[{input}]无效，请输入一个有效的数字！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"IP:{ip}通讯失败: {ex.Message}");
                WriteLog($"IP:{ip}通讯失败: {ex.Message}");
            }
        }

        //当单个信息录入完成后修改json
        public void InputToJson(string ip, int address, string newValue)
        {
            string jsonPath = $"{ip}.json"; // JSON 文件路径

            // 读取JSON文件
            var jsonContent = File.ReadAllText(jsonPath);

            // 解析JSON为JArray
            var modbusArray = JArray.Parse(jsonContent);

            // 遍历JSON对象，找到address并更新value
            foreach (var modbusObject in modbusArray)
            {
                var groups = modbusObject["group"];
                foreach (var group in groups)
                {
                    var items = group["items"];
                    foreach (var item in items)
                    {
                        if (item["address"] != null && (int)item["address"] == address)
                        {
                            // 更新value
                            item["value"] = newValue;
                        }
                    }
                }
            }

            // 将修改后的内容重新写入文件
            File.WriteAllText(jsonPath, modbusArray.ToString());
            WriteLog($"IP:{ip},更新json文件,address:" + address + ",value:" + newValue);
        }

        //清理所有TabPage及其子控件
        private void ClearTabControl()
        {
            // 停止并释放定时器
            if (connectionTimer != null)
            {
                connectionTimer.Stop(); // 停止定时器
                connectionTimer = null; // 清空引用
            }

            // 清空 TabControl 的所有 TabItem
            tabControl.Items.Clear(); // 使用 TabControl.Items 清空所有选项卡
        }

        //将10进制的地址转换为16进制
        private static string ConvertAddressToHex(int address)
        {
            // 减去40000和1
            ushort decimalAddress = (ushort)(address - 40000 - 1);

            // 转换为十六进制字符串并返回
            return decimalAddress.ToString("X"); // 返回大写的十六进制字符串
        }

        //读取从机的数据更新json
        private async Task ReadSlaveToJson(IniData data)
        {
            for (int i = 0; ; i++)
            {
                var hostKey = $"host[{i}]";
                if (!data["Settings"].ContainsKey(hostKey))
                    break;

                var hostInfo = data["Settings"][hostKey].Split(',');
                if (hostInfo.Length == 4)
                {
                    string ip = hostInfo[0];
                    string port = hostInfo[1];
                    string slaveId = hostInfo[2];
                    string name = hostInfo[3];

                    byte SlaveId = (byte)int.Parse(slaveId);
                    var factory = new ModbusFactory();

                    // 读取JSON数据
                    string jsonFileName = $"{ip}.json";
                    if (!File.Exists(jsonFileName))
                    {
                        File.Copy("original.json", jsonFileName, true);
                    }

                    WriteLog($"读取json将从机数据同步回本地: {jsonFileName}");

                    var jsonContent = File.ReadAllText(jsonFileName);
                    var modbusArray = JArray.Parse(jsonContent);

                    TcpClient tcpClient = null; // 在try块外部声明
                    try
                    {
                        using (tcpClient = new TcpClient())
                        {
                            // 设置连接超时为2秒
                            var connectTask = tcpClient.ConnectAsync(ip, int.Parse(port));
                            if (await Task.WhenAny(connectTask, Task.Delay(2000)) == connectTask && tcpClient.Connected)
                            {
                                // 确保连接成功后才创建 modbusMaster
                                tcpClient.NoDelay = true;
                                tcpClient.LingerState = new LingerOption(true, 0);
                                tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                                tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.KeepAlive, true);

                                var modbusMaster = factory.CreateMaster(tcpClient);

                                // 遍历 JSON，逐个处理 items
                                foreach (var modbusObject in modbusArray)
                                {
                                    var groups = modbusObject["group"];
                                    foreach (var group in groups)
                                    {
                                        var items = group["items"];
                                        foreach (var item in items)
                                        {
                                            if (item["address"] != null)
                                            {
                                                ushort startAddress = (ushort)item["address"];
                                                ushort length = ushort.Parse(item["length"].ToString());
                                                if (length > 2) length = 1;
                                                if (length >= 0) length = 1;

                                                try
                                                {
                                                    // 设置Modbus读取操作超时
                                                    var readTask = Task.Run(() =>
                                                        modbusMaster.ReadHoldingRegisters(SlaveId, (ushort)(startAddress - 40001), length)
                                                    );
                                                    if (await Task.WhenAny(readTask, Task.Delay(2000)) == readTask)
                                                    {
                                                        ushort[] registerValues = readTask.Result;
                                                        item["value"] = string.Join("", registerValues);
                                                        WriteLog($"IP:{ip},读取到寄存器地址 {startAddress} 的值: {registerValues[0]}");
                                                    }
                                                    else
                                                    {
                                                        WriteLog($"IP:{ip},读取寄存器超时");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    // 处理读取寄存器时的错误
                                                    WriteLog($"IP:{ip},读取寄存器时发生错误: {ex.Message}");
                                                }
                                            }
                                        }
                                    }
                                }

                                // 将更新后的 JSON 写回文件
                                File.WriteAllText(jsonFileName, modbusArray.ToString());
                                WriteLog($"IP:{ip}读取远程寄存器数据并更新json文件");
                            }
                            else
                            {
                                WriteLog($"IP:{ip},连接超时或连接失败");
                            }
                        }
                    }
                    catch (SocketException)
                    {
                        WriteLog($"IP:{ip},无法连接到从机，将在下次调用时重试...");
                        continue; // 跳过这个从机的处理，继续下一个从机
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"IP:{ip},读取从机数据发生错误: {ex.Message}");
                    }
                    finally
                    {
                        // 确保连接被释放
                        tcpClient?.Close();
                    }
                }
            }
        }

        // 释放资源
        private void CleanUpResources()
        {
            // 写入日志，记录程序关闭
            WriteLog("关闭程序!");

            // 停止并释放所有 Modbus 连接
            if (modbusMasters != null && modbusMasters.Count > 0)
            {
                foreach (var master in modbusMasters.ToList())
                {
                    try
                    {
                        master.Dispose(); // 释放 Modbus 连接
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"释放Modbus连接时出错: {ex.Message}");
                    }
                }
                modbusMasters.Clear(); // 清空集合
            }

            // 停止并释放定时器
            StopAndDisposeTimer(connectionTimer);
            StopAndDisposeTimer(loadTimer);
            StopAndDisposeTimer(readTimer);
        }

        // 释放定时器
        private void StopAndDisposeTimer(object timer)
        {
            if (timer is System.Timers.Timer sysTimer)
            {
                sysTimer.Stop();
                sysTimer.Dispose(); // 释放定时器资源
            }
            else if (timer is System.Windows.Threading.DispatcherTimer dispatcherTimer)
            {
                dispatcherTimer.Stop(); // 只需停止 DispatcherTimer
            }
        }

        // 关闭窗体
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 仅在未确认退出时弹出提示框
            if (!isClosing)
            {
                var result = MessageBox.Show("确认要退出吗？", "退出确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    isClosing = true; // 设置标志变量，表示正在关闭
                    CleanUpResources(); // 调用清理资源的方法
                }
                else
                {
                    e.Cancel = true; // 取消关闭事件
                }
            }
        }

        // 检测键盘输入, 实现按ESC键退出
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                var result = MessageBox.Show("确认要退出吗？", "退出确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    isClosing = true; // 设置标志变量，表示正在关闭
                    CleanUpResources();
                    this.Close(); // 退出程序
                }

                e.Handled = true; // 标记事件已处理
            }
        }

        // 定时器用于检测连接状态
        private async void InitializeConnectionTimer(string ip, string port, string slaveId, Label statusLabel)
        {
            while (true)
            {
                // 使用 Task.Run 将连接检查操作放到后台线程中
                string connectText = await Task.Run(() => CheckSlaveConnection(ip, port, slaveId)); // 调用检测函数

                // 更新 UI，确保在 UI 线程中执行
                if (statusLabel != null) // 确保 statusLabel 不为 null
                {
                    // 更新连接状态
                    if (connectText == "已连接")
                    {
                        statusLabel.Content = "连接状态: 已连接";
                        statusLabel.Foreground = Brushes.Green; // 绿色显示已连接
                    }
                    else
                    {
                        statusLabel.Content = "连接状态: " + connectText;
                        statusLabel.Foreground = Brushes.Red; // 红色显示未连接
                    }
                }

                // 每次检测完成后等待 5 秒
                await Task.Delay(5000);
            }
        }

        //检测和从机的连接:间隔5秒调用
        private string CheckSlaveConnection(string ip, string port, string slaveId)
        {
            string ret = "未连接";

            TcpClient tcpClient = null; // 在try块外部声明
            try
            {
                using (tcpClient = new TcpClient())
                {
                    tcpClient.NoDelay = true; // 禁用Nagle算法，确保及时发送和关闭
                    tcpClient.LingerState = new LingerOption(true, 0); // 强制关闭连接
                    tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true); // 设置端口复用
                    tcpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.KeepAlive, true); // 启用保活

                    var result = tcpClient.BeginConnect(ip, int.Parse(port), null, null);
                    bool success = result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(2)); // 设置2秒的连接超时

                    if (!success)
                    {
                        throw new SocketException((int)SocketError.TimedOut);
                    }

                    // 完成连接
                    tcpClient.EndConnect(result);

                    byte SlaveId = (byte)int.Parse(slaveId);
                    var factory = new ModbusFactory();
                    modbusMaster = factory.CreateMaster(tcpClient);

                    ushort[] testRegister = modbusMaster.ReadHoldingRegisters(SlaveId, 0, 1);

                    ret = "已连接";
                }
            }
            catch (NModbus.SlaveException slaveEx)
            {
                ret = "Modbus异常(非法地址,从机ID重复等)";
                WriteLog($"IP:{ip},从机异常: {slaveEx.Message}");
            }
            catch (SocketException socketEx)
            {
                ret = "网络连接错误或超时";
                //WriteLog($"IP:{ip},网络异常: {socketEx.Message}");
            }
            catch (Exception ex)
            {
                ret = "其他异常";
                WriteLog($"IP:{ip},连接检测失败: {ex.Message}");
            }
            finally
            {
                // 确保连接被释放
                tcpClient?.Close();
            }

            return ret;
        }

        //写日志
        public static void WriteLog(string message)
        {
            try
            {
                // 创建log文件夹路径
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log");

                // 如果log文件夹不存在，创建它
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // 按日期生成日志文件名
                string logFileName = $"{DateTime.Now.ToString("yyyy-MM-dd")}.txt";
                string logFilePath = Path.Combine(logDirectory, logFileName);

                // 创建日志条目
                string logEntry = $"{DateTime.Now}: {message}";

                // 将日志写入对应的文件
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // 这里可以记录异常信息
                // File.AppendAllText("error_log.txt", $"{DateTime.Now}: {ex.Message}" + Environment.NewLine);
            }
        }

        public class ModbusCategory
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int IsDisplay { get; set; }
            public List<ModbusGroup> Group { get; set; }
        }

        public class ModbusGroup
        {
            public string Name { get; set; }
            public List<ModbusItem> Items { get; set; }
        }

        public class ModbusItem
        {
            public string Name { get; set; }
            public string Range { get; set; }
            public string Input { get; set; }
            public int Address { get; set; }
            public int Length { get; set; }
            public dynamic Value { get; set; }
            public int IsDisplay { get; set; }
        }

    }
}