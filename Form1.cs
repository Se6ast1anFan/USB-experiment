using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management; // 必须添加引用: 项目->添加引用->System.Management
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp_USB
{
    public partial class Form1 : Form
    {
        // 界面控件声明
        private DataGridView gridDevices;
        private ComboBox cmbDrives;
        private ListView listFiles;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblSpeed;
        private Label lblUserInfo;
        private TextBox txtLog; // 日志窗口

        public Form1()
        {
            InitializeComponent();
            BuildUI();             // 手动构建界面

            // 初始化数据
            lblUserInfo.Text = $"当前操作用户: {Environment.UserName} | 计算机名: {Environment.MachineName}";
            RefreshUsbList();
            RefreshDriveList();
            Log("系统初始化完成，等待 USB 设备...");
        }

        // --- 1. 界面构建 (已去除特殊符号) ---
        private void BuildUI()
        {
            this.Text = "USB 总线与设备分析实验平台";
            this.Size = new Size(1000, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            // 使用通用字体，防止兼容性问题
            this.Font = new Font("Microsoft Sans Serif", 9);

            // 顶部信息
            lblUserInfo = new Label { Location = new Point(15, 10), AutoSize = true, Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold), ForeColor = Color.DarkBlue };
            this.Controls.Add(lblUserInfo);

            // 左侧：USB设备树
            var grpDev = new GroupBox { Text = "USB 总线设备列表 (WMI)", Location = new Point(15, 40), Size = new Size(550, 250) };
            gridDevices = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, RowHeadersVisible = false, AllowUserToAddRows = false, BackgroundColor = Color.White };
            gridDevices.Columns.Add("Name", "设备名称");
            gridDevices.Columns.Add("ID", "设备ID (VID/PID)");
            gridDevices.Columns.Add("Status", "状态");
            grpDev.Controls.Add(gridDevices);
            this.Controls.Add(grpDev);

            // 右侧：实时日志
            var grpLog = new GroupBox { Text = "系统事件日志", Location = new Point(580, 40), Size = new Size(390, 250) };
            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.Black, ForeColor = Color.Lime };
            grpLog.Controls.Add(txtLog);
            this.Controls.Add(grpLog);

            // 中部：U盘操作区
            var grpOp = new GroupBox { Text = "U盘挂载与文件操作", Location = new Point(15, 300), Size = new Size(955, 380) };

            Label lblSel = new Label { Text = "选择挂载点:", Location = new Point(20, 30), AutoSize = true };
            cmbDrives = new ComboBox { Location = new Point(100, 26), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDrives.SelectedIndexChanged += (s, e) => ListFiles(); // 选U盘自动列出文件

            Button btnRefresh = new Button { Text = "刷新列表", Location = new Point(260, 25), Width = 70 };
            btnRefresh.Click += (s, e) => { RefreshDriveList(); RefreshUsbList(); };

            Button btnWrite = new Button { Text = "写入测试", Location = new Point(340, 25), Width = 80 };
            btnWrite.Click += BtnWrite_Click;

            Button btnDel = new Button { Text = "删除选中", Location = new Point(430, 25), Width = 80 };
            btnDel.Click += BtnDel_Click;

            Button btnCopy = new Button { Text = "拷贝文件(测速)", Location = new Point(520, 25), Width = 120, BackColor = Color.LightSkyBlue };
            btnCopy.Click += BtnCopy_Click;

            grpOp.Controls.AddRange(new Control[] { lblSel, cmbDrives, btnRefresh, btnWrite, btnDel, btnCopy });

            // 文件列表
            listFiles = new ListView { Location = new Point(20, 60), Size = new Size(915, 250), View = View.Details, GridLines = true, FullRowSelect = true };
            listFiles.Columns.Add("文件名", 300);
            listFiles.Columns.Add("大小", 100);
            listFiles.Columns.Add("属性", 100);
            listFiles.Columns.Add("修改时间", 150);
            grpOp.Controls.Add(listFiles);

            // 底部状态条
            progressBar = new ProgressBar { Location = new Point(20, 320), Size = new Size(600, 20), Visible = false };
            lblSpeed = new Label { Location = new Point(630, 320), AutoSize = true, Font = new Font("Consolas", 10, FontStyle.Bold), ForeColor = Color.Red };
            lblStatus = new Label { Location = new Point(20, 350), AutoSize = true, Text = "就绪" };

            grpOp.Controls.Add(progressBar);
            grpOp.Controls.Add(lblSpeed);
            grpOp.Controls.Add(lblStatus);

            this.Controls.Add(grpOp);
        }

        // --- 2. 核心：USB 总线枚举 (增强版：修复转换无效错误) ---
        private void RefreshUsbList()
        {
            gridDevices.Rows.Clear();
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity WHERE DeviceID LIKE 'USB%'");

                // 关键修改：使用 ManagementBaseObject，不要用 var 或 ManagementObject
                foreach (ManagementBaseObject obj in searcher.Get())
                {
                    // 使用 Convert.ToString 处理可能的 null 值，防止报错
                    string name = Convert.ToString(obj["Name"]);
                    string id = Convert.ToString(obj["DeviceID"]);
                    string status = Convert.ToString(obj["Status"]);

                    if (!string.IsNullOrEmpty(name))
                    {
                        gridDevices.Rows.Add(name, id, status);
                    }
                }
            }
            catch (Exception ex)
            {
                // 即使出错也只记录日志，不弹窗打扰用户
                Log("枚举 USB 列表警告: " + ex.Message);
            }
        }

        // --- 3. 核心：U盘盘符检测 ---
        private void RefreshDriveList()
        {
            string oldSel = cmbDrives.SelectedItem as string;
            cmbDrives.Items.Clear();
            var drives = DriveInfo.GetDrives().Where(d => d.DriveType == DriveType.Removable && d.IsReady);

            foreach (var d in drives)
            {
                cmbDrives.Items.Add(d.Name);
            }

            if (cmbDrives.Items.Count > 0)
            {
                if (oldSel != null && cmbDrives.Items.Contains(oldSel)) cmbDrives.SelectedItem = oldSel;
                else cmbDrives.SelectedIndex = 0;
            }
            else
            {
                listFiles.Items.Clear();
            }
        }

        // --- 4. 核心：热拔插监听 (插入和拔出都增加延迟，彻底解决报错) ---
        protected override void WndProc(ref Message m)
        {
            const int WM_DEVICECHANGE = 0x0219;
            const int DBT_DEVICEARRIVAL = 0x8000;
            const int DBT_DEVICEREMOVECOMPLETE = 0x8004;

            if (m.Msg == WM_DEVICECHANGE)
            {
                int eventType = m.WParam.ToInt32();

                // 无论是插入还是拔出，都统一处理
                if (eventType == DBT_DEVICEARRIVAL || eventType == DBT_DEVICEREMOVECOMPLETE)
                {
                    string action = (eventType == DBT_DEVICEARRIVAL) ? "插入" : "拔出";
                    Log($"[硬件事件] 检测到设备{action}！正在同步系统状态...");

                    // 统一延迟 1 秒，等 Windows 忙完
                    Task.Delay(1000).ContinueWith(t =>
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            RefreshDriveList();
                            RefreshUsbList();
                            Log($"[系统] 设备列表刷新完毕 ({action})");
                        });
                    });
                }
            }
            base.WndProc(ref m);
        }

        // --- 5. 文件操作逻辑 ---

        // 列出文件 (含隐藏)
        private void ListFiles()
        {
            listFiles.Items.Clear();
            if (cmbDrives.SelectedItem == null) return;
            string path = cmbDrives.SelectedItem.ToString();

            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                foreach (FileInfo f in dir.GetFiles())
                {
                    var item = new ListViewItem(f.Name);
                    item.SubItems.Add((f.Length / 1024.0).ToString("F1") + " KB");
                    item.SubItems.Add(f.Attributes.ToString());
                    item.SubItems.Add(f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"));

                    // 隐藏文件变灰
                    if ((f.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                        item.ForeColor = Color.Gray;

                    listFiles.Items.Add(item);
                }
            }
            catch { Log("无法读取 U 盘文件列表"); }
        }

        // 写入测试
        private void BtnWrite_Click(object sender, EventArgs e)
        {
            if (cmbDrives.SelectedItem == null) return;
            string path = Path.Combine(cmbDrives.SelectedItem.ToString(), "USB_TEST_LOG.txt");
            File.WriteAllText(path, $"写入时间: {DateTime.Now}\r\n操作员: {Environment.UserName}");
            Log("写入文本文件成功: " + path);
            ListFiles();
        }

        // 删除文件
        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (listFiles.SelectedItems.Count == 0 || cmbDrives.SelectedItem == null) return;
            string fileName = listFiles.SelectedItems[0].Text;
            string path = Path.Combine(cmbDrives.SelectedItem.ToString(), fileName);

            if (MessageBox.Show($"确定删除 {fileName}?", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                File.Delete(path);
                Log("已删除文件: " + fileName);
                ListFiles();
            }
        }

        // 拷贝与测速 (异步)
        private async void BtnCopy_Click(object sender, EventArgs e)
        {
            if (cmbDrives.SelectedItem == null) return;
            string targetDrive = cmbDrives.SelectedItem.ToString();

            OpenFileDialog dlg = new OpenFileDialog { Title = "选择要拷贝到U盘的大文件" };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string src = dlg.FileName;
            string dest = Path.Combine(targetDrive, Path.GetFileName(src));

            progressBar.Visible = true;
            progressBar.Value = 0;
            lblStatus.Text = "正在拷贝...";
            Log($"开始拷贝: {Path.GetFileName(src)} -> {targetDrive}");

            // 异步任务，防止界面卡死
            await Task.Run(() =>
            {
                byte[] buffer = new byte[1024 * 1024]; // 1MB 缓存
                using (FileStream fsIn = new FileStream(src, FileMode.Open, FileAccess.Read))
                using (FileStream fsOut = new FileStream(dest, FileMode.Create, FileAccess.Write))
                {
                    long totalLen = fsIn.Length;
                    long totalRead = 0;
                    int readSize;

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    long lastRead = 0;
                    long lastTime = 0;

                    while ((readSize = fsIn.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        fsOut.Write(buffer, 0, readSize);
                        totalRead += readSize;

                        long currTime = sw.ElapsedMilliseconds;
                        if (currTime - lastTime > 500) // 每0.5秒刷新一次速度
                        {
                            double speed = ((totalRead - lastRead) / 1024.0 / 1024.0) / ((currTime - lastTime) / 1000.0);

                            // 切换回主线程更新UI
                            this.Invoke((MethodInvoker)delegate {
                                progressBar.Value = (int)((double)totalRead / totalLen * 100);
                                lblSpeed.Text = $"{speed:F1} MB/s";
                            });

                            lastRead = totalRead;
                            lastTime = currTime;
                        }
                    }
                    sw.Stop();
                }
            });

            Log("拷贝完成！");
            lblStatus.Text = "就绪";
            progressBar.Visible = false;
            lblSpeed.Text = "";
            ListFiles();
        }

        private void Log(string msg)
        {
            // 使用标准的日期时间格式
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n");
        }
    }
}