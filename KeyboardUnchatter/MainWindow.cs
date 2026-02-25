using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace KeyboardUnchatter
{
    public partial class MainWindow : Form
    {
        private Color _activeColor;
        private DataGridController _dataGridController;
        private List<long> _intervals = new List<long>();
        private DateTime _lastKeyTime = DateTime.MinValue;
        private int _maxEntries = 50;

        public MainWindow()
        {
            _activeColor = Color.FromArgb(255, 255, 102, 0);

            InitializeComponent();

            // Set window title using assembly title and version
            this.Text = GetAppTitleWithVersion();

            Program.KeyboardMonitor.OnKeyBlocked += OnKeyBlocked;
            Program.KeyboardMonitor.OnKeyPress += OnKeyPress;

            _notifyIcon.Visible = true;

            _thresholdTimeInput.Value = Properties.Settings.Default.chatterThreshold;
            _minimizeCheckBox.Checked = Properties.Settings.Default.openMinimized;
            _activateOnLaunchCheckBox.Checked = Properties.Settings.Default.activateOnLaunch;
            _runAtStartupCheckBox.Checked = Program.GetStartup();
            _startAsAdminCheckBox.Checked = Properties.Settings.Default.startAsAdmin;
            _runAtStartupCheckBox.CheckedChanged += OnRunAtStartupCheckBoxChanged;
            _startAsAdminCheckBox.CheckedChanged += OnStartAsAdminCheckBoxChanged;

            _dataGridController = new DataGridController(_mainDataGrid);

            listBoxIntervals.DrawMode = DrawMode.OwnerDrawFixed;
            listBoxIntervals.DrawItem += ListBoxIntervals_DrawItem;
            Program.InputHook.OnTypingMedianChanged += UpdateTypingSpeedLabel;

            UpdateAdminButtonState();

            if (Properties.Settings.Default.activateOnLaunch)
            {
                ActivateKeyboardMonitor();
            }
        }

        // Helper method to get title and version
        private static string GetAppTitleWithVersion()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var titleAttr = assembly.GetCustomAttribute<System.Reflection.AssemblyTitleAttribute>();
            var version = assembly.GetName().Version;
            string title = titleAttr != null ? titleAttr.Title : assembly.GetName().Name;
            return $"{title} v{version}";
        }

        private void UpdateAdminButtonState()
        {
            if (Program.IsRunningAsAdmin())
            {
                btnRelaunchAsAdmin.Text = "Running as Admin";
                btnRelaunchAsAdmin.Enabled = false;
                return;
            }

            btnRelaunchAsAdmin.Text = "Relaunch as Admin";
            btnRelaunchAsAdmin.Enabled = true;
        }

        private void OnRelaunchAsAdminClick(object sender, EventArgs e)
        {
            if (Program.IsRunningAsAdmin())
            {
                UpdateAdminButtonState();
                return;
            }

            Program.RelaunchAsAdmin();
        }

        private void OnKeyPress(Keys key)
        {
            _dataGridController.AddKeyPress(key);
            //_diagnoseWindow?.RegisterKeyPress();
            RegisterKeyPress();
        }

        private void OnKeyBlocked(Keys key)
        {
            _dataGridController.AddKeyBlock(key);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            Program.KeyboardMonitor.OnKeyBlocked -= OnKeyBlocked;
            Program.KeyboardMonitor.OnKeyPress -= OnKeyPress;
            Program.InputHook.OnTypingMedianChanged -= UpdateTypingSpeedLabel;
            _notifyIcon.Visible = false;
        }

        private void ShowProgramWindow()
        {
            if (!Visible)
            {
                Show();
                WindowState = FormWindowState.Normal;
            }
            Activate();
        }
        public void ResetDiagnostics()
        {
            _intervals.Clear();
            _lastKeyTime = DateTime.MinValue;
            labelTypingSpeed.Text = "Typing speed: N/A";
        }
        public void RegisterKeyPress()
        {
            var now = DateTime.Now;
            if (_lastKeyTime != DateTime.MinValue)
            {
                var interval = (long)(now - _lastKeyTime).TotalMilliseconds;
                _intervals.Insert(0, interval);
                if (_intervals.Count > _maxEntries)
                    _intervals.RemoveAt(_intervals.Count - 1);
                listBoxIntervals.Items.Insert(0, interval);
                if (listBoxIntervals.Items.Count > _maxEntries)
                    listBoxIntervals.Items.RemoveAt(listBoxIntervals.Items.Count - 1);
            }
            _lastKeyTime = now;
        }
        private void UpdateTypingSpeedLabel(double median)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<double>(UpdateTypingSpeedLabel), median);
                return;
            }
            labelTypingSpeed.Text = $"Median: {median:0} ms | {(median > 0 ? 1000.0 / median : 0):0.0} keys/sec";
        }
        #region Events

        private void ListBoxIntervals_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var interval = (long)listBoxIntervals.Items[e.Index];
            bool isChatter = interval < Program.KeyboardMonitor.ChatterTimeMs;
            e.DrawBackground();
            var color = isChatter ? Brushes.Red : Brushes.Black;
            e.Graphics.DrawString(interval + " ms", e.Font, color, e.Bounds);
            e.DrawFocusRectangle();
        }

        private void OnActivateButtonClick(object sender, EventArgs e)
        {
            if (!Program.KeyboardMonitor.Active)
            {
                ActivateKeyboardMonitor();
            }
            else
            {
                DeactivateKeyboardMonitor();
            }
            ActiveControl = null;
        }

        private void ActivateKeyboardMonitor()
        {
            _buttonActivate.BackColor = _activeColor;
            _statusPanel.BackColor = _activeColor;
            _buttonActivate.Text = "Stop";
            _statusPanelLabel.Text = "Active";

            Program.InputHook.TypingSpeedEnabled = true;
            Program.KeyboardMonitor.Activate();
        }

        public void DeactivateKeyboardMonitor()
        {
            _buttonActivate.BackColor = SystemColors.Control;
            _statusPanel.BackColor = SystemColors.ControlDark;
            _buttonActivate.Text = "Activate";
            _statusPanelLabel.Text = "Not Active";

            Program.InputHook.TypingSpeedEnabled = false;
            Program.KeyboardMonitor.Deactivate();
        }

        private void OnMainWindowResize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                Hide();
            }
        }

        private void OnNotifyIconClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ShowProgramWindow();
            }
        }

        private void OnNotifyIconMenuStripItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if(e.ClickedItem == _openMenuItem)
            {
                ShowProgramWindow();
            }

            if(e.ClickedItem == _exitMenuItem)
            {
                Close();
            }
        }

        private void OnThresholdTimeInputValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.chatterThreshold = _thresholdTimeInput.Value;
            Properties.Settings.Default.Save();
            Program.KeyboardMonitor.ChatterTimeMs = System.Convert.ToDouble(_thresholdTimeInput.Value);
        }

        private void OnMinimizeCheckBoxValueChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.openMinimized = _minimizeCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void OnActivateOnLaunchCheckBox(object sender, EventArgs e)
        {
            Properties.Settings.Default.activateOnLaunch = _activateOnLaunchCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void OnRunAtStartupCheckBoxChanged(object sender, EventArgs e)
        {
            Program.SetStartup(_runAtStartupCheckBox.Checked);
        }

        private void OnStartAsAdminCheckBoxChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.startAsAdmin = _startAsAdminCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void DataSortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index == 1 || e.Column.Index == 2)
            {
                e.SortResult = int.Parse(e.CellValue1.ToString()).CompareTo(int.Parse(e.CellValue2.ToString()));
                e.Handled = true;
            }

            if (e.Column.Index == 3)
            {
                var value1 = e.CellValue1.ToString().Replace("%","");
                var value2 = e.CellValue2.ToString().Replace("%", "");

                e.SortResult = int.Parse(value1.ToString()).CompareTo(int.Parse(value2.ToString()));
                e.Handled = true;
            }
        }

        private void OnResetDiagnosticsClick(object sender, EventArgs e)
        {
            Program.InputHook?.ResetDiagnostics();
            Program.KeyboardMonitor?.ResetStatistics();
            ResetDiagnostics();
            _dataGridController.Reset();
            tbTestInput.Clear();
        }

        #endregion
    }
}
