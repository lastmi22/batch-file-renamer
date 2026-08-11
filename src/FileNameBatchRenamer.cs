using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("File Name Batch Renamer")]
[assembly: System.Reflection.AssemblyDescription("Batch file renaming utility by AI DO")]
[assembly: System.Reflection.AssemblyCompany("AI DO")]
[assembly: System.Reflection.AssemblyProduct("File Name Batch Renamer")]
[assembly: System.Reflection.AssemblyVersion("0.1.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("0.1.0.0")]

namespace FileNameBatchRenamer
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    internal sealed class RowDraft
    {
        public bool Selected { get; set; }
        public string NewName { get; set; }
    }

    internal sealed class RenameEntry
    {
        public string Source { get; set; }
        public string Target { get; set; }
        public string Temp { get; set; }
    }

    internal sealed class MainForm : Form
    {
        private const string SelectColumn = "Select";
        private const string CurrentNameColumn = "CurrentName";
        private const string ExtensionColumn = "Extension";
        private const string ModifiedColumn = "Modified";
        private const string NewNameColumn = "NewName";

        private readonly TextBox folderTextBox;
        private readonly TextBox suffixTextBox;
        private readonly CheckBox modifyExtensionCheckBox;
        private readonly ComboBox sortComboBox;
        private readonly ComboBox ruleComboBox;
        private readonly TextBox ruleStartTextBox;
        private readonly TextBox ruleEndTextBox;
        private readonly DataGridView filesGrid;
        private readonly CheckBox selectAllCheckBox;
        private readonly ToolStripStatusLabel statusLabel;
        private bool isShowingExtensions;
        private bool isApplyingBulkSelection;
        private bool isSynchronizingSelectAllCheckBox;
        private readonly Dictionary<string, RowDraft> drafts =
            new Dictionary<string, RowDraft>(StringComparer.OrdinalIgnoreCase);

        public MainForm()
        {
            Text = "文件名批量修改工具";
            try
            {
                Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch (Exception)
            {
                // A missing shell icon association must not prevent the tool from opening.
            }
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(980, 620);
            Size = new Size(1280, 800);
            Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point);
            BackColor = Color.FromArgb(244, 246, 248);

            var titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(16, 24, 39)
            };
            var titleLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold, GraphicsUnit.Point),
                Location = new Point(68, 17),
                Text = "文件名批量修改"
            };
            var brandIconBox = new PictureBox
            {
                BackColor = Color.Transparent,
                Image = Icon == null ? null : Icon.ToBitmap(),
                Location = new Point(20, 14),
                Size = new Size(36, 36),
                SizeMode = PictureBoxSizeMode.Zoom,
                TabStop = false
            };
            titleLabel.BackColor = Color.Transparent;
            titleLabel.Text = "\u6587\u4ef6\u540d\u6279\u91cf\u4fee\u6539";
            titlePanel.Controls.Add(brandIconBox);
            titlePanel.Controls.Add(titleLabel);

            var controlsPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 126,
                Padding = new Padding(16, 10, 16, 8),
                BackColor = Color.White
            };

            var folderLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 34,
                ColumnCount = 4,
                RowCount = 1
            };
            folderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            folderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            folderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 104));
            folderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));

            folderTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Text = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar)
            };
            var folderLabel = CreateLabel("文件夹", ContentAlignment.MiddleLeft);
            var chooseFolderButton = CreateButton("选择文件夹", Color.FromArgb(229, 234, 240));
            var refreshButton = CreateButton("刷新列表", Color.FromArgb(16, 24, 39), Color.White);
            chooseFolderButton.Click += delegate { ChooseFolder(); };
            refreshButton.Click += delegate { RefreshFiles(); };
            folderLayout.Controls.Add(folderLabel, 0, 0);
            folderLayout.Controls.Add(folderTextBox, 1, 0);
            folderLayout.Controls.Add(chooseFolderButton, 2, 0);
            folderLayout.Controls.Add(refreshButton, 3, 0);

            var filterLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 34,
                ColumnCount = 7,
                RowCount = 1
            };
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 98));
            filterLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            suffixTextBox = new TextBox { Dock = DockStyle.Fill };
            modifyExtensionCheckBox = new CheckBox
            {
                Text = "修改后缀",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Checked = false
            };
            modifyExtensionCheckBox.CheckedChanged += ModifyExtensionCheckBoxCheckedChanged;
            sortComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            sortComboBox.Items.AddRange(new object[]
            {
                "名称（A-Z）",
                "名称（Z-A）",
                "修改时间（新到旧）",
                "修改时间（旧到新）",
                "后缀（A-Z）"
            });
            sortComboBox.SelectedIndex = 0;
            sortComboBox.SelectedIndexChanged += delegate { RefreshFiles(false); };
            var applyFilterButton = CreateButton("应用筛选", Color.FromArgb(229, 234, 240));
            applyFilterButton.Click += delegate { RefreshFiles(); };
            var filterHint = CreateLabel("支持多个后缀，以逗号或分号分隔", ContentAlignment.MiddleLeft);
            filterHint.ForeColor = Color.FromArgb(96, 110, 128);
            filterHint.Font = new Font("Microsoft YaHei UI", 8.5f, FontStyle.Regular, GraphicsUnit.Point);

            var ruleLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 34,
                ColumnCount = 7,
                RowCount = 1
            };
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 54));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 208));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 68));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 68));
            ruleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));

            ruleComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            ruleComboBox.Items.AddRange(new object[]
            {
                "选择规则",
                "删除日期前缀",
                "删除日期后缀",
                "数字改中文",
                "只保留区间内文字"
            });
            ruleComboBox.SelectedIndex = 0;
            ruleComboBox.SelectedIndexChanged += delegate { UpdateRuleInputState(); };

            ruleStartTextBox = new TextBox { Dock = DockStyle.Fill, Enabled = false };
            ruleEndTextBox = new TextBox { Dock = DockStyle.Fill, Enabled = false };
            var applyRuleButton = CreateButton("处理", Color.FromArgb(229, 234, 240));
            applyRuleButton.Click += delegate { ApplySelectedRule(); };

            ruleLayout.Controls.Add(CreateLabel("规则", ContentAlignment.MiddleLeft), 0, 0);
            ruleLayout.Controls.Add(ruleComboBox, 1, 0);
            ruleLayout.Controls.Add(CreateLabel("起始", ContentAlignment.MiddleLeft), 2, 0);
            ruleLayout.Controls.Add(ruleStartTextBox, 3, 0);
            ruleLayout.Controls.Add(CreateLabel("结束", ContentAlignment.MiddleLeft), 4, 0);
            ruleLayout.Controls.Add(ruleEndTextBox, 5, 0);
            ruleLayout.Controls.Add(applyRuleButton, 6, 0);

            filterLayout.Controls.Add(CreateLabel("后缀筛选", ContentAlignment.MiddleLeft), 0, 0);
            filterLayout.Controls.Add(suffixTextBox, 1, 0);
            filterLayout.Controls.Add(modifyExtensionCheckBox, 2, 0);
            filterLayout.Controls.Add(CreateLabel("排序", ContentAlignment.MiddleLeft), 3, 0);
            filterLayout.Controls.Add(sortComboBox, 4, 0);
            filterLayout.Controls.Add(applyFilterButton, 5, 0);
            filterLayout.Controls.Add(filterHint, 6, 0);
            controlsPanel.Controls.Add(folderLayout);
            controlsPanel.Controls.Add(ruleLayout);
            controlsPanel.Controls.Add(filterLayout);

            filesGrid = CreateGrid();
            filesGrid.ColumnHeaderMouseClick += FilesGridColumnHeaderMouseClick;
            filesGrid.KeyDown += FilesGridKeyDown;
            filesGrid.CurrentCellDirtyStateChanged += FilesGridCurrentCellDirtyStateChanged;
            filesGrid.CellValueChanged += FilesGridCellValueChanged;
            filesGrid.CellContentClick += FilesGridCellContentClick;

            // A real control avoids DataGridView's header-click/edit-state conflict.
            selectAllCheckBox = new CheckBox
            {
                AutoSize = false,
                Appearance = Appearance.Normal,
                BackColor = Color.FromArgb(233, 238, 244),
                FlatStyle = FlatStyle.Standard,
                Text = "\u5168\u9009",
                TextAlign = ContentAlignment.MiddleLeft,
                CheckAlign = ContentAlignment.MiddleLeft,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            selectAllCheckBox.CheckedChanged += SelectAllCheckBoxCheckedChanged;
            filesGrid.Controls.Add(selectAllCheckBox);
            filesGrid.SizeChanged += delegate { PositionSelectAllCheckBox(); };
            filesGrid.Scroll += delegate { PositionSelectAllCheckBox(); };
            filesGrid.ColumnWidthChanged += delegate { PositionSelectAllCheckBox(); };
            filesGrid.ColumnDisplayIndexChanged += delegate { PositionSelectAllCheckBox(); };
            filesGrid.ColumnHeadersHeightChanged += delegate { PositionSelectAllCheckBox(); };
            filesGrid.Layout += delegate { PositionSelectAllCheckBox(); };

            var actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                Padding = new Padding(16, 10, 16, 10),
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.White
            };
            var renameButton = CreateButton("批量更改", Color.FromArgb(22, 114, 71), Color.White, 112);
            var clearButton = CreateButton("清空修改框", Color.FromArgb(229, 234, 240), Color.Black, 108);
            renameButton.Click += delegate { BatchRename(); };
            clearButton.Click += delegate { ClearNewNames(); };
            actionPanel.Controls.Add(renameButton);
            actionPanel.Controls.Add(clearButton);

            var statusStrip = new StatusStrip
            {
                Dock = DockStyle.Bottom,
                SizingGrip = false,
                BackColor = Color.FromArgb(235, 239, 244)
            };
            statusLabel = new ToolStripStatusLabel { Text = "准备就绪" };
            statusStrip.Items.Add(statusLabel);

            Controls.Add(filesGrid);
            Controls.Add(actionPanel);
            Controls.Add(statusStrip);
            Controls.Add(controlsPanel);
            Controls.Add(titlePanel);

            Shown += delegate { RefreshFiles(); };
        }

        private static Label CreateLabel(string text, ContentAlignment alignment)
        {
            return new Label
            {
                Text = text,
                TextAlign = alignment,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(39, 48, 61)
            };
        }

        private static Button CreateButton(string text, Color backColor, Color? foreColor = null, int width = 0)
        {
            var button = new Button
            {
                Text = text,
                Dock = width == 0 ? DockStyle.Fill : DockStyle.None,
                Height = 34,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = foreColor ?? Color.FromArgb(27, 35, 48),
                Cursor = Cursors.Hand,
                Margin = new Padding(5, 0, 0, 0)
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(190, 199, 210);
            button.FlatAppearance.BorderSize = backColor == Color.FromArgb(16, 24, 39) || backColor == Color.FromArgb(22, 114, 71) ? 0 : 1;
            if (width > 0)
            {
                button.Width = width;
            }
            return button;
        }

        private static DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = true,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(226, 231, 237),
                RowTemplate = { Height = 34 }
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(233, 238, 244);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 41, 55);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold, GraphicsUnit.Point);
            grid.ColumnHeadersHeight = 38;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(222, 235, 250);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(20, 28, 39);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 251, 253);

            grid.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = SelectColumn,
                HeaderText = "☐ 全选",
                Width = 92,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = CurrentNameColumn,
                HeaderText = "当前文件名",
                ReadOnly = true,
                MinimumWidth = 240,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 42,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = ExtensionColumn,
                HeaderText = "后缀",
                ReadOnly = true,
                Width = 100,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = ModifiedColumn,
                HeaderText = "修改时间",
                ReadOnly = true,
                Width = 165,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = NewNameColumn,
                HeaderText = "新文件名（可编辑）",
                MinimumWidth = 300,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 58,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            return grid;
        }

        private void ChooseFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择要读取和批量重命名的文件夹";
                dialog.SelectedPath = Directory.Exists(folderTextBox.Text) ? folderTextBox.Text : AppDomain.CurrentDomain.BaseDirectory;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    folderTextBox.Text = dialog.SelectedPath;
                    RefreshFiles();
                }
            }
        }

        private void RefreshFiles(bool showErrors = true, bool captureDrafts = true)
        {
            if (captureDrafts)
            {
                CaptureVisibleDrafts();
            }
            UpdateNameColumnHeaders();
            string folderPath;
            try
            {
                folderPath = Path.GetFullPath(folderTextBox.Text.Trim());
            }
            catch (Exception)
            {
                if (showErrors)
                {
                    MessageBox.Show(this, "文件夹路径无效。", "无法读取", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }
            if (!Directory.Exists(folderPath))
            {
                if (showErrors)
                {
                    MessageBox.Show(this, "找不到指定的文件夹。", "无法读取", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return;
            }

            try
            {
                var extensions = ParseExtensions(suffixTextBox.Text);
                var selfPath = Path.GetFullPath(Application.ExecutablePath);
                var files = new DirectoryInfo(folderPath).GetFiles()
                    .Where(delegate(FileInfo file)
                    {
                        return !string.Equals(Path.GetFullPath(file.FullName), selfPath, StringComparison.OrdinalIgnoreCase)
                            && (extensions.Count == 0 || extensions.Contains(file.Extension.ToLowerInvariant()));
                    })
                    .ToList();

                SortFiles(files);
                filesGrid.Rows.Clear();
                foreach (var file in files)
                {
                    RowDraft draft;
                    drafts.TryGetValue(file.FullName, out draft);
                    int rowIndex = filesGrid.Rows.Add(
                        draft != null && draft.Selected,
                        GetDisplayedCurrentName(file),
                        string.IsNullOrEmpty(file.Extension) ? "（无后缀）" : file.Extension,
                        file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                        draft == null ? string.Empty : GetDisplayedDraftName(draft.NewName, file.Extension));
                    filesGrid.Rows[rowIndex].Tag = file.FullName;
                }
                filesGrid.ClearSelection();
                UpdateSelectAllHeader();
                statusLabel.Text = string.Format("已显示 {0} 个文件。批量更改只会处理当前列表中已勾选且新文件名不为空的项目。", files.Count);
            }
            catch (UnauthorizedAccessException)
            {
                if (showErrors)
                {
                    MessageBox.Show(this, "没有读取该文件夹的权限。", "无法读取", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                if (showErrors)
                {
                    MessageBox.Show(this, ex.Message, "无法读取", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ModifyExtensionCheckBoxCheckedChanged(object sender, EventArgs e)
        {
            CaptureVisibleDrafts();
            isShowingExtensions = modifyExtensionCheckBox.Checked;
            RefreshFiles(false, false);
        }

        private void UpdateRuleInputState()
        {
            bool rangeRuleSelected = ruleComboBox.SelectedIndex == 4;
            ruleStartTextBox.Enabled = rangeRuleSelected;
            ruleEndTextBox.Enabled = rangeRuleSelected;
            if (!rangeRuleSelected)
            {
                ruleStartTextBox.BackColor = SystemColors.Control;
                ruleEndTextBox.BackColor = SystemColors.Control;
            }
            else
            {
                ruleStartTextBox.BackColor = Color.White;
                ruleEndTextBox.BackColor = Color.White;
            }
        }

        private void ApplySelectedRule()
        {
            int ruleIndex = ruleComboBox.SelectedIndex;
            if (ruleIndex <= 0)
            {
                MessageBox.Show(this, "请先从规则下拉菜单中选择一项处理规则。", "请选择规则", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int start = 0;
            int end = 0;
            if (ruleIndex == 4)
            {
                if (!int.TryParse(ruleStartTextBox.Text.Trim(), out start)
                    || !int.TryParse(ruleEndTextBox.Text.Trim(), out end)
                    || start < 1
                    || end < start)
                {
                    MessageBox.Show(this, "请输入有效的文字区间，例如起始 3、结束 12。", "区间无效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            filesGrid.EndEdit();
            isApplyingBulkSelection = true;
            int processedCount = 0;
            try
            {
                foreach (DataGridViewRow row in filesGrid.Rows)
                {
                    string input = GetRuleInputName(row);
                    string result = ApplyNameRule(input, ruleIndex, start, end);
                    row.Cells[NewNameColumn].Value = result;
                    row.Cells[SelectColumn].Value = !string.IsNullOrWhiteSpace(result);
                    processedCount++;
                }
            }
            finally
            {
                isApplyingBulkSelection = false;
            }

            CaptureVisibleDrafts();
            UpdateSelectAllHeader();
            string ruleName = Convert.ToString(ruleComboBox.SelectedItem);
            statusLabel.Text = string.Format("已对当前筛选结果的 {0} 行应用规则“{1}”，可继续叠加下一条规则。", processedCount, ruleName);
        }

        private string GetRuleInputName(DataGridViewRow row)
        {
            string value = Convert.ToString(row.Cells[NewNameColumn].Value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                value = Convert.ToString(row.Cells[CurrentNameColumn].Value ?? string.Empty).Trim();
            }
            if (!modifyExtensionCheckBox.Checked)
            {
                var path = row.Tag as string;
                string originalExtension = string.IsNullOrEmpty(path) ? string.Empty : Path.GetExtension(path);
                if (!string.IsNullOrEmpty(originalExtension)
                    && value.EndsWith(originalExtension, StringComparison.OrdinalIgnoreCase))
                {
                    value = value.Substring(0, value.Length - originalExtension.Length);
                }
            }
            return value;
        }

        private static string ApplyNameRule(string value, int ruleIndex, int start, int end)
        {
            switch (ruleIndex)
            {
                case 1:
                    return RemoveDatePrefix(value);
                case 2:
                    return RemoveDateSuffix(value);
                case 3:
                    return ConvertDigitsToChinese(value);
                case 4:
                    return KeepTextRange(value, start, end);
                default:
                    return value;
            }
        }

        private static string RemoveDatePrefix(string value)
        {
            return new Regex(
                @"^\s*[\[\(【]?(?:(?:19|20)\d{2}[-_.]\d{1,2}[-_.]\d{1,2}|\d{8}|(?:19|20)\d{2}年\d{1,2}月\d{1,2}日?)[\]\)】]?[\s_.-]*",
                RegexOptions.CultureInvariant).Replace(value, string.Empty, 1);
        }

        private static string RemoveDateSuffix(string value)
        {
            return new Regex(
                @"[\s_.-]*[\[\(【]?(?:(?:19|20)\d{2}[-_.]\d{1,2}[-_.]\d{1,2}|\d{8}|(?:19|20)\d{2}年\d{1,2}月\d{1,2}日?)[\]\)】]?\s*$",
                RegexOptions.CultureInvariant).Replace(value, string.Empty, 1);
        }

        private static string ConvertDigitsToChinese(string value)
        {
            const string chineseDigits = "〇一二三四五六七八九";
            var builder = new StringBuilder(value.Length);
            foreach (char character in value)
            {
                if (character >= '0' && character <= '9')
                {
                    builder.Append(chineseDigits[character - '0']);
                }
                else
                {
                    builder.Append(character);
                }
            }
            return builder.ToString();
        }

        private static string KeepTextRange(string value, int start, int end)
        {
            if (string.IsNullOrEmpty(value) || start > value.Length)
            {
                return string.Empty;
            }
            int startIndex = start - 1;
            int length = Math.Min(end, value.Length) - startIndex;
            return length <= 0 ? string.Empty : value.Substring(startIndex, length);
        }

        private void UpdateNameColumnHeaders()
        {
            if (filesGrid == null || filesGrid.Columns.Count == 0)
            {
                return;
            }
            filesGrid.Columns[CurrentNameColumn].HeaderText = modifyExtensionCheckBox.Checked ? "当前文件名（含后缀）" : "当前文件名";
            filesGrid.Columns[NewNameColumn].HeaderText = "新文件名";
        }

        private string GetDisplayedCurrentName(FileInfo file)
        {
            return modifyExtensionCheckBox.Checked ? file.Name : Path.GetFileNameWithoutExtension(file.Name);
        }

        private string GetDisplayedDraftName(string canonicalName, string originalExtension)
        {
            if (string.IsNullOrEmpty(canonicalName) || modifyExtensionCheckBox.Checked || string.IsNullOrEmpty(originalExtension))
            {
                return canonicalName;
            }
            if (canonicalName.EndsWith(originalExtension, StringComparison.OrdinalIgnoreCase))
            {
                return canonicalName.Substring(0, canonicalName.Length - originalExtension.Length);
            }
            return Path.GetFileNameWithoutExtension(canonicalName);
        }

        private static HashSet<string> ParseExtensions(string rawInput)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(rawInput))
            {
                return result;
            }
            var parts = rawInput.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var extension = part.Trim();
                if (extension == "*" || extension == "*.*" || extension == "全部")
                {
                    result.Clear();
                    return result;
                }
                extension = extension.TrimStart('*');
                if (!extension.StartsWith(".", StringComparison.Ordinal))
                {
                    extension = "." + extension;
                }
                if (extension.Length > 1)
                {
                    result.Add(extension.ToLowerInvariant());
                }
            }
            return result;
        }

        private void SortFiles(List<FileInfo> files)
        {
            switch (sortComboBox.SelectedIndex)
            {
                case 0:
                    files.Sort(delegate(FileInfo left, FileInfo right) { return StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name); });
                    break;
                case 1:
                    files.Sort(delegate(FileInfo left, FileInfo right) { return StringComparer.CurrentCultureIgnoreCase.Compare(right.Name, left.Name); });
                    break;
                case 2:
                    files.Sort(delegate(FileInfo left, FileInfo right) { return right.LastWriteTime.CompareTo(left.LastWriteTime); });
                    break;
                case 3:
                    files.Sort(delegate(FileInfo left, FileInfo right) { return left.LastWriteTime.CompareTo(right.LastWriteTime); });
                    break;
                case 4:
                    files.Sort(delegate(FileInfo left, FileInfo right)
                    {
                        int extensionComparison = StringComparer.CurrentCultureIgnoreCase.Compare(left.Extension, right.Extension);
                        return extensionComparison != 0 ? extensionComparison : StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name);
                    });
                    break;
                default:
                    files.Sort(delegate(FileInfo left, FileInfo right) { return StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name); });
                    break;
            }
        }

        private void CaptureVisibleDrafts()
        {
            foreach (DataGridViewRow row in filesGrid.Rows)
            {
                var path = row.Tag as string;
                if (string.IsNullOrEmpty(path))
                {
                    continue;
                }
                drafts[path] = new RowDraft
                {
                    Selected = Convert.ToBoolean(row.Cells[SelectColumn].Value ?? false),
                    NewName = ToCanonicalDraftName(
                        Convert.ToString(row.Cells[NewNameColumn].Value ?? string.Empty),
                        Path.GetExtension(path),
                        isShowingExtensions)
                };
            }
        }

        private static string ToCanonicalDraftName(string displayedName, string originalExtension, bool displayIncludesExtension)
        {
            var name = (displayedName ?? string.Empty).Trim();
            if (name.Length == 0)
            {
                return string.Empty;
            }
            if (!displayIncludesExtension)
            {
                if (!string.IsNullOrEmpty(originalExtension) && name.EndsWith(originalExtension, StringComparison.OrdinalIgnoreCase))
                {
                    name = name.Substring(0, name.Length - originalExtension.Length);
                }
                return name + originalExtension;
            }
            return string.IsNullOrEmpty(Path.GetExtension(name)) ? name + originalExtension : name;
        }

        private void CopyCurrentNames()
        {
            CopyNames(filesGrid.Rows.Cast<DataGridViewRow>());
        }

        private void PasteNames()
        {
            PasteNamesIntoRows(filesGrid.Rows.Cast<DataGridViewRow>());
        }

        private void PasteNamesIntoRows(IEnumerable<DataGridViewRow> rowsToFill)
        {
            string clipboardText;
            if (!TryGetClipboardText(out clipboardText))
            {
                return;
            }
            filesGrid.EndEdit();
            var rows = rowsToFill.OrderBy(delegate(DataGridViewRow row) { return row.Index; }).ToList();
            var lines = clipboardText.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            int count = Math.Min(lines.Length, rows.Count);
            for (int index = 0; index < count; index++)
            {
                var value = lines[index].Trim();
                var row = rows[index];
                row.Cells[NewNameColumn].Value = value;
                row.Cells[SelectColumn].Value = !string.IsNullOrWhiteSpace(value);
            }
            CaptureVisibleDrafts();
            UpdateSelectAllHeader();
            statusLabel.Text = string.Format("已按所选行顺序填入 {0} 行，非空行已自动勾选。", count);
        }

        private void CopyNames(IEnumerable<DataGridViewRow> rowsToCopy)
        {
            var names = rowsToCopy
                .OrderBy(delegate(DataGridViewRow row) { return row.Index; })
                .Select(delegate(DataGridViewRow row) { return Convert.ToString(row.Cells[CurrentNameColumn].Value); })
                .ToArray();
            if (names.Length == 0)
            {
                return;
            }
            if (!TrySetClipboardText(string.Join(Environment.NewLine, names)))
            {
                return;
            }
            statusLabel.Text = string.Format("已复制 {0} 个当前文件名，可编辑后粘贴到新文件名列。", names.Length);
        }

        private void SelectAllCheckBoxCheckedChanged(object sender, EventArgs e)
        {
            if (!isSynchronizingSelectAllCheckBox)
            {
                SetAllSelections(selectAllCheckBox.Checked);
            }
        }

        private void PositionSelectAllCheckBox()
        {
            if (selectAllCheckBox == null || filesGrid.Columns.Count == 0)
            {
                return;
            }

            var bounds = filesGrid.GetCellDisplayRectangle(filesGrid.Columns[SelectColumn].Index, -1, true);
            bool visible = bounds.Width > 0 && bounds.Height > 0;
            selectAllCheckBox.Visible = visible;
            if (!visible)
            {
                return;
            }

            selectAllCheckBox.Bounds = new Rectangle(
                bounds.X + 4,
                bounds.Y + 3,
                Math.Max(1, bounds.Width - 8),
                Math.Max(1, bounds.Height - 6));
            selectAllCheckBox.BringToFront();
        }

        private void SynchronizeSelectAllCheckBox()
        {
            if (selectAllCheckBox == null)
            {
                return;
            }

            isSynchronizingSelectAllCheckBox = true;
            try
            {
                selectAllCheckBox.Checked = AreAllVisibleRowsSelected();
            }
            finally
            {
                isSynchronizingSelectAllCheckBox = false;
            }
            PositionSelectAllCheckBox();
        }

        private void FilesGridColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0)
            {
                return;
            }
            int selectIndex = filesGrid.Columns[SelectColumn].Index;
            int currentNameIndex = filesGrid.Columns[CurrentNameColumn].Index;
            int newNameIndex = filesGrid.Columns[NewNameColumn].Index;
            if (e.ColumnIndex == selectIndex)
            {
                return;
            }
            if (e.ColumnIndex != currentNameIndex && e.ColumnIndex != newNameIndex)
            {
                return;
            }
            filesGrid.ClearSelection();
            if (filesGrid.Rows.Count > 0)
            {
                filesGrid.CurrentCell = filesGrid.Rows[0].Cells[e.ColumnIndex];
            }
            foreach (DataGridViewRow row in filesGrid.Rows)
            {
                row.Cells[e.ColumnIndex].Selected = true;
            }
            statusLabel.Text = e.ColumnIndex == currentNameIndex
                ? "已选择当前文件名整列，按 Ctrl+C 可复制。"
                : "已选择新文件名整列，按 Ctrl+V 可按行粘贴。";
        }

        private void FilesGridKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C && filesGrid.SelectedCells.Count > 0)
            {
                CopySelectedCellsAsExcel();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.Control && e.KeyCode == Keys.V && filesGrid.SelectedCells.Count > 0)
            {
                PasteExcelValuesIntoNewNames();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void CopySelectedCellsAsExcel()
        {
            var selectedCells = filesGrid.SelectedCells.Cast<DataGridViewCell>().ToList();
            if (selectedCells.All(delegate(DataGridViewCell cell) { return cell.ColumnIndex == selectedCells[0].ColumnIndex; }))
            {
                var copiedColumn = selectedCells
                    .OrderBy(delegate(DataGridViewCell cell) { return cell.RowIndex; })
                    .Select(delegate(DataGridViewCell cell) { return Convert.ToString(cell.Value ?? string.Empty); })
                    .ToArray();
                if (!TrySetClipboardText(string.Join(Environment.NewLine, copiedColumn)))
                {
                    return;
                }
                statusLabel.Text = string.Format("已复制 {0} 个单元格，可从任意新文件名行开始逐行粘贴。", copiedColumn.Length);
                return;
            }
            int minRow = selectedCells.Min(delegate(DataGridViewCell cell) { return cell.RowIndex; });
            int maxRow = selectedCells.Max(delegate(DataGridViewCell cell) { return cell.RowIndex; });
            int minColumn = selectedCells.Min(delegate(DataGridViewCell cell) { return cell.ColumnIndex; });
            int maxColumn = selectedCells.Max(delegate(DataGridViewCell cell) { return cell.ColumnIndex; });
            var copiedText = new StringBuilder();
            for (int rowIndex = minRow; rowIndex <= maxRow; rowIndex++)
            {
                if (rowIndex > minRow)
                {
                    copiedText.AppendLine();
                }
                for (int columnIndex = minColumn; columnIndex <= maxColumn; columnIndex++)
                {
                    if (columnIndex > minColumn)
                    {
                        copiedText.Append('\t');
                    }
                    var cell = filesGrid.Rows[rowIndex].Cells[columnIndex];
                    if (cell.Selected)
                    {
                        copiedText.Append(Convert.ToString(cell.Value ?? string.Empty));
                    }
                }
            }
            if (!TrySetClipboardText(copiedText.ToString()))
            {
                return;
            }
            statusLabel.Text = "已按 Excel 格式复制选中单元格。";
        }

        private void PasteExcelValuesIntoNewNames()
        {
            string clipboardText;
            if (!TryGetClipboardText(out clipboardText))
            {
                return;
            }
            int newNameIndex = filesGrid.Columns[NewNameColumn].Index;
            var selectedCells = filesGrid.SelectedCells.Cast<DataGridViewCell>().ToList();
            var selectedNewNameCells = selectedCells.Where(delegate(DataGridViewCell cell) { return cell.ColumnIndex == newNameIndex; }).ToList();
            if (selectedNewNameCells.Count == 0)
            {
                MessageBox.Show(this, "请先选择“新文件名”列中的一个单元格或点击该列表头。", "请选择粘贴位置", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            filesGrid.EndEdit();
            int startRow = selectedNewNameCells.Min(delegate(DataGridViewCell cell) { return cell.RowIndex; });
            var lines = clipboardText.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').ToList();
            if (lines.Count > 1 && lines[lines.Count - 1].Length == 0)
            {
                lines.RemoveAt(lines.Count - 1);
            }
            int pastedCount = 0;
            for (int lineIndex = 0; lineIndex < lines.Count && startRow + lineIndex < filesGrid.Rows.Count; lineIndex++)
            {
                var value = lines[lineIndex].Split('\t')[0].Trim();
                var row = filesGrid.Rows[startRow + lineIndex];
                row.Cells[NewNameColumn].Value = value;
                row.Cells[SelectColumn].Value = !string.IsNullOrWhiteSpace(value);
                pastedCount++;
            }
            CaptureVisibleDrafts();
            UpdateSelectAllHeader();
            statusLabel.Text = string.Format("已按 Excel 行顺序粘贴 {0} 个新文件名，非空行已自动勾选。", pastedCount);
        }

        private bool TryGetClipboardText(out string clipboardText)
        {
            clipboardText = string.Empty;
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    if (!Clipboard.ContainsText())
                    {
                        MessageBox.Show(this, "剪贴板中没有可粘贴的文本。", "无法粘贴", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return false;
                    }
                    clipboardText = Clipboard.GetText();
                    return true;
                }
                catch (ExternalException)
                {
                    Thread.Sleep(80);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "无法访问剪贴板", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            MessageBox.Show(this, "剪贴板正在被其他程序占用，请稍后重试。", "无法访问剪贴板", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        private bool TrySetClipboardText(string text)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    Clipboard.SetText(text);
                    return true;
                }
                catch (ExternalException)
                {
                    Thread.Sleep(80);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "无法访问剪贴板", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            MessageBox.Show(this, "剪贴板正在被其他程序占用，请稍后重试。", "无法访问剪贴板", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        private void FilesGridCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (filesGrid.IsCurrentCellDirty && filesGrid.CurrentCell != null && filesGrid.CurrentCell.ColumnIndex == filesGrid.Columns[SelectColumn].Index)
            {
                filesGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void FilesGridCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!isApplyingBulkSelection && e.RowIndex >= 0 && e.ColumnIndex == filesGrid.Columns[SelectColumn].Index)
            {
                UpdateSelectAllHeader();
            }
        }

        private void FilesGridCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == filesGrid.Columns[SelectColumn].Index)
            {
                BeginInvoke((MethodInvoker)delegate { UpdateSelectAllHeader(); });
            }
        }

        private void ClearNewNames()
        {
            foreach (DataGridViewRow row in filesGrid.Rows)
            {
                row.Cells[NewNameColumn].Value = string.Empty;
                row.Cells[SelectColumn].Value = false;
            }
            CaptureVisibleDrafts();
            UpdateSelectAllHeader();
            statusLabel.Text = "已清空当前列表的修改框并取消勾选。";
        }

        private void SetAllSelections(bool selected)
        {
            // Commit a pending row checkbox before writing the full filtered result.
            filesGrid.EndEdit();
            isApplyingBulkSelection = true;
            filesGrid.SuspendLayout();
            try
            {
                foreach (DataGridViewRow row in filesGrid.Rows)
                {
                    row.Cells[SelectColumn].Value = selected;
                }
            }
            finally
            {
                filesGrid.ResumeLayout();
                isApplyingBulkSelection = false;
            }
            filesGrid.InvalidateColumn(filesGrid.Columns[SelectColumn].Index);
            filesGrid.Update();
            CaptureVisibleDrafts();
            UpdateSelectAllHeader();
            statusLabel.Text = selected ? "已勾选当前筛选结果。" : "已取消勾选当前筛选结果。";
        }

        private bool AreAllVisibleRowsSelected()
        {
            return filesGrid.Rows.Count > 0
                && filesGrid.Rows.Cast<DataGridViewRow>().All(delegate(DataGridViewRow row)
                {
                    return Convert.ToBoolean(row.Cells[SelectColumn].Value ?? false);
                });
        }

        private void UpdateSelectAllHeader()
        {
            if (filesGrid == null || filesGrid.Columns.Count == 0)
            {
                return;
            }
            int selectedCount = filesGrid.Rows.Cast<DataGridViewRow>().Count(delegate(DataGridViewRow row)
            {
                return Convert.ToBoolean(row.Cells[SelectColumn].Value ?? false);
            });
            string headerText = selectedCount == filesGrid.Rows.Count && selectedCount > 0 ? "☑ 全选" : "☐ 全选";
            filesGrid.Columns[SelectColumn].HeaderCell.Value = headerText;
            filesGrid.InvalidateColumn(filesGrid.Columns[SelectColumn].Index);
            filesGrid.InvalidateCell(filesGrid.Columns[SelectColumn].HeaderCell);
            SynchronizeSelectAllCheckBox();
        }

        private void BatchRename()
        {
            filesGrid.EndEdit();
            var entries = new List<RenameEntry>();
            var errors = new List<string>();
            string folderPath;
            try
            {
                folderPath = Path.GetFullPath(folderTextBox.Text.Trim());
            }
            catch (Exception)
            {
                MessageBox.Show(this, "文件夹路径无效。", "无法更改", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (DataGridViewRow row in filesGrid.Rows)
            {
                bool selected = Convert.ToBoolean(row.Cells[SelectColumn].Value ?? false);
                string requestedName = Convert.ToString(row.Cells[NewNameColumn].Value ?? string.Empty).Trim();
                if (!selected || requestedName.Length == 0)
                {
                    continue;
                }
                var sourcePath = row.Tag as string;
                if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                {
                    errors.Add(string.Format("找不到源文件：{0}", Convert.ToString(row.Cells[CurrentNameColumn].Value)));
                    continue;
                }
                try
                {
                    var finalName = NormalizeRequestedName(requestedName, Path.GetExtension(sourcePath), modifyExtensionCheckBox.Checked);
                    if (string.Equals(Path.GetFileName(sourcePath), finalName, StringComparison.Ordinal))
                    {
                        continue;
                    }
                    entries.Add(new RenameEntry
                    {
                        Source = sourcePath,
                        Target = Path.Combine(folderPath, finalName)
                    });
                }
                catch (ArgumentException ex)
                {
                    errors.Add(string.Format("{0}：{1}", Convert.ToString(row.Cells[CurrentNameColumn].Value), ex.Message));
                }
            }

            ValidateEntries(entries, errors);
            if (errors.Count > 0)
            {
                var errorText = string.Join(Environment.NewLine, errors.Take(10).ToArray());
                if (errors.Count > 10)
                {
                    errorText += Environment.NewLine + string.Format("另有 {0} 项错误。", errors.Count - 10);
                }
                MessageBox.Show(this, errorText, "请先修正以下问题", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (entries.Count == 0)
            {
                MessageBox.Show(this, "没有可更改的项目。请勾选文件并填写新文件名。", "无需更改", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var preview = new StringBuilder();
            foreach (var entry in entries.Take(8))
            {
                preview.AppendLine(Path.GetFileName(entry.Source) + "  ->  " + Path.GetFileName(entry.Target));
            }
            if (entries.Count > 8)
            {
                preview.AppendLine(string.Format("……以及另外 {0} 项", entries.Count - 8));
            }
            var confirmation = MessageBox.Show(
                this,
                string.Format("将更改 {0} 个文件名：{1}{2}{1}此操作会立即修改文件名，是否继续？", entries.Count, Environment.NewLine, preview),
                "确认批量更改",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.Yes)
            {
                return;
            }

            ExecuteRename(entries);
        }

        private static string NormalizeRequestedName(string requestedName, string originalExtension, bool allowExtensionChange)
        {
            if (requestedName == "." || requestedName == "..")
            {
                throw new ArgumentException("文件名无效。");
            }
            if (requestedName.EndsWith(".", StringComparison.Ordinal) || requestedName.EndsWith(" ", StringComparison.Ordinal))
            {
                throw new ArgumentException("文件名不能以句点或空格结尾。");
            }
            if (requestedName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || requestedName.IndexOf('\\') >= 0 || requestedName.IndexOf('/') >= 0)
            {
                throw new ArgumentException("包含不允许的文件名字符。");
            }
            string finalName;
            if (!allowExtensionChange)
            {
                var baseName = requestedName;
                if (!string.IsNullOrEmpty(originalExtension) && baseName.EndsWith(originalExtension, StringComparison.OrdinalIgnoreCase))
                {
                    baseName = baseName.Substring(0, baseName.Length - originalExtension.Length);
                }
                finalName = baseName + originalExtension;
            }
            else
            {
                finalName = string.IsNullOrEmpty(Path.GetExtension(requestedName))
                    ? requestedName + originalExtension
                    : requestedName;
            }
            if (finalName.Length > 240)
            {
                throw new ArgumentException("文件名过长。");
            }
            return finalName;
        }

        private static void ValidateEntries(List<RenameEntry> entries, List<string> errors)
        {
            var targetPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sourcePaths = new HashSet<string>(entries.Select(delegate(RenameEntry entry) { return entry.Source; }), StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries)
            {
                if (!targetPaths.Add(entry.Target))
                {
                    errors.Add(string.Format("目标文件名重复：{0}", Path.GetFileName(entry.Target)));
                }
                if (File.Exists(entry.Target) && !sourcePaths.Contains(entry.Target))
                {
                    errors.Add(string.Format("目标文件已存在：{0}", Path.GetFileName(entry.Target)));
                }
            }
        }

        private void ExecuteRename(List<RenameEntry> entries)
        {
            var movedToTemp = new List<RenameEntry>();
            var movedToTarget = new List<RenameEntry>();
            try
            {
                foreach (var entry in entries)
                {
                    entry.Temp = Path.Combine(Path.GetDirectoryName(entry.Source), ".rename_tmp_" + Guid.NewGuid().ToString("N"));
                    File.Move(entry.Source, entry.Temp);
                    movedToTemp.Add(entry);
                }
                foreach (var entry in entries)
                {
                    File.Move(entry.Temp, entry.Target);
                    movedToTarget.Add(entry);
                }
                drafts.Clear();
                RefreshFiles(false);
                MessageBox.Show(this, string.Format("已成功更改 {0} 个文件名。", entries.Count), "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                RollBackRename(movedToTarget, movedToTemp);
                RefreshFiles(false);
                MessageBox.Show(this, "批量更改未完成，已尝试恢复原文件名。" + Environment.NewLine + ex.Message, "更改失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RollBackRename(List<RenameEntry> movedToTarget, List<RenameEntry> movedToTemp)
        {
            foreach (var entry in movedToTarget.AsEnumerable().Reverse())
            {
                try
                {
                    if (File.Exists(entry.Target) && !File.Exists(entry.Source))
                    {
                        File.Move(entry.Target, entry.Source);
                    }
                }
                catch (Exception)
                {
                }
            }
            foreach (var entry in movedToTemp.AsEnumerable().Reverse())
            {
                try
                {
                    if (File.Exists(entry.Temp) && !File.Exists(entry.Source))
                    {
                        File.Move(entry.Temp, entry.Source);
                    }
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
