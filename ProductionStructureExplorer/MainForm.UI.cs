using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ProductionStructureExplorer
{
    public partial class MainForm : Form
    {
        // UI Controls
        private SplitContainer splitContainer;
        private DataGridView dgvBom;
        private DataGridView dgvRouting;

        private Button btnLoadBom;
        private Button btnLoadRouting;
        private Button btnExpandAll;
        private Button btnCollapseAll;
        private Button btnExport;
        private Label lblInfo;

        private Label lblFeature;
        private TextBox txtFeature;
        private Label lblRoot;
        private TextBox txtRoot;

        private readonly IExcelService _excelService;

        // Data Storage
        internal DataTable _bomRaw = new();
        internal DataTable _bomSorted = new();
        internal Dictionary<string, List<DataRow>> _childrenByParent = new();

        internal DataTable _routingRaw = new();
        internal Dictionary<string, List<DataRow>> _routingByChild = new();

        internal string _featureCode = "";
        internal string _rootCode = "";

        // --- Constants ---
        internal const string COL_PARENT = "PARENT";
        internal const string COL_CHILD = "CHILD";
        internal const string COL_LEVEL = "階層";
        internal const string COL_EXPAND = "_EXPAND_";

        // Routing 中文欄位定義 (對應 Excel 標頭)
        internal const string COL_R_OPSEQ = "工序";
        internal const string COL_R_DESC = "說明";
        internal const string COL_R_MANPOWER = "單位人力";
        internal const string COL_R_WC = "工作中心";
        internal const string COL_R_SETUP = "準備工時";
        internal const string COL_R_RUN = "標準工時";
        internal const string COL_R_MACH = "機器工時";

        // Grid Tag State
        internal class BomNodeState
        {
            public DataRow Row { get; set; } = null!;
            public bool IsExpanded { get; set; } = false;
            public int SourceIndex { get; set; } = -1;
        }

        public MainForm()
        {
            _excelService = new ExcelService();

            Text = "Production Structure Explorer (Master-Detail)";
            Width = 1400;
            Height = 800;
            StartPosition = FormStartPosition.CenterScreen;

            // Initialize Controls
            btnLoadBom = CreateBtn("載入 BOM", 20);
            btnLoadRouting = CreateBtn("載入 Routing", 140);
            btnExpandAll = CreateBtn("全部展開", 260);
            btnCollapseAll = CreateBtn("全部收合", 380);
            btnExport = CreateBtn("匯出 Excel", 500);

            lblFeature = new Label { Left = 630, Top = 24, Width = 60, Text = "特性編碼:" };
            txtFeature = new TextBox { Left = 690, Top = 20, Width = 100 };
            lblRoot = new Label { Left = 810, Top = 24, Width = 50, Text = "ROOT:" };
            txtRoot = new TextBox { Left = 860, Top = 20, Width = 120 };
            lblInfo = new Label { Left = 20, Top = 60, Width = 1300, Height = 30, Text = "請載入資料..." };

            // SplitContainer
            splitContainer = new SplitContainer
            {
                Left = 20,
                Top = 100,
                Width = 1340,
                Height = 640,
                Orientation = Orientation.Vertical,
                SplitterDistance = 600,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            // Left Grid (BOM) - [關鍵修正] 加入 EditMode 設定
            dgvBom = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                // [關鍵] 禁止進入編輯模式，確保點擊事件 100% 觸發
                EditMode = DataGridViewEditMode.EditProgrammatically
            };

            // Right Grid (Routing)
            dgvRouting = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.WhiteSmoke,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            splitContainer.Panel1.Controls.Add(dgvBom);
            splitContainer.Panel2.Controls.Add(dgvRouting);

            Controls.AddRange(new Control[] {
                btnLoadBom, btnLoadRouting, btnExpandAll, btnCollapseAll, btnExport,
                lblRoot, txtRoot, lblFeature, txtFeature, lblInfo, splitContainer
            });

            // Events
            btnLoadBom.Click += BtnLoadBom_Click;
            btnLoadRouting.Click += BtnLoadRouting_Click;
            btnExpandAll.Click += BtnExpandAll_Click;
            btnCollapseAll.Click += BtnCollapseAll_Click;
            btnExport.Click += BtnExport_Click;

            dgvBom.SelectionChanged += DgvBom_SelectionChanged;
            dgvBom.CellClick += DgvBom_CellClick;
        }

        private Button CreateBtn(string text, int left)
        {
            return new Button { Text = text, Left = left, Top = 20, Width = 100, Height = 30 };
        }

        internal void BuildBomColumns()
        {
            dgvBom.Columns.Clear();

            // [優化] 展開欄位設為置中，像按鈕一樣
            var expandCol = new DataGridViewTextBoxColumn
            {
                Name = COL_EXPAND,
                HeaderText = "",
                Width = 30,
                ReadOnly = true
            };
            expandCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            expandCol.DefaultCellStyle.Font = new Font("Consolas", 10, FontStyle.Bold); // 用等寬字體讓 + 號漂亮一點
            dgvBom.Columns.Add(expandCol);

            foreach (DataColumn col in _bomSorted.Columns)
            {
                if (new[] { COL_PARENT, "組合項次", "FULL_PATH", "特性編碼", "ROOT", "是否使用", "SOURCE" }
                    .Any(x => x.Equals(col.ColumnName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var gridCol = new DataGridViewTextBoxColumn
                {
                    Name = col.ColumnName,
                    HeaderText = col.ColumnName == COL_CHILD ? "料號" : col.ColumnName,
                    ReadOnly = true
                };
                dgvBom.Columns.Add(gridCol);
            }
        }

        internal void BuildRoutingColumns()
        {
            dgvRouting.Columns.Clear();

            // 使用中文常數
            AddRoutingCol(COL_R_OPSEQ, COL_R_OPSEQ);
            AddRoutingCol(COL_R_DESC, COL_R_DESC);
            AddRoutingCol(COL_R_MANPOWER, COL_R_MANPOWER);
            AddRoutingCol(COL_R_WC, COL_R_WC);
            AddRoutingCol(COL_R_SETUP, COL_R_SETUP);
            AddRoutingCol(COL_R_RUN, COL_R_RUN);
            AddRoutingCol(COL_R_MACH, COL_R_MACH);
        }

        private void AddRoutingCol(string name, string header)
        {
            dgvRouting.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                ReadOnly = true
            });
        }

        internal bool HasBomChildren(string childKey)
        {
            if (string.IsNullOrEmpty(childKey)) return false;
            return _childrenByParent.ContainsKey(childKey) && _childrenByParent[childKey].Count > 0;
        }

        internal int AddBomGridRow(DataRow srcRow, int sourceIdx, int? insertIndex = null)
        {
            var newRow = new DataGridViewRow();
            newRow.CreateCells(dgvBom);

            var state = new BomNodeState { Row = srcRow, IsExpanded = false, SourceIndex = sourceIdx };
            newRow.Tag = state;

            string childKey = srcRow[COL_CHILD]?.ToString()?.Trim() ?? "";
            newRow.Cells[dgvBom.Columns[COL_EXPAND].Index].Value = HasBomChildren(childKey) ? "+" : "";

            if (_routingByChild.ContainsKey(childKey))
            {
                newRow.DefaultCellStyle.ForeColor = Color.Blue;
                newRow.DefaultCellStyle.Font = new Font(dgvBom.Font, FontStyle.Bold);
            }

            FillRowData(newRow, srcRow, dgvBom);

            if (insertIndex.HasValue && insertIndex.Value >= 0 && insertIndex.Value <= dgvBom.Rows.Count)
            {
                dgvBom.Rows.Insert(insertIndex.Value, newRow);
                return insertIndex.Value;
            }
            return dgvBom.Rows.Add(newRow);
        }

        private void FillRowData(DataGridViewRow gridRow, DataRow dataRow, DataGridView grid)
        {
            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (col.Name == COL_EXPAND) continue;

                object? value = null;
                if (dataRow.Table.Columns.Contains(col.Name))
                {
                    value = dataRow[col.Name];
                }
                gridRow.Cells[col.Index].Value = value == null || value == DBNull.Value ? "" : value.ToString();
            }
        }

        internal void ShowOnlyLevel1()
        {
            dgvBom.Rows.Clear();
            for (int i = 0; i < _bomSorted.Rows.Count; i++)
            {
                DataRow row = _bomSorted.Rows[i];
                int level = ParseInt(row[COL_LEVEL]);
                if (level == 1)
                {
                    AddBomGridRow(row, i);
                }
            }
        }

        internal int ParseInt(object? v)
        {
            if (v == null || v == DBNull.Value) return int.MaxValue;
            var s = v.ToString()?.Trim();
            if (int.TryParse(s, out var n)) return n;
            return int.MaxValue;
        }
    }
}