using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ProductionStructureExplorer
{
    public partial class MainForm : Form
    {
        private Button btnLoadBom;
        private Button btnLoadRouting;
        private Button btnExpandAll;
        private Button btnCollapseAll;
        private Button btnExport;
        private DataGridView dgv;
        private Label lblInfo;
        private Button btnSearchDB;

        private Label lblFeature;
        private TextBox txtFeature;
        private Label lblRoot;
        private TextBox txtRoot;

        private readonly IExcelService _excelService;

        // BOM / Routing 資料
        internal DataTable _bomRaw = new();
        internal DataTable _bomSorted = new();
        internal Dictionary<string, List<DataRow>> _childrenByParent = new();

        internal DataTable _routingRaw = new();
        internal Dictionary<string, List<DataRow>> _routingByChild = new();

        internal string _featureCode = "";
        internal string _rootCode = "";

        // 欄位名稱
        internal const string COL_PARENT = "PARENT";
        internal const string COL_CHILD = "CHILD";
        internal const string COL_LEVEL = "階層";
        internal const string COL_OPSEQ = "工序";
        internal const string COL_EXPAND = "_EXPAND_";

        internal const string COL_ROOT = "ROOT";
        internal const string COL_FULLPATH = "FULL_PATH";
        internal const string COL_FEATURE = "特性編碼";
        internal const string COL_ORDER = "組合項次";
        internal const string COL_USED = "是否使用";

        // Grid 用 Tag 狀態
        internal class BomNodeState
        {
            public DataRow Row { get; set; } = null!;
            public bool IsExpanded { get; set; } = false;
            public bool IsRouting { get; set; } = false;
        }

        public MainForm()
        {
            _excelService = new ExcelService();

            Text = "ProductionStructureExplorer";
            Width = 1200;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            btnLoadBom = new Button
            {
                Text = "載入 BOM",
                Left = 20,
                Top = 20,
                Width = 100,
                Height = 30
            };
            btnLoadRouting = new Button
            {
                Text = "載入 Routing",
                Left = 140,
                Top = 20,
                Width = 100,
                Height = 30
            };
            btnExpandAll = new Button
            {
                Text = "全部展開",
                Left = 260,
                Top = 20,
                Width = 100,
                Height = 30
            };
            btnCollapseAll = new Button
            {
                Text = "全部收合",
                Left = 380,
                Top = 20,
                Width = 100,
                Height = 30
            };
            btnExport = new Button
            {
                Text = "匯出 Excel",
                Left = 500,
                Top = 20,
                Width = 100,
                Height = 30
            };

            lblFeature = new Label
            {
                Left = 630,
                Top = 24,
                Width = 60,
                Height = 20,
                Text = "特性編碼:"
            };
            txtFeature = new TextBox
            {
                Left = 690,
                Top = 20,
                Width = 100,
                Height = 24,
                //ReadOnly = true
            };

            lblRoot = new Label
            {
                Left = 810,
                Top = 24,
                Width = 50,
                Height = 20,
                Text = "ROOT:"
            };
            txtRoot = new TextBox
            {
                Left = 860,
                Top = 20,
                Width = 120,
                Height = 24,
                //ReadOnly = true
            };

            btnSearchDB = new Button
            {
                Text = "查詢",
                Left = 1000,
                Top = 20,
                Width = 100,
                Height = 30
            };

            lblInfo = new Label
            {
                Left = 20,
                Top = 60,
                Width = 1100,
                Height = 30,
                Text = "請先載入 BOM，必要欄位：階層 / ROOT / PARENT / CHILD / 組合項次 / 特性編碼 / FULL_PATH / 是否使用 / 失效日期。"
            };

            dgv = new DataGridView
            {
                Left = 20,
                Top = 100,
                Width = 1140,
                Height = 560,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            };

            Controls.Add(btnLoadBom);
            Controls.Add(btnLoadRouting);
            Controls.Add(btnExpandAll);
            Controls.Add(btnCollapseAll);
            Controls.Add(btnExport);
            Controls.Add(lblRoot);
            Controls.Add(txtRoot);
            Controls.Add(lblFeature);
            Controls.Add(txtFeature);
            Controls.Add(lblInfo);
            Controls.Add(dgv);
            Controls.Add(btnSearchDB);

            btnLoadBom.Click += BtnLoadBom_Click;
            btnLoadRouting.Click += BtnLoadRouting_Click;
            btnExpandAll.Click += BtnExpandAll_Click;
            btnCollapseAll.Click += BtnCollapseAll_Click;
            btnExport.Click += BtnExport_Click;
            dgv.CellClick += Dgv_CellClick;
        }

        internal void BuildGridColumns()
        {
            dgv.Columns.Clear();

            var expandCol = new DataGridViewTextBoxColumn
            {
                Name = COL_EXPAND,
                HeaderText = "",
                Width = 30,
                ReadOnly = true
            };
            dgv.Columns.Add(expandCol);

            // 其他欄位從 _bomSorted 建立，排除 helper 欄位
            foreach (DataColumn col in _bomSorted.Columns)
            {
                if (string.Equals(col.ColumnName, COL_PARENT, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_ORDER, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_FULLPATH, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_FEATURE, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_ROOT, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.ColumnName, COL_USED, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var gridCol = new DataGridViewTextBoxColumn
                {
                    Name = col.ColumnName,
                    HeaderText = col.ColumnName,
                    ReadOnly = true
                };

                if (string.Equals(col.ColumnName, COL_CHILD, StringComparison.OrdinalIgnoreCase))
                    gridCol.HeaderText = "料號";

                dgv.Columns.Add(gridCol);
            }

            // 工序欄（給 Routing 用）
            if (!dgv.Columns.Contains(COL_OPSEQ))
            {
                var opCol = new DataGridViewTextBoxColumn
                {
                    Name = COL_OPSEQ,
                    HeaderText = COL_OPSEQ,
                    ReadOnly = true
                };
                dgv.Columns.Add(opCol);
            }
        }

        internal int AddBomGridRow(DataRow srcRow, int? insertIndex = null)
        {
            int rowIndex;
            if (insertIndex.HasValue)
            {
                dgv.Rows.Insert(insertIndex.Value, 1);
                rowIndex = insertIndex.Value;
            }
            else
            {
                rowIndex = dgv.Rows.Add();
            }

            var gr = dgv.Rows[rowIndex];

            var state = new BomNodeState { Row = srcRow, IsExpanded = false, IsRouting = false };
            gr.Tag = state;

            string childKey = srcRow[COL_CHILD]?.ToString()?.Trim() ?? "";
            bool hasBomChildren = !string.IsNullOrEmpty(childKey) &&
                                  _childrenByParent.ContainsKey(childKey) &&
                                  _childrenByParent[childKey].Count > 0;
            bool hasRouting = !string.IsNullOrEmpty(childKey) &&
                              _routingByChild.ContainsKey(childKey) &&
                              _routingByChild[childKey].Count > 0;

            gr.Cells[COL_EXPAND].Value = (hasBomChildren || hasRouting) ? "+" : "";

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Name == COL_EXPAND) continue;

                object? value = null;
                if (srcRow.Table.Columns.Contains(col.Name))
                {
                    value = srcRow[col.Name];
                }

                gr.Cells[col.Name].Value =
                    value == null || value == DBNull.Value ? "" : value.ToString();
            }

            return rowIndex;
        }

        internal int AddRoutingGridRow(DataRow routingRow, string baseLevelText, int? insertIndex = null)
        {
            int rowIndex;
            if (insertIndex.HasValue)
            {
                dgv.Rows.Insert(insertIndex.Value, 1);
                rowIndex = insertIndex.Value;
            }
            else
            {
                rowIndex = dgv.Rows.Add();
            }

            var gr = dgv.Rows[rowIndex];

            var state = new BomNodeState { Row = routingRow, IsExpanded = false, IsRouting = true };
            gr.Tag = state;

            gr.Cells[COL_EXPAND].Value = "";

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Name == COL_EXPAND) continue;

                if (col.Name == COL_LEVEL)
                {
                    gr.Cells[col.Name].Value = string.IsNullOrEmpty(baseLevelText)
                        ? "R"
                        : baseLevelText + "R";
                    continue;
                }

                object? value = null;
                if (routingRow.Table.Columns.Contains(col.Name))
                {
                    value = routingRow[col.Name];
                }

                gr.Cells[col.Name].Value =
                    value == null || value == DBNull.Value ? "" : value.ToString();
            }

            return rowIndex;
        }

        internal int InsertRoutingRows(string childKey, string baseLevelText, int? insertIndex)
        {
            if (string.IsNullOrEmpty(childKey)) return insertIndex ?? dgv.Rows.Count;
            if (!_routingByChild.TryGetValue(childKey, out var ops) || ops.Count == 0)
                return insertIndex ?? dgv.Rows.Count;

            int idx = insertIndex ?? dgv.Rows.Count;

            foreach (var opRow in ops)
            {
                bool alreadyVisible = false;
                foreach (DataGridViewRow gr in dgv.Rows)
                {
                    if (gr.Tag is BomNodeState s && s.IsRouting && s.Row == opRow)
                    {
                        alreadyVisible = true;
                        break;
                    }
                }
                if (alreadyVisible) continue;

                AddRoutingGridRow(opRow, baseLevelText, idx);
                idx++;
            }

            return idx;
        }

        internal void ShowOnlyLevel1()
        {
            dgv.Rows.Clear();

            foreach (DataRow row in _bomSorted.Rows)
            {
                int level = ParseInt(row[COL_LEVEL]);
                if (level == 1)
                {
                    AddBomGridRow(row);
                }
            }
        }

        internal int ParseInt(object? v)
        {
            if (v == null || v == DBNull.Value) return int.MaxValue;
            var s = v.ToString()?.Trim();

            if (!string.IsNullOrEmpty(s))
            {
                int i = 0;
                while (i < s.Length && char.IsDigit(s[i])) i++;
                if (i > 0 && int.TryParse(s.Substring(0, i), out var n))
                    return n;
            }

            return int.TryParse(s, out var n2) ? n2 : int.MaxValue;
        }
    }
}
