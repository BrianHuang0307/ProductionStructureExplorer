using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ProductionStructureExplorer
{
    public partial class MainForm
    {
        private void BtnLoadBom_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls", Title = "選擇 BOM 檔" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                _bomRaw = _excelService.LoadBom(ofd.FileName, out _featureCode, out _rootCode);
                var order = BomService.BuildDfsOrder(_bomRaw);
                _bomSorted = BomService.ApplyOrder(_bomRaw, order);
                _childrenByParent = BomService.BuildChildrenMap(_bomSorted);

                BuildBomColumns();
                ShowOnlyLevel1();

                txtRoot.Text = _rootCode;
                txtFeature.Text = _featureCode;
                lblInfo.Text = $"已載入 BOM：{_bomSorted.Rows.Count} 列。";
            }
            catch (Exception ex) { MessageBox.Show("載入失敗: " + ex.Message); }
        }

        private void BtnLoadRouting_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.csv", Title = "選擇 Routing 檔" };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                _routingRaw = _excelService.LoadRouting(ofd.FileName);
                _routingByChild = RoutingService.BuildRoutingMap(_routingRaw);

                BuildRoutingColumns();

                // 刷新左側
                if (_bomSorted != null && _bomSorted.Rows.Count > 0)
                {
                    foreach (DataGridViewRow row in dgvBom.Rows)
                    {
                        if (row.Tag is BomNodeState state)
                        {
                            string child = state.Row[COL_CHILD]?.ToString() ?? "";
                            if (_routingByChild.ContainsKey(child))
                            {
                                row.DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;
                                row.DefaultCellStyle.Font = new System.Drawing.Font(dgvBom.Font, System.Drawing.FontStyle.Bold);
                            }
                        }
                    }
                }

                lblInfo.Text = $"已載入 Routing：{_routingRaw.Rows.Count} 列。";
            }
            catch (Exception ex) { MessageBox.Show("載入失敗: " + ex.Message); }
        }

        // --- 核心修改：使用中文鍵值更新右側表格 ---
        private void DgvBom_SelectionChanged(object? sender, EventArgs e)
        {
            try
            {
                dgvRouting.Rows.Clear();

                if (dgvBom.SelectedRows.Count == 0) return;

                var selectedRow = dgvBom.SelectedRows[0];
                if (selectedRow.Tag is not BomNodeState state) return;

                string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";

                if (!string.IsNullOrEmpty(childKey) && _routingByChild.TryGetValue(childKey, out var routingRows))
                {
                    // 排序：嘗試使用 "工序" 欄位排序
                    var sortedRouting = routingRows.OrderBy(r =>
                    {
                        string val = GetSafeValue(r, COL_R_OPSEQ);
                        if (int.TryParse(val, out int seq)) return seq;
                        return 9999;
                    });

                    foreach (var rRow in sortedRouting)
                    {
                        int idx = dgvRouting.Rows.Add();
                        var gridRow = dgvRouting.Rows[idx];

                        // 使用 UI 定義的中文欄位名稱填入資料
                        gridRow.Cells[COL_R_OPSEQ].Value = GetSafeValue(rRow, COL_R_OPSEQ);
                        gridRow.Cells[COL_R_DESC].Value = GetSafeValue(rRow, COL_R_DESC);
                        gridRow.Cells[COL_R_MANPOWER].Value = GetSafeValue(rRow, COL_R_MANPOWER);
                        gridRow.Cells[COL_R_WC].Value = GetSafeValue(rRow, COL_R_WC);
                        gridRow.Cells[COL_R_SETUP].Value = GetSafeValue(rRow, COL_R_SETUP);
                        gridRow.Cells[COL_R_RUN].Value = GetSafeValue(rRow, COL_R_RUN);
                        gridRow.Cells[COL_R_MACH].Value = GetSafeValue(rRow, COL_R_MACH);
                    }
                }
            }
            catch (Exception)
            {
                // Ignore
            }
        }

        private string GetSafeValue(DataRow row, string colName)
        {
            // 精確比對
            if (row.Table.Columns.Contains(colName))
                return row[colName]?.ToString() ?? "";

            // 模糊比對
            foreach (DataColumn col in row.Table.Columns)
            {
                if (col.ColumnName.Trim().Equals(colName.Trim(), StringComparison.OrdinalIgnoreCase))
                    return row[col]?.ToString() ?? "";
            }

            return "";
        }

        private void DgvBom_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            if (dgvBom.Columns[e.ColumnIndex].Name == COL_EXPAND)
            {
                var row = dgvBom.Rows[e.RowIndex];
                if (row.Tag is not BomNodeState state) return;

                string childKey = state.Row[COL_CHILD]?.ToString() ?? "";
                if (!HasBomChildren(childKey)) return;

                if (state.IsExpanded) CollapseRow(e.RowIndex);
                else ExpandRow(e.RowIndex);
            }
        }

        private void ExpandRow(int rowIndex)
        {
            var row = dgvBom.Rows[rowIndex];
            var state = row.Tag as BomNodeState;
            int currentLevel = ParseInt(state.Row[COL_LEVEL]);
            int insertIndex = rowIndex + 1;
            int sourceIndex = state.SourceIndex;

            if (sourceIndex >= 0 && sourceIndex < _bomSorted.Rows.Count - 1)
            {
                for (int i = sourceIndex + 1; i < _bomSorted.Rows.Count; i++)
                {
                    DataRow nextRow = _bomSorted.Rows[i];
                    int nextLevel = ParseInt(nextRow[COL_LEVEL]);

                    if (nextLevel <= currentLevel) break;

                    if (nextLevel == currentLevel + 1)
                    {
                        AddBomGridRow(nextRow, i, insertIndex);
                        insertIndex++;
                    }
                }
            }
            state.IsExpanded = true;
            row.Cells[COL_EXPAND].Value = "-";
        }

        private void CollapseRow(int rowIndex)
        {
            var row = dgvBom.Rows[rowIndex];
            var state = row.Tag as BomNodeState;
            int myLevel = ParseInt(state.Row[COL_LEVEL]);
            int i = rowIndex + 1;

            while (i < dgvBom.Rows.Count)
            {
                var gr = dgvBom.Rows[i];
                if (gr.Tag is not BomNodeState s) { i++; continue; }

                int level = ParseInt(s.Row[COL_LEVEL]);
                if (level <= myLevel) break;

                dgvBom.Rows.RemoveAt(i);
            }
            state.IsExpanded = false;
            row.Cells[COL_EXPAND].Value = "+";
        }

        private void BtnExpandAll_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null) return;
            dgvBom.Rows.Clear();
            for (int i = 0; i < _bomSorted.Rows.Count; i++)
            {
                DataRow row = _bomSorted.Rows[i];
                AddBomGridRow(row, i);

                var gridRow = dgvBom.Rows[dgvBom.Rows.Count - 1];
                var state = gridRow.Tag as BomNodeState;
                string childKey = row[COL_CHILD]?.ToString() ?? "";

                if (HasBomChildren(childKey))
                {
                    state.IsExpanded = true;
                    gridRow.Cells[COL_EXPAND].Value = "-";
                }
            }
        }

        private void BtnCollapseAll_Click(object? sender, EventArgs e)
        {
            ShowOnlyLevel1();
        }

        private void BtnExport_Click(object? sender, EventArgs e)
        {
            _excelService.ExportGridToExcel(dgvBom, "BOM_Tree.xlsx");
        }
    }
}