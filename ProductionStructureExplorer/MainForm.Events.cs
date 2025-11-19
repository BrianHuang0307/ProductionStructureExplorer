using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionStructureExplorer
{
    public partial class MainForm
    {
        private void BtnLoadBom_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls",
                Title = "選擇 BOM 檔"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                _bomRaw = _excelService.LoadBom(ofd.FileName, out _featureCode, out _rootCode);

                var order = BomService.BuildDfsOrder(_bomRaw);
                _bomSorted = BomService.ApplyOrder(_bomRaw, order);

                _childrenByParent = BomService.BuildChildrenMap(_bomSorted);

                BuildGridColumns();
                ShowOnlyLevel1();

                txtRoot.Text = _rootCode;
                txtFeature.Text = _featureCode;

                lblInfo.Text = $"已載入 BOM：{_bomSorted.Rows.Count} 列 (已排除失效日期有值的列)。目前只顯示階層 = 1。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入 / 處理 BOM 失敗： " + ex.Message);
            }
        }

        private void BtnLoadRouting_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls",
                Title = "選擇 Routing 檔"
            };
            if (ofd.ShowDialog() != DialogResult.OK) return;

            try
            {
                _routingRaw = _excelService.LoadRouting(ofd.FileName);
                _routingByChild = RoutingService.BuildRoutingMap(_routingRaw);

                lblInfo.Text = $"已載入 Routing：{_routingRaw.Rows.Count} 列（依 CHILD + 工序 可掛到 BOM 上）。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入 Routing 失敗： " + ex.Message);
            }
        }

        private void BtnExpandAll_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null || _bomSorted.Rows.Count == 0)
            {
                MessageBox.Show("請先載入 BOM。");
                return;
            }

            dgv.Rows.Clear();

            foreach (DataRow row in _bomSorted.Rows)
            {
                AddBomGridRow(row);
                string childKey = row[COL_CHILD]?.ToString()?.Trim() ?? "";
                string levelText = row[COL_LEVEL]?.ToString()?.Trim() ?? "";

                InsertRoutingRows(childKey, levelText, null);
            }

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                var gr = dgv.Rows[i];
                if (gr.Tag is BomNodeState state && !state.IsRouting)
                {
                    string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(childKey) &&
                        _childrenByParent.ContainsKey(childKey) &&
                        _childrenByParent[childKey].Count > 0)
                    {
                        state.IsExpanded = true;
                        gr.Cells[COL_EXPAND].Value = "-";
                    }
                }
            }

            lblInfo.Text = $"全部展開：顯示 {dgv.Rows.Count} 列（含 Routing 製程）。";
        }

        private void BtnCollapseAll_Click(object? sender, EventArgs e)
        {
            if (_bomSorted == null || _bomSorted.Rows.Count == 0)
            {
                MessageBox.Show("請先載入 BOM。");
                return;
            }

            ShowOnlyLevel1();
            lblInfo.Text = "已收合：只顯示階層 = 1（不顯示 Routing）。";
        }

        private void BtnExport_Click(object? sender, EventArgs e)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("目前畫面沒有資料可以匯出。");
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "BomRouting_View.xlsx"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                //  直接匯出當前的 DataGridView 狀態
                _excelService.ExportGridToExcel(dgv, sfd.FileName);
                MessageBox.Show("匯出完成！（以目前 Grid 顯示為準）");
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗： " + ex.Message);
            }
        }

        private void Dgv_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var col = dgv.Columns[e.ColumnIndex];
            if (col.Name != COL_EXPAND) return;

            var row = dgv.Rows[e.RowIndex];
            if (row.Tag is not BomNodeState state) return;

            if (state.IsRouting) return;

            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            bool hasBomChildren = !string.IsNullOrEmpty(childKey) &&
                                  _childrenByParent.ContainsKey(childKey) &&
                                  _childrenByParent[childKey].Count > 0;
            bool hasRouting = !string.IsNullOrEmpty(childKey) &&
                              _routingByChild.ContainsKey(childKey) &&
                              _routingByChild[childKey].Count > 0;

            if (!hasBomChildren && !hasRouting)
                return;

            if (state.IsExpanded)
                CollapseRow(e.RowIndex);
            else
                ExpandRow(e.RowIndex);
        }

        private void ExpandRow(int rowIndex)
        {
            var row = dgv.Rows[rowIndex];
            if (row.Tag is not BomNodeState state) return;
            if (state.IsRouting) return;

            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(childKey))
                return;

            string levelText = state.Row[COL_LEVEL]?.ToString()?.Trim() ?? "";

            int insertIndex = rowIndex + 1;

            insertIndex = InsertRoutingRows(childKey, levelText, insertIndex);

            if (_childrenByParent.TryGetValue(childKey, out var children) && children.Count > 0)
            {
                foreach (var childRow in children)
                {
                    bool alreadyVisible = false;
                    foreach (DataGridViewRow gr in dgv.Rows)
                    {
                        if (gr.Tag is BomNodeState s && !s.IsRouting && s.Row == childRow)
                        {
                            alreadyVisible = true;
                            break;
                        }
                    }
                    if (alreadyVisible) continue;

                    AddBomGridRow(childRow, insertIndex);
                    insertIndex++;
                }
            }

            state.IsExpanded = true;
            row.Cells[COL_EXPAND].Value = "-";
        }

        private void CollapseRow(int rowIndex)
        {
            var row = dgv.Rows[rowIndex];
            if (row.Tag is not BomNodeState state) return;
            if (state.IsRouting) return;

            int myLevel = ParseInt(state.Row[COL_LEVEL]);
            int i = rowIndex + 1;

            while (i < dgv.Rows.Count)
            {
                var gr = dgv.Rows[i];
                if (gr.Tag is not BomNodeState s)
                {
                    i++;
                    continue;
                }

                if (s.IsRouting)
                {
                    dgv.Rows.RemoveAt(i);
                    continue;
                }

                int level = ParseInt(s.Row[COL_LEVEL]);
                if (level <= myLevel)
                    break;

                dgv.Rows.RemoveAt(i);
            }

            state.IsExpanded = false;

            string childKey = state.Row[COL_CHILD]?.ToString()?.Trim() ?? "";
            bool hasBomChildren = !string.IsNullOrEmpty(childKey) &&
                                  _childrenByParent.ContainsKey(childKey) &&
                                  _childrenByParent[childKey].Count > 0;
            bool hasRouting = !string.IsNullOrEmpty(childKey) &&
                              _routingByChild.ContainsKey(childKey) &&
                              _routingByChild[childKey].Count > 0;

            if (hasBomChildren || hasRouting)
                row.Cells[COL_EXPAND].Value = "+";
            else
                row.Cells[COL_EXPAND].Value = "";
        }
    }
}
