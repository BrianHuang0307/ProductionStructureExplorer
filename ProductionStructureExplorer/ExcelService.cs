using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ProductionStructureExplorer
{
    public class ExcelService : IExcelService
    {
        // 共用欄位名稱常數
        private const string COL_PARENT = "PARENT";
        private const string COL_CHILD = "CHILD";
        private const string COL_ORDER = "組合項次";   // Excel 欄名還是這個
        private const string COL_FULLPATH = "FULL_PATH";
        private const string COL_EXPIRE = "失效日期";
        private const string COL_FEATURE = "特性編碼";
        private const string COL_LEVEL = "階層";
        private const string COL_ROOT = "ROOT";
        private const string COL_USED = "是否使用";

        private const string COL_EXPAND = "_EXPAND_";   // Grid 第一欄用

        public DataTable LoadBom(string path, out string featureCode, out string rootCode)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);  // 第一個工作表

            var dt = new DataTable();
            bool firstRow = true;

            foreach (var row in ws.RowsUsed())
            {
                if (firstRow)
                {
                    foreach (var cell in row.CellsUsed())
                        dt.Columns.Add(cell.GetString().Trim());
                    firstRow = false;
                }
                else
                {
                    var newRow = dt.NewRow();
                    int colIndex = 0;
                    foreach (var cell in row.Cells(1, dt.Columns.Count))
                    {
                        newRow[colIndex++] = cell.Value;
                    }
                    dt.Rows.Add(newRow);
                }
            }

            if (!dt.Columns.Contains(COL_PARENT) ||
                !dt.Columns.Contains(COL_CHILD) ||
                !dt.Columns.Contains(COL_ORDER) ||
                !dt.Columns.Contains(COL_LEVEL))
                throw new Exception("BOM 檔缺少必要欄位：階層 / PARENT / CHILD / 組合項次。");

            if (!dt.Columns.Contains(COL_FULLPATH))
                dt.Columns.Add(COL_FULLPATH, typeof(string));

            bool hasExpire = dt.Columns.Contains(COL_EXPIRE);
            bool hasFeature = dt.Columns.Contains(COL_FEATURE);
            bool hasRoot = dt.Columns.Contains(COL_ROOT);

            // 抓特性編碼
            featureCode = "";
            if (hasFeature)
            {
                foreach (DataRow r in dt.Rows)
                {
                    var v = r[COL_FEATURE]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v))
                    {
                        featureCode = v;
                        break;
                    }
                }
            }

            // 抓 ROOT（第一個非空的 ROOT）
            rootCode = "";
            if (hasRoot)
            {
                foreach (DataRow r in dt.Rows)
                {
                    var v = r[COL_ROOT]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v))
                    {
                        rootCode = v;
                        break;
                    }
                }
            }

            // 失效日期有值的列全部刪掉
            if (hasExpire)
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    var v = dt.Rows[i][COL_EXPIRE];
                    if (v != null && v != DBNull.Value &&
                        !string.IsNullOrWhiteSpace(v.ToString()))
                    {
                        dt.Rows.RemoveAt(i);
                    }
                }
            }

            dt.AcceptChanges();
            return dt;
        }

        public DataTable LoadRouting(string path)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);  // 第一個工作表

            var dt = new DataTable();
            bool firstRow = true;

            foreach (var row in ws.RowsUsed())
            {
                if (firstRow)
                {
                    foreach (var cell in row.CellsUsed())
                        dt.Columns.Add(cell.GetString().Trim());
                    firstRow = false;
                }
                else
                {
                    var newRow = dt.NewRow();
                    int colIndex = 0;
                    foreach (var cell in row.Cells(1, dt.Columns.Count))
                    {
                        newRow[colIndex++] = cell.Value;
                    }
                    dt.Rows.Add(newRow);
                }
            }

            if (!dt.Columns.Contains(COL_CHILD) ||
                !dt.Columns.Contains("工序"))
                throw new Exception("Routing 檔缺少必要欄位：CHILD / 工序。");

            dt.AcceptChanges();
            return dt;
        }

        // 直接匯出「當下 DataGridView 的內容」
        public void ExportGridToExcel(DataGridView grid, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("GRID_EXPORT");

            // 1. 欄位：跳過 _EXPAND_ 欄
            int col = 1;
            var exportColumns = grid.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible && !string.Equals(c.Name, COL_EXPAND, StringComparison.OrdinalIgnoreCase))
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            foreach (var c in exportColumns)
            {
                ws.Cell(1, col).Value = c.HeaderText;
                col++;
            }

            // 2. 資料列
            int rowIndex = 2;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;

                col = 1;
                foreach (var c in exportColumns)
                {
                    var cell = row.Cells[c.Name];
                    var v = cell.Value;
                    ws.Cell(rowIndex, col).Value = v == null ? "" : v.ToString();
                    col++;
                }

                rowIndex++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }
    }
}
