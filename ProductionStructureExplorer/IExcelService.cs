using System.Data;
using System.Windows.Forms;

namespace ProductionStructureExplorer
{
    public interface IExcelService
    {
        /// 載入 BOM Excel，回傳 DataTable，並帶出特性編碼與 ROOT。
        DataTable LoadBom(string path, out string featureCode, out string rootCode);

        /// 載入 Routing Excel，回傳 DataTable。
        DataTable LoadRouting(string path);

        /// 直接把當前 DataGridView 的畫面匯出到 Excel。
        void ExportGridToExcel(DataGridView grid, string path);
    }
}
