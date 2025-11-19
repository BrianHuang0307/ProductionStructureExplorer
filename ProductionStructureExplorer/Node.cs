using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionStructureExplorer
{
    public class Node
    {
        public string Parent { get; set; } = "";
        public string Child { get; set; } = "";

        // 組合項次
        public int Order { get; set; }

        // 在 DataTable 裡的列索引
        public int RowIndex { get; set; }
    }
}
