using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ProductionStructureExplorer
{
    public static class BomService
    {
        private const string COL_PARENT = "PARENT";
        private const string COL_CHILD = "CHILD";
        private const string COL_ORDER = "組合項次";

        public static List<int> BuildDfsOrder(DataTable dt)
        {
            var nodes = new List<Node>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                string parent = row[COL_PARENT]?.ToString()?.Trim() ?? "";
                string child = row[COL_CHILD]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(child))
                    continue;

                int order = int.MaxValue;
                var orderObj = row[COL_ORDER];
                if (orderObj != null && orderObj != DBNull.Value &&
                    int.TryParse(orderObj.ToString().Trim(), out var tmp))
                    order = tmp;

                nodes.Add(new Node
                {
                    Parent = parent,
                    Child = child,
                    Order = order,
                    RowIndex = i
                });
            }

            var map = nodes
                .GroupBy(e => e.Parent)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(x => x.Order).ThenBy(x => x.Child).ToList()
                );

            var parents = new HashSet<string>(nodes.Select(e => e.Parent));
            var childs = new HashSet<string>(nodes.Select(e => e.Child));
            var roots = parents.Except(childs).ToList();

            if (!roots.Any())
                throw new Exception("找不到 Root（成品），請確認 BOM 資料。");

            var orderList = new List<int>();
            var visited = new HashSet<string>();

            foreach (var root in roots.OrderBy(r => r))
            {
                DfsNode(root, map, visited, orderList);
            }

            return orderList;
        }

        private static void DfsNode(
            string parent,
            Dictionary<string, List<Node>> map,
            HashSet<string> visitedInPath,
            List<int> order)
        {
            if (visitedInPath.Contains(parent))
                return;

            visitedInPath.Add(parent);

            if (map.TryGetValue(parent, out var children))
            {
                foreach (var edge in children)
                {
                    order.Add(edge.RowIndex);
                    DfsNode(edge.Child, map, visitedInPath, order);
                }
            }

            visitedInPath.Remove(parent);
        }

        public static DataTable ApplyOrder(DataTable dt, List<int> order)
        {
            var result = dt.Clone();

            foreach (int idx in order)
            {
                if (idx >= 0 && idx < dt.Rows.Count)
                    result.ImportRow(dt.Rows[idx]);
            }

            return result;
        }

        public static Dictionary<string, List<DataRow>> BuildChildrenMap(DataTable bomSorted)
        {
            var dict = new Dictionary<string, List<DataRow>>(StringComparer.OrdinalIgnoreCase);

            foreach (DataRow row in bomSorted.Rows)
            {
                string parent = row[COL_PARENT]?.ToString()?.Trim() ?? "";
                string child = row[COL_CHILD]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(parent) || string.IsNullOrEmpty(child))
                    continue;

                if (!dict.TryGetValue(parent, out var list))
                {
                    list = new List<DataRow>();
                    dict[parent] = list;
                }

                list.Add(row);
            }

            return dict;
        }
    }
}
