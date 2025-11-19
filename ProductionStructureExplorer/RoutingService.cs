using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ProductionStructureExplorer
{
    public static class RoutingService
    {
        private const string COL_CHILD = "CHILD";
        private const string COL_OPSEQ = "工序";

        public static Dictionary<string, List<DataRow>> BuildRoutingMap(DataTable routing)
        {
            return routing.AsEnumerable()
                .Where(r => !IsNullOrWhiteSpace(r, COL_CHILD))
                .GroupBy(r => r[COL_CHILD].ToString().Trim())
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderBy(r => ParseOpSeq(r[COL_OPSEQ]))
                          .ThenBy(r => r.Table.Rows.IndexOf(r))
                          .ToList(),
                    StringComparer.OrdinalIgnoreCase
                );
        }

        private static bool IsNullOrWhiteSpace(DataRow row, string colName)
        {
            if (!row.Table.Columns.Contains(colName)) return true;
            var v = row[colName];
            if (v == null || v == DBNull.Value) return true;
            return string.IsNullOrWhiteSpace(v.ToString());
        }

        private static int ParseOpSeq(object? v)
        {
            if (v == null || v == DBNull.Value) return int.MaxValue;
            var s = v.ToString()?.Trim();
            return int.TryParse(s, out var n) ? n : int.MaxValue;
        }
    }
}
