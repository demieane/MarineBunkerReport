using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetKits.Extend.LinqToExcel
{
    /// <summary> 指示数据源中的对应映射列名
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Parameter, AllowMultiple = false)]
    public class MapFromAttribute : Attribute
    {
        public string ColumnName { get; set; }

        public MapFromAttribute(string columnName)
        {
            this.ColumnName = columnName;
        }
    }
}
