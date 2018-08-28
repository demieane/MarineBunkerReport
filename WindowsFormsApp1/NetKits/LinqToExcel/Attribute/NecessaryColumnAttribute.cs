using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetKits.Extend.LinqToExcel
{
    /// <summary> 指示标记的列为必需项
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Parameter, AllowMultiple = false)]
    public class NecessaryColumnAttribute : Attribute
    {
    }
}
