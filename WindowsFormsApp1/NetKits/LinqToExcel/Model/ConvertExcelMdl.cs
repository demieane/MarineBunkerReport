using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetKits.Extend.LinqToExcel
{
    /// <summary> 转换 读取Excel数据结果
    /// </summary>
    /// <typeparam name="TResultMdl"></typeparam>
    public class ConvertExcelMdl<TResultMdl>
    {
        /// <summary> 列信息有错误
        /// </summary>
        public bool HasErr
        {
            get { return ErrColumns != null && ErrColumns.Count > 0; }
        }

        /// <summary> 缺失的列集合
        /// </summary>
        public List<string> ErrColumns { get; set; }

        /// <summary> 列无问题后 返回的读取数据
        /// </summary>
        public List<TResultMdl> ResultData { get; set; }

        public ConvertExcelMdl(List<string> errColumns, List<TResultMdl> resultData)
        {
            this.ErrColumns = errColumns;
            this.ResultData = resultData;
        }
    }
}
