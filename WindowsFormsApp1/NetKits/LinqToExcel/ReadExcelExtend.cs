using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace NetKits.Extend.LinqToExcel
{
    public class ReadExcelExtend
    {
        /// <summary> 以异步方式 从excel中读取数据
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static async Task<ConvertExcelMdl<TMdl>> GetExcelDataAsync<TMdl>(string path)
        {
            return await Task.Factory.StartNew(() =>
            {
                return GetExcelData<TMdl>(path);
            });
        }

        /// <summary> 从excel中读取数据
        /// </summary>
        /// <typeparam name="TMdl"></typeparam>
        /// <param name="path"></param>
        /// <returns></returns>
        public static ConvertExcelMdl<TMdl> GetExcelData<TMdl>(string path)
        {
            ExcelQueryFactory userXLS = new ExcelQueryFactory(path);
            var errColumns = SetXlsMapping<TMdl>(userXLS);
            if (errColumns.Count > 0)
                return new ConvertExcelMdl<TMdl>(errColumns, null);

            var userQuery = userXLS.Worksheet<TMdl>(0).ToList();
            if (userQuery.Count == 0)
                userQuery = null;
            return new ConvertExcelMdl<TMdl>(errColumns, userQuery);
        }

        /// <summary> 设置xls 映射关系
        /// </summary>
        /// <param name="userXLS"></param>
        /// <returns> 错误列 集合
        /// </returns>
        private static List<string> SetXlsMapping<VMdl>(ExcelQueryFactory userXLS)
        {
            List<string> errColumns = new List<string>();

            //xls现有数据字段
            var firstSheet = userXLS.GetWorksheetNames().First<string>();
            List<string> cnlist = userXLS.GetColumnNames(firstSheet).ToList<string>();

            Type type = typeof(VMdl);
            var ProArr = type.GetProperties();
            foreach (var pro in ProArr)
            {
                var attributeArr = pro.GetCustomAttributes(true);

                var title = pro.Name;
                var name = pro.Name;

                var mapAttribute = attributeArr.OfType<MapFromAttribute>().FirstOrDefault();
                if (mapAttribute != null)
                    title = mapAttribute.ColumnName;

                var requiredAttri = attributeArr.OfType<NecessaryColumnAttribute>().FirstOrDefault();
                if (requiredAttri != null)
                {
                    if (!cnlist.Exists(cn => cn == title)) //不存在
                    {
                        errColumns.Add(title);
                        continue;
                    }
                }

                if (mapAttribute != null)
                    userXLS.AddMapping(name, title);
            }
            return errColumns;
        }
    }
}
