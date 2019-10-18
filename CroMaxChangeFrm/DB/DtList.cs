using System;
using System.Data;

namespace CroMaxChangeFrm.DB
{
    public class DtList
    {
        /// <summary>
        /// 从EXCEL导入的临时表（横向）
        /// </summary>
        /// <returns></returns>
        public DataTable Get_Importdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 35; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "主配方色号（差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "差异色名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 12:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                

                    case 13:
                        dc.ColumnName = "色母1";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母量1";
                        dc.DataType = Type.GetType("System.Decimal"); 
                        break;
                    case 15:
                        dc.ColumnName = "色母2";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "色母量2";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 17:
                        dc.ColumnName = "色母3";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 18:
                        dc.ColumnName = "色母量3";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 19:
                        dc.ColumnName = "色母4";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 20:
                        dc.ColumnName = "色母量4";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 21:
                        dc.ColumnName = "色母5";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 22:
                        dc.ColumnName = "色母量5";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 23:
                        dc.ColumnName = "色母6";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 24:
                        dc.ColumnName = "色母量6";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 25:
                        dc.ColumnName = "色母7";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 26:
                        dc.ColumnName = "色母量7";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 27:
                        dc.ColumnName = "色母8";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 28:
                        dc.ColumnName = "色母量8";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 29:
                        dc.ColumnName = "色母9";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 30:
                        dc.ColumnName = "色母量9";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 31:
                        dc.ColumnName = "色母10";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 32:
                        dc.ColumnName = "色母量10";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 33:
                        dc.ColumnName = "色母11";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 34:
                        dc.ColumnName = "色母量11";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导入纵向EXCEL模板
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ImportHdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 16; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方色号(差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //case 6:
                    //    dc.ColumnName = "差异色名称";
                    //    dc.DataType = Type.GetType("System.String");
                    //    break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;

                    case 11:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "色母量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 14:
                        dc.ColumnName = "累积量(可不填)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 15:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出EXCEL的临时表(最终导出使用)--旧系统使用
        /// </summary>
        /// <returns></returns>
        public DataTable Get_Exportdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 24; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "车厂";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "颜色代码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "颜色名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "适用车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "品牌";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "差异色";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "年份";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "色版来源";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "配方号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "颜色索引号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "制作日期";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    case 12:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "录入日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "录入人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 15:
                        dc.ColumnName = "审核日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "审核人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 17:
                        dc.ColumnName = "备注";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 18:
                        dc.ColumnName = "来源分类";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 19:
                        dc.ColumnName = "色母";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 20:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 21:
                        dc.ColumnName = "量(克)";
                        dc.DataType = Type.GetType("System.Double");  
                        break;
                    case 22:
                        dc.ColumnName = "累计量(克)";
                        dc.DataType = Type.GetType("System.Double"); 
                        break;
                    case 23:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出EXCEL的临时表(最终导出使用)--新系统使用
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportNewdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 17; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方色号(差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "差异色名称";
                        dc.DataType=Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;

                    case 12:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 15:
                        dc.ColumnName = "累积量(可不填)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 16:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出模板(以横向模式)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportVdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 33; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方色号（差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    //case 6:
                    //    dc.ColumnName = "差异色名称";
                    //    dc.DataType = Type.GetType("System.String");
                    //    break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;


                    case 11:
                        dc.ColumnName = "色母1";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "色母量1";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 13:
                        dc.ColumnName = "色母2";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母量2";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 15:
                        dc.ColumnName = "色母3";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "色母量3";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 17:
                        dc.ColumnName = "色母4";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 18:
                        dc.ColumnName = "色母量4";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 19:
                        dc.ColumnName = "色母5";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 20:
                        dc.ColumnName = "色母量5";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 21:
                        dc.ColumnName = "色母6";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 22:
                        dc.ColumnName = "色母量6";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 23:
                        dc.ColumnName = "色母7";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 24:
                        dc.ColumnName = "色母量7";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 25:
                        dc.ColumnName = "色母8";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 26:
                        dc.ColumnName = "色母量8";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 27:
                        dc.ColumnName = "色母9";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 28:
                        dc.ColumnName = "色母量9";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 29:
                        dc.ColumnName = "色母10";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 30:
                        dc.ColumnName = "色母量10";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 31:
                        dc.ColumnName = "色母11";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 32:
                        dc.ColumnName = "色母量11";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }


        /// <summary>
        /// 获取整理后的临时表-表头(旧系统使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_gendt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 21; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32"); 
                        break;
                    case 1:
                        dc.ColumnName = "车厂";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "颜色代码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "适用车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "品牌";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "差异色";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "年份";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "色版来源";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "配方号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "颜色索引号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "制作日期";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 13:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "录入日期";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 15:
                        dc.ColumnName = "录入人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "审核日期";
                        dc.DataType = Type.GetType("System.DateTime");
                        break;
                    case 17:
                        dc.ColumnName = "审核人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 18:
                        dc.ColumnName = "备注";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 19:
                        dc.ColumnName = "来源分类";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 20:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 获取整理后的临时表-表体(旧系统使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_genenptydt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 5; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "色母";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 4:
                        dc.ColumnName = "累计量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 获取整理后的临时表-表头(新系统使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_NewTempdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 13; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "主配方色号（差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "差异色名称";
                        dc.DataType = Type.GetType("System.String");
                        break;

                    case 8:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 获取整理后的临时表-表体(新系统使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_NewTempdtldt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 3; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "ID";
                        dc.DataType = Type.GetType("System.Int32");
                        break;
                    case 1:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "量(克)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        public DataTable Get_ErrorRecorddt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 1; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

    }
}
