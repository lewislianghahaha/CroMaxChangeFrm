using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using CroMaxChangeFrm.DB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CroMaxChangeFrm.Logic
{
    public class ImportDt
    {
        DtList dtList = new DtList();

        /// <summary>
        /// 打开及导入至DT
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="selectid"></param>
        /// <returns></returns>
        public DataTable OpenExcelImporttoDt(string fileAddress,int selectid)
        {
            var dt=new DataTable();
            try
            {
                //使用NPOI技术进行导入EXCEL至DATATABLE
                var importExcelDt = OpenExcelToDataTable(fileAddress,selectid);
                //将从EXCEL过来的记录集为空的行清除
                dt = RemoveEmptyRows(importExcelDt);
            }
            catch (Exception)
            {
                dt.Rows.Clear();
                dt.Columns.Clear();
            }
             return dt;
        }

        private DataTable OpenExcelToDataTable(string fileAddress,int selectid)
        {
            IWorkbook wk;
            //定义列ID
            var colid = 0;
            //定义ID变量
            var id = 1;
            //定义内部色号
            var code = string.Empty;
            //定义版本日期
            var confirmdt = string.Empty;
            //定义层
            var layer = 0;
            //创建表标题
            var dt=new DataTable();
            dt = selectid == 3 ? dtList.Get_ImportHdt() : dtList.Get_Importdt();

            using (var fsRead = File.OpenRead(fileAddress))
            {
                wk = new XSSFWorkbook(fsRead);
                //获取第一个sheet
                var sheet = wk.GetSheetAt(0);
                //获取第一行
                //var hearRow = sheet.GetRow(0);

                //创建完标题后,开始从第二行起读取对应列的值
                for (var r = 1; r <= sheet.LastRowNum; r++)
                {
                    var result = false;
                    var dr = dt.NewRow();
                    //获取当前行(注:只能获取行中有值的项,为空的项不能获取;即row.Cells.Count得出的总列数就只会汇总"有值的列"之和)
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    //读取每列(固定了共列值37)  (固定列为34列=>新导入模板使用 add date:20191009)
                    colid = selectid == 3 ? 17 : 35;

                    for (var j = 0; j < colid/*37*//*row.Cells.Count*/; j++)
                    {
                        if (j == 0 && selectid !=3)
                        {
                            dr[0] = id;
                        }

                        else
                        {
                            //循环获取行中的单元格
                            var cell = row.GetCell(j);
                            var cellValue = GetCellValue(cell);
                            if (cellValue == string.Empty)
                            {
                                if (j == 4 && selectid == 3)
                                {
                                    dr[j] = code;
                                }
                                else if (j == 9 && selectid == 3)
                                {
                                    dr[j] = confirmdt;
                                }
                                else if (j == 10 && selectid == 3)
                                {
                                    dr[j] = layer;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                dr[j] = cellValue;
                                if (j == 4 && selectid == 3)
                                {
                                    code = Convert.ToString(dr[j]);
                                }
                                else if (j == 9 && selectid == 3)
                                {
                                    confirmdt = Convert.ToString(dr[j]);
                                }
                                else if (j == 10 && selectid == 3)
                                {
                                    layer = Convert.ToInt32(dr[j]);
                                }
                            }

                            //全为空就不取
                            if (dr[j].ToString() != "")
                            {
                                result = true;
                            }
                        }
                    }

                    if (result == true)
                    {
                        //把每行增加到DataTable
                        dt.Rows.Add(dr);
                    }
                    //自增ID值
                    if(selectid!=3)
                        id++;
                }
            }
            return dt;
        }

        /// <summary>
        /// 检查单元格的数据类型并获其中的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (DateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();

                    }

                case CellType.Unknown: //无法识别类型
                default: //默认类型                    
                    return cell.ToString();
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        var e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }

            }
        }

        /// <summary>
        ///  将从EXCEL导入的DATATABLE的空白行清空
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            var removeList = new List<DataRow>();
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var isNull = true;
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    //将不为空的行标记为False
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        isNull = false;
                    }
                }
                //将整行都为空白的记录进行记录
                if (isNull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }

            //将整理出来的所有空白行通过循环进行删除
            for (var i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
            return dt;
        }
    }
}
