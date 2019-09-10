using System;
using System.Data;
using System.IO;
using CroMaxChangeFrm.DB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CroMaxChangeFrm.Logic
{
    //导出
    public class ExportDt
    {
        DtList dtList=new DtList();

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress">导出地址</param>
        /// <param name="tempdt">运算结果-表头</param>
        /// <param name="tempdtldt">运算结果-表体</param>
        /// <param name="comselectid">获取下拉框ID(注:通过此判断以旧系统模板 还是 新系统模板 导出) 1:导出至旧数据库 2:导出至新数据库</param>
        public bool ExportDtToExcel(string fileAddress, DataTable tempdt, DataTable tempdtldt,int comselectid)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;

            //定义运算DT
            var temp=new DataTable();

            try
            {
                //声明一个WorkBook
                var xssfWorkbook = new XSSFWorkbook();
                //通过运算得出的表头及表体合并最终DT
                //1:旧系统使用 //2:新系统使用
                temp = comselectid == 1 ? Margedt(tempdt, tempdtldt) : UseDtChangeNewdt(Margedt(tempdt, tempdtldt));

                //执行sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充; 以一个sheet表有9W行记录填充为基准)
                sheetcount = temp.Rows.Count % 100000 == 0 ? temp.Rows.Count / 100000 : temp.Rows.Count / 100000 + 1;
                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("Sheet" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);
                    //创建sheet页各列标题
                    for (var j = 0; j < temp.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        //旧系统使用
                        if (comselectid == 1)
                        {
                            switch (j)
                            {
                                #region SetCellValue 旧系统模板使用
                                case 0:
                                    row.CreateCell(j).SetCellValue("车厂");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("颜色代码");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("颜色名称");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("适用车型");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("品牌");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("差异色");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("年份");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("色版来源");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("配方号");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("颜色索引号");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("制作日期");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                case 13:
                                    row.CreateCell(j).SetCellValue("录入日期");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("录入人");
                                    break;
                                case 15:
                                    row.CreateCell(j).SetCellValue("审核日期");
                                    break;
                                case 16:
                                    row.CreateCell(j).SetCellValue("审核人");
                                    break;
                                case 17:
                                    row.CreateCell(j).SetCellValue("备注");
                                    break;
                                case 18:
                                    row.CreateCell(j).SetCellValue("来源分类");
                                    break;
                                case 19:
                                    row.CreateCell(j).SetCellValue("色母");
                                    break;
                                case 20:
                                    row.CreateCell(j).SetCellValue("色母名称");
                                    break;
                                case 21:
                                    row.CreateCell(j).SetCellValue("量(克)");
                                    break;
                                case 22:
                                    row.CreateCell(j).SetCellValue("累计量(克)");
                                    break;
                                    #endregion
                            }
                        }
                        //新系统使用
                        else
                        {
                            switch (j)
                            {
                                #region SetCellValue 新系统模板使用
                                case 0:
                                    row.CreateCell(j).SetCellValue("制造商");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("车型");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("颜色描述");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("主配方色号(差异色)");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("颜色组别");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("标准色号");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("RGBValue");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("层");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("色母编码");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("色母名称");
                                    break;
                                case 13:
                                    row.CreateCell(j).SetCellValue("色母量");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("累积量(可不填)");
                                    break;
                                case 15:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                    #endregion
                            }
                        }

                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 100000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? temp.Rows.Count : i * 100000;

                    //每一个sheet表显示90000行  
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < temp.Columns.Count; k++)
                        {
                            if(Convert.ToString(temp.Rows[j][k]) == "") continue;
                            else
                            {
                                //当ColNum=21 或 22时,执行(注:要注意值小数位数保留两位;当超出三位小数的时候,会出现OutofMemory异常.)
                                //注:当需要转出模板为旧系统时 使用 列ID为21 22;当为新系统时 使用 列ID为13 14
                                if (comselectid == 1)
                                {
                                    if (k == 21 || k == 22)
                                    {
                                        row.CreateCell(k, CellType.Numeric).SetCellValue(Convert.ToDouble(temp.Rows[j][k]));
                                    }
                                }
                                else
                                {
                                    if (k == 13 || k == 14)
                                    {
                                        row.CreateCell(k, CellType.Numeric).SetCellValue(Convert.ToDouble(temp.Rows[j][k]));
                                    }
                                }
                                //除‘色母量’以及‘累积量’外的值的转换赋值
                                row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(temp.Rows[j][k]));
                            }
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }

                //写入数据
                var file = new FileStream(fileAddress, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();           //关闭文件流
                xssfWorkbook.Close();   //关闭工作簿
                file.Dispose();         //释放文件流
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        /// <summary>
        /// 合并表格-将运算过来的临时表头及表体进行合并(作导出数据之用)
        /// </summary>
        /// <param name="tempdt">运算结果-表头</param>
        /// <param name="tempemptydt">运算结果-表体</param>
        /// <returns></returns>
        private DataTable Margedt(DataTable tempdt,DataTable tempemptydt)
        {
            //获取导出EXCEL临时表
            var resultdt = dtList.Get_Exportdt();
            //循环表头信息
            foreach (DataRow rows in tempdt.Rows)
            {
                //根据ID值获取表头对应的表体信息
                var emptyrow = tempemptydt.Select("ID='" + Convert.ToInt32(rows[0]) + "'");
                //循环表体信息(注:若i=0即表示插入第一行记录,也就可以将表头信息插入至临时表对应的列;而除i=0外,其余的表头信息对应的列都为空)
                for (var i = 0; i < emptyrow.Length; i++)
                {
                    var newrows = resultdt.NewRow();
                    newrows[0] = i == 0 ? rows[1] : DBNull.Value;              //车厂
                    newrows[1] = i == 0 ? rows[2] : DBNull.Value;              //颜色代码
                    newrows[2] = i == 0 ? rows[3] : DBNull.Value;              //颜色名称
                    newrows[3] = i == 0 ? rows[4] : DBNull.Value;              //适用车型
                    newrows[4] = i == 0 ? rows[5] : DBNull.Value;              //品牌
                    newrows[5] = i == 0 ? rows[6] : DBNull.Value;              //涂层
                    newrows[6] = i == 0 ? rows[7] : DBNull.Value;              //差异色
                    newrows[7] = i == 0 ? rows[8] : DBNull.Value;              //年份
                    newrows[8] = i == 0 ? rows[9] : DBNull.Value;              //色版来源
                    newrows[9] = i == 0 ? rows[10] : DBNull.Value;             //配方号
                    newrows[10] = i == 0 ? rows[11] : DBNull.Value;            //颜色索引号
                    newrows[11] = i == 0 ? rows[12] : DBNull.Value;            //制作日期
                    newrows[12] = i == 0 ? rows[13] : DBNull.Value;            //制作人
                    newrows[13] = i == 0 ? rows[14] : DBNull.Value;            //录入日期
                    newrows[14] = i == 0 ? rows[15] : DBNull.Value;            //录入人
                    newrows[15] = i == 0 ? rows[16] : DBNull.Value;            //审核日期
                    newrows[16] = i == 0 ? rows[17] : DBNull.Value;            //审核人
                    newrows[17] = i == 0 ? rows[18] : DBNull.Value;            //备注
                    newrows[18] = i == 0 ? rows[19] : DBNull.Value;            //来源分类
                    newrows[19] = emptyrow[i][1];                              //色母
                    newrows[20] = emptyrow[i][2];                              //色母名称
                    newrows[21] = Math.Round(Convert.ToDouble(emptyrow[i][3]),2);  //量(克)
                    newrows[22] = Math.Round(Convert.ToDouble(emptyrow[i][4]),2);  //累计量(克)
                    newrows[23] = i == 0 ? rows[20] : DBNull.Value;                //颜色组别
                    resultdt.Rows.Add(newrows);
                }
            }
            return resultdt;
        }

        /// <summary>
        /// 将旧数据库模板DT 转换至 新数据库模板DT
        /// </summary>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        private DataTable UseDtChangeNewdt(DataTable sourcedt)
        {
            //获取导出EXCEL临时表
            var resultdt = dtList.Get_ExportNewdt();
            //循环使用sourcedt读取记录;目的:将旧数据库模板DT 转换 至 新数据库模板DT
            foreach (DataRow rows in sourcedt.Rows)
            {
                var newrows = resultdt.NewRow();
                newrows[0] = rows[0];       //制造商
                newrows[1] = rows[3];       //车型
                newrows[2] = rows[5];       //涂层
                newrows[3] = rows[2];       //颜色描述
                newrows[4] = "";            //内部色号
                newrows[5] = "";            //主配方色号(差异色)
                newrows[6] = rows[23];      //颜色组别
                newrows[7] = rows[1];       //标准色号
                newrows[8] = "";            //RBGValue
                newrows[9] = rows[15];      //版本日期
                newrows[10] = "";           //层
                newrows[11] = rows[20];     //色母编码
                newrows[12] = "";           //色母名称
                newrows[13] = rows[21];     //色母量(克)
                newrows[14] = rows[22];     //累积量
                newrows[15] = rows[12];     //制作人
                resultdt.Rows.Add(newrows);
            }
            return resultdt;
        }

    }
}
