using System;
using System.Data;
using CroMaxChangeFrm.DB;

namespace CroMaxChangeFrm.Logic
{
    //运算
    public class GenerateDt
    {
        DtList dtList=new DtList();

        /// <summary>
        /// 运算-通过从EXCEL导入的DT获取表头信息
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="comselectid"></param>
        /// <returns></returns>
        public DataTable Generatetemp(DataTable dt, int comselectid)
        {
            var resultdt=new DataTable();

            //保存‘配方代码’字段,用于排除重复值
            var colorcode = string.Empty;
            //保存‘版本日期’字段,用于排除重复值
            var confrimdt = string.Empty;
            //保存‘层’字段,用于排除重复值
            var layer = 0;

            try
            {
                #region 旧系统使用

                if (comselectid == 1)
                {
                    //获取表头临时表
                    resultdt = dtList.Get_gendt();
                    //循环从模板EXCEL获取的DT
                    foreach (DataRow row in dt.Rows)
                    {
                        var newrow = resultdt.NewRow();
                        newrow[0] = row[0];                                          //ID
                        newrow[1] = row[1];                                           //车厂
                        newrow[2] = "";                                               //颜色代码
                        newrow[3] = "";                                               //颜色名称
                        newrow[4] = row[2];                                           //适用车型
                        newrow[5] = "伊施威";                                         //品牌
                        newrow[6] = row[11];                                          //涂层
                        newrow[7] = "";                                               //差异色
                        newrow[8] = Convert.ToDateTime(row[10]).Year.ToString();      //年份(注:取‘制作日期’中的年份)
                        newrow[9] = "原车板";                                         //色版来源
                        newrow[10] = row[3];                                         //配方号
                        newrow[11] = "";                                             //颜色索引号
                        newrow[12] = row[10];                                        //制作日期
                        newrow[13] = "陈富明";                                        //制作人
                        newrow[14] = DateTime.Now.Date;                              //录入日期
                        newrow[15] = "冯惠娴";                                       //录入人
                        newrow[16] = row[9];                                          //审核日期
                        newrow[17] = "谭晓红";                                       //审核人
                        newrow[18] = "";                                              //备注
                        newrow[19] = "";                                              //来源分类
                        newrow[20] = row[7];                                         //颜色组别(导出新数据库模板使用)
                        resultdt.Rows.Add(newrow);
                        
                    }
                }
                #endregion
                //以新系统模板纵向导出
                else if(comselectid==2)
                {
                    //获取表头临时表
                    resultdt = dtList.Get_NewTempdt();
                    //循环从模板EXCEL获取的DT
                    foreach (DataRow rows in dt.Rows)
                    {
                        var newrow = resultdt.NewRow();
                        newrow[0] = rows[0];        //ID
                        newrow[1] = rows[1];        //制造商
                        newrow[2] = rows[2];        //车型
                        newrow[3] = rows[3];        //涂层
                        newrow[4] = rows[4];        //颜色描述
                        newrow[5] = rows[5];        //内部色号
                        newrow[6] = rows[6];        //主配方色号（差异色)
                        newrow[7] = rows[7];        //差异色名称
                        newrow[8] = rows[8];        //颜色组别
                        newrow[9] = rows[9];        //标准色号
                        newrow[10] = rows[10];      //RGBValue
                        newrow[11] = rows[11];      //版本日期
                        newrow[12] = rows[12];      //层
                        resultdt.Rows.Add(newrow);
                    }
                }
                //以新系统模板横向导出
                else if (comselectid == 3)
                {
                    //获取导出模板(横向)
                    resultdt = dtList.Get_ExportVdt(); //dtList.Get_ErrorRecorddt();

                    //先循环导入EXCEL数据源
                    foreach (DataRow rows in dt.Rows)
                    {
                        //若循环获取的‘内部色号’与变量一致,即不用继续
                        if (colorcode == Convert.ToString(rows[4]) && confrimdt == Convert.ToString(rows[9]) && layer == Convert.ToInt32(rows[10])) continue;
                        //若不相同,先将当前循环行的值进行赋值至变量
                        colorcode = Convert.ToString(rows[4]);
                        confrimdt = Convert.ToString(rows[9]);
                        layer = Convert.ToInt32(rows[10]);

                        resultdt.Merge(GetVdt(rows,dt,resultdt));
                    }
                }
            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }
            
            return resultdt;
        }

        /// <summary>
        /// 以横向方式导出使用
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="sourcedt"></param>
        /// <param name="resultdt"></param>
        /// <returns></returns>
        private DataTable GetVdt(DataRow rows,DataTable sourcedt,DataTable resultdt)
        {
            //先将‘制造商’等表头相关信息插入,再插入色母等信息
            var newrow = resultdt.NewRow();
            newrow[0] = rows[0];//制造商
            newrow[1] = rows[1];//车型
            newrow[2] = rows[2];//涂层
            newrow[3] = rows[3];//颜色描述
            newrow[4] = rows[4];//内部色号
            newrow[5] = rows[5];//主配方色号（差异色)
            newrow[6] = rows[6];//颜色组别
            newrow[7] = rows[7];//标准色号
            newrow[8] = rows[8];//RGBValue
            newrow[9] = rows[9];//版本日期
            newrow[10] = rows[10];//层

            //将‘色母’相关信息，插入至对应的项内
            var rowsdtl = sourcedt.Select("内部色号='"+Convert.ToString(rows[4])+"' and 版本日期='"+ Convert.ToString(rows[9])+"' and 层='"+Convert.ToInt32(rows[10])+"'");
            //if (rowsdtl.Length > 11)
            //{
            //    newrow[0] = rows[4];

            //    //var a1 = Convert.ToString(rows[4]);
            //    //var a = Convert.ToString(rows[9]);
            //}
            for (var i = 0; i < rowsdtl.Length; i++)
            {
                newrow[11 + i + i] = rowsdtl[i][11];        //色母编码
                newrow[11 + i + i + 1] = rowsdtl[i][13];    //色母量(保留两位小数)
            }
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 运算-获取要生成的表体信息
        /// </summary>
        /// <param name="typeid">获取格式转换类型ID(0:格式转换 1:色母相关格式转换)</param>
        /// <param name="dt">从EXCEL导入的DT</param>
        /// <param name="tempdt">获取已运算成功的表头信息</param>
        /// <param name="comselectid"></param>
        /// <returns></returns>
        public DataTable GeneratetempEnpty(int typeid,DataTable dt,DataTable tempdt,int comselectid)
        {
            var resultdt = new DataTable();

            try
            {
                if (comselectid == 1)
                {
                    //获取对应临时表(表体)
                    resultdt = dtList.Get_genenptydt();
                    //循环获取已运算成功的表头信息
                    foreach (DataRow row in tempdt.Rows)
                    {
                        //根据表头的ID信息查询从EXCEL模板得出的DT内的相关内容
                        var rows = dt.Select("ID='" + Convert.ToInt32(row[0]) + "'");
                        //执行插入相关信息至临时表
                        resultdt.Merge(GenerColorantWeight(comselectid,typeid, resultdt, rows));
                    }
                }
                else
                {
                    //获取对应临时表(表体)
                    resultdt = dtList.Get_NewTempdtldt();
                    //循环获取已运算成功的表头信息
                    foreach (DataRow row in tempdt.Rows)
                    {
                        //根据表头的ID信息查询从EXCEL模板得出的DT内的相关内容
                        var rows = dt.Select("ID='" + Convert.ToInt32(row[0]) + "'");
                        //执行插入相关信息至临时表
                        resultdt.Merge(GenerColorantWeight(comselectid,typeid, resultdt, rows));
                    }
                }

            }
            catch (Exception)
            {
                resultdt.Rows.Clear();
                resultdt.Columns.Clear();
            }
            return resultdt;
        }

        /// <summary>
        /// 根据状态标记-整理色母量明细
        /// </summary>
        /// <param name="comselectid"></param>
        /// <param name="typeid">获取格式转换类型ID(0:格式转换 1:色母相关格式转换)</param>
        /// <param name="sourcedt">临时表</param>
        /// <param name="rows">从EXCEL获取的行数组</param>
        /// <returns></returns>
        private DataTable GenerColorantWeight(int comselectid,int typeid,DataTable sourcedt,DataRow[] rows)
        {
            //累加量(克)
            decimal sumweight = 0;

            if (comselectid == 1)
            {
                //循环执行获取11个色母量明细记录
                for (var i = 1; i < 12; i++)
                {
                    //格式转换(只需计算累加量)
                    if (typeid == 0)
                    {
                        //先根据循环ID获取对应的列色母名称
                        var colorantname = Convert.ToString(rows[0][13 + i + i]);
                        //判断若获取的色母为空,就不作添加
                        if (colorantname == "") continue;
                        var newrows = sourcedt.NewRow();
                        newrows[0] = rows[0][0];                                                  //ID
                        newrows[1] = rows[0][13 + i + i];                                        //色母
                        newrows[2] = "";                                                        //色母名称
                        newrows[3] = rows[0][13 + i + i + 1];                                  //量(克)
                        newrows[4] = sumweight += Convert.ToDecimal(rows[0][13 + i + i + 1]); //累计量(克)
                        sourcedt.Rows.Add(newrows);
                    }
                    //色母相关格式转换(即需计算色母量及累加量)
                    else
                    {
                    
                    }
                }
            }
            else
            {
                //循环执行获取11个色母量明细记录
                for (var i = 1; i < 12; i++)
                {
                    //先根据循环ID获取对应的列色母编码
                    var colorantname = Convert.ToString(rows[0][11 + i + i]);
                    //判断若获取的色母编码为空,就不作添加
                    if (colorantname == "") continue;
                    var newrows = sourcedt.NewRow();
                    newrows[0] = rows[0][0];                //ID
                    newrows[1] = colorantname;              //色母编码
                    newrows[2] = rows[0][11 + i + i + 1];   //量(克)
                    sourcedt.Rows.Add(newrows);
                }
            }
            return sourcedt;
        }
    }
}
