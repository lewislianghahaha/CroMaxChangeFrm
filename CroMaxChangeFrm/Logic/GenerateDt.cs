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
        /// <returns></returns>
        public DataTable Generatetemp(DataTable dt)
        {
            var resultdt=new DataTable();
            try
            {
                //获取表头临时表
                resultdt = dtList.Get_gendt();
                //循环从模板EXCEL获取的DT
                foreach (DataRow row in dt.Rows)
                {
                    var newrow = resultdt.NewRow();
                    newrow[0] = row[0];                                          //ID
                    newrow[1] =row[1];                                          //车厂
                    newrow[2] ="";                                              //颜色代码
                    newrow[3] ="";                                              //颜色名称
                    newrow[4] =row[2];                                          //适用车型
                    newrow[5] = "伊施威";                                       //品牌
                    newrow[6] =row[11];                                        //涂层
                    newrow[7] ="";                                             //差异色
                    newrow[8] =Convert.ToDateTime(row[10]).Year.ToString();    //年份(注:取‘制作日期’中的年份)
                    newrow[9] ="原车板";                                       //色版来源
                    newrow[10] =row[3];                                       //配方号
                    newrow[11] ="";                                           //颜色索引号
                    newrow[12] =row[10];                                     //制作日期
                    newrow[13] ="陈富明";                                    //制作人
                    newrow[14] =DateTime.Now.Date;                          //录入日期
                    newrow[15] = "冯惠娴";                                  //录入人
                    newrow[16] =row[9];                                    //审核日期
                    newrow[17] = "谭晓红";                                 //审核人
                    newrow[18] ="";                                       //备注
                    newrow[19] ="";                                      //来源分类
                    resultdt.Rows.Add(newrow);
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
        /// 运算-获取要生成的表体信息
        /// </summary>
        /// <param name="typeid">获取格式转换类型ID(0:格式转换 1:色母相关格式转换)</param>
        /// <param name="dt">从EXCEL导入的DT</param>
        /// <param name="tempdt">获取已运算成功的表头信息</param>
        /// <returns></returns>
        public DataTable GeneratetempEnpty(int typeid,DataTable dt,DataTable tempdt)
        {
            var resultdt=new DataTable();

            try
            {
                //获取对应临时表(表体)
                resultdt = dtList.Get_genenptydt();
                //循环获取已运算成功的表头信息
                foreach (DataRow row in tempdt.Rows)
                {
                    //根据表头的ID信息查询从EXCEL模板得出的DT内的相关内容
                    var rows = dt.Select("ID='" + Convert.ToInt32(row[0]) + "'");
                    //将相关值赋给resultdt临时表对应的项内
                    //if (rows.Length > 0)
                    //{
                        //循环执行获取11个色母量明细记录
                        for (var i = 0; i < 12; i++)
                        {
                            var newrow = resultdt.NewRow();
                            newrow = GenerColorantWeight(typeid, newrow, rows);
                            resultdt.Rows.Add(newrow);
                        }
                    //}
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
        /// <param name="typeid">获取格式转换类型ID(0:格式转换 1:色母相关格式转换)</param>
        /// <param name="newrows">新创建的行</param>
        /// <param name="rows">从EXCEL获取的行数组</param>
        /// <returns></returns>
        private DataRow GenerColorantWeight(int typeid,DataRow newrows,DataRow[] rows)
        {
            //累加量(克)
            decimal sumweight=0;
            //格式转换(只需计算累加量)
            if (typeid == 0)
            {
                newrows[0] = rows[0][3];
                newrows[1] = 
            }
            //色母相关格式转换(即需计算色母量及累加量)
            else
            {

            }
            return newrows;
        }
    }
}
