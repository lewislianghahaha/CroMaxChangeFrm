using System.Data;
using System.Threading;

namespace CroMaxChangeFrm.Logic
{
    public class TaskLogic
    {
        ImportDt importDt=new ImportDt();
        GenerateDt generateDt=new GenerateDt();
        ExportDt exportDt=new ExportDt();

        private int _taskid;
        private string _fileAddress;       //文件地址
        private DataTable _dt;             //获取dt(从EXCEL获取的DT)
        private int _typeid;               //获取格式转换类型ID(0:格式转换 1:色母相关格式转换)

        private DataTable _resultTable;   //返回DT
        private bool _resultMark;        //返回是否成功标记

        #region Set
            /// <summary>
            /// 中转ID
            /// </summary>
            public int TaskId { set { _taskid = value; } }

            /// <summary>
            /// //接收文件地址信息
            /// </summary>
            public string FileAddress { set { _fileAddress = value; } }

            /// <summary>
            /// 获取dt(从EXCEL获取的DT)
            /// </summary>
            public DataTable Data { set { _dt = value; } }

            /// <summary>
            /// 获取格式转换类型ID(0:格式转换 1:色母相关格式转换)
            /// </summary>
            public int Typeid { set { _typeid = value; } }
        #endregion

        #region Get
            /// <summary>
            ///返回DataTable至主窗体
            /// </summary>
            public DataTable RestulTable => _resultTable;

            /// <summary>
            ///  返回是否成功标记
            /// </summary>
            public bool ResultMark => _resultMark;
        #endregion

        public void StartTask()
        {
            Thread.Sleep(1000);

            switch (_taskid)
            {
                //导入
                case 0:
                    OpenExcelImporttoDt(_fileAddress);
                    break;
                //运算
                case 1:

                    break;
                //导出
                case 2:

                    break;
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="fileAddress"></param>
        private void OpenExcelImporttoDt(string fileAddress)
        {
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress);
        }



    }
}
