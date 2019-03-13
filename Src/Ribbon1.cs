using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using DataTable = System.Data.DataTable;
using ExcelAddIn1;
using System;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using ExcelDna.Integration.Extensibility;
using System.Resources;

namespace ClassLibrary7
{
    [ComVisible(true)]

    public class Ribbon : ExcelRibbon
    {
        static Application app = ExcelDnaUtil.Application as Application;


        public override string GetCustomUI(string RibbonID)
        {
            return base.GetCustomUI(RibbonID);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override object LoadImage(string imageId)
        {
            return base.LoadImage(imageId);
        }

        public override void OnAddInsUpdate(ref Array custom)
        {
            base.OnAddInsUpdate(ref custom);
        }

        public override void OnBeginShutdown(ref Array custom)
        {
            base.OnBeginShutdown(ref custom);
        }

        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            base.OnConnection(Application, ConnectMode, AddInInst, ref custom);
        }

        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            base.OnDisconnection(RemoveMode, ref custom);
        }

        public override void OnStartupComplete(ref Array custom)
        {
            base.OnStartupComplete(ref custom);
        }

        public override void RunTagMacro(IRibbonControl control)
        {
            //   base.RunTagMacro(control);

            var method = typeof(Ribbon).GetMethod(control.Tag);
            if (method != null)
            {
                method.Invoke(this, null);
            }
        }



        public void button1_Click()
        {
            Query frm = new Query();
            frm.ImportDataAction += Frm_ImportDataAction;
            frm.ShowDialog();
        }

       


        private static void fillSheet(DataTable dt)
        {
            var workbook = app.ActiveWorkbook;
            _Worksheet sheet = workbook.Sheets.Add();
            if (dt.TableName.Contains("/"))
            {
                dt.TableName = dt.TableName.Replace("/", "_");
            }
            else if (dt.TableName.Contains("\\"))
            {
                dt.TableName = dt.TableName.Replace("\\", "_");
            }
            else if (dt.TableName.Contains("*"))
            {
                dt.TableName = dt.TableName.Replace("*", "_");
            }
            else if (dt.TableName.Contains("["))
            {
                dt.TableName = dt.TableName.Replace("[", "_");
            }
            else if (dt.TableName.Contains("]"))
            {
                dt.TableName = dt.TableName.Replace("]", "_");
            }
            else if (dt.TableName.Contains("?"))
            {
                dt.TableName = dt.TableName.Replace("?", "_");
            }
            sheet.Name = dt.TableName;
            for (var i = 0; i < dt.Columns.Count; i++)
            {
                Range cell = sheet.Cells[1, i + 1] as Range;
                cell.Value = dt.Columns[i].Caption;
            }
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    Range cell = sheet.Cells[i + 2, j + 1] as Range;
                    cell.Value = dt.Rows[i][j].ToString();
                }
            }
        }
        private static void Frm_ImportDataAction(string dataset,string table,string limit,string offset,string tbname)
        {

            var dt = getDataTable(dataset,table,limit,offset);
            dt.Columns.Remove("available");
            dt.TableName = tbname;
            fillSheet(dt);

        }


        private static DataTable getDataTable(string dataset, string table, string limit, string offset)
        {
            //#region 构建DataTable 
            //DataTable dt = new DataTable();
            //dt.Columns.Add(new DataColumn(importDate + "Header"));
            //for (var i = 0; i < 5; i++)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr[0] = importDate + "_RowData" + i;
            //    dt.Rows.Add(dr);
            //}
            //return dt;
            return BLL.GetDataTable(dataset,table,limit,offset);
            //#endregion
        }
        private static DataTable Showtables()
        {
            return BLL.ShowTables();
        }
        public void sd_ImportDataAction()
        {

            var dt = Showtables();
            dt.TableName = "tables";
            fillSheet(dt);

        }
    }


}
