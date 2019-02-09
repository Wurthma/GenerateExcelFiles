using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GenerateEscelFiles_Webform
{
    public partial class GerarExcelComNPOI : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            GerarXLSXComNPOI();
        }

        /// <summary>
        /// Retorna um arquivo excel com o NPOIS sem salvar arquivo em disco rigido (gerado em memory e retornado para navegador).
        /// </summary>
        /// <returns>
        /// Retorna um arquivo excel. 
        /// Mime: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        /// Extensão:.xlsx
        /// </returns>
        /// <remarks>Planilha de exemplo retirada de: https://github.com/tonyqus/npoi/blob/master/examples/xssf/BigGridTest/Program.cs</remarks>
        private void GerarXLSXComNPOI()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");

            for (int rownum = 0; rownum < 10000; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                for (int celnum = 0; celnum < 20; celnum++)
                {
                    ICell Cell = row.CreateCell(celnum);
                    Cell.SetCellValue("Cell: Row-" + rownum + ";CellNo:" + celnum);
                }
            }

            MemoryStream ms = new MemoryStream();
            using (MemoryStream tempStream = new MemoryStream())
            {
                workbook.Write(tempStream);
                var byteArray = tempStream.ToArray();
                ms.Write(byteArray, 0, byteArray.Length);
                Response.Clear();
                Response.ContentType = "application/force-download";
                Response.AddHeader("content-disposition", "attachment;    filename = Excel.xls");
                Response.BinaryWrite(byteArray);
                Response.End();
            }
        }
    }
}