using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;

namespace FastStream.office
{
  public  class EecelTool
    {

        /// <summary>
        /// 导出EXCEL
        /// </summary>
        /// <param name="filePath">excel文件</param>
        /// <param name="dts">数据</param>
       public void ExportTable(string filePath,List<DataTable> dts)
        {
            HSSFWorkbook wb;
            FileStream file;
            file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            wb = new HSSFWorkbook(file);
            int num = 0;
            foreach (DataTable dt in dts)
            {
                string sheetName = dt.TableName == "" ? "sheet"+(++num).ToString() : dt.TableName;
                var sheet = wb.CreateSheet(sheetName);
                //标题
                var row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.CreateCell(i, CellType.String).SetCellValue(dt.Columns[i].ColumnName);
                }
                //数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());

                    }
                }
            }
            file.Close();
 
        }

       /// <summary>
       /// 按照DataTble的名称填充
       /// </summary>
       /// <param name="templepte">模板</param>
       /// <param name="fileDir">导出文件的目录</param>
       /// <param name="dts">数据</param>
        public void ExportTempleteByName(string templepte,string fileDir, List<DataTable> dts)
        {
            HSSFWorkbook wb;
            FileStream file;
            FileInfo fileInfo = new FileInfo(templepte);
            string filePath = Path.Combine(fileDir, DateTime.Now.ToString("yyyyMMddHHmmss") + fileInfo.Extension);
            fileInfo.CopyTo(filePath);
            file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Read);
            wb = new HSSFWorkbook(file);
            int num = 0;
            foreach (DataTable dt in dts)
            {
                string sheetName = dt.TableName == "" ? "sheet" + (++num).ToString() : dt.TableName;
                var sheet = wb.GetSheet(sheetName);
                //标题
                var row = sheet.GetRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.GetCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                //数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = sheet.GetRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row.GetCell(j).SetCellValue(dt.Rows[i][j].ToString());

                    }
                }
            }
            file.Close();
        }

        /// <summary>
        /// 按照顺序填充模板
        /// </summary>
        /// <param name="templepte">模板</param>
        /// <param name="fileDir">导出文件目录</param>
        /// <param name="dts">数据</param>
        public void ExportTemplete(string templepte, string fileDir, List<DataTable> dts)
        {
            HSSFWorkbook wb;
            FileStream file;
            FileInfo fileInfo = new FileInfo(templepte);
            string filePath = Path.Combine(fileDir, DateTime.Now.ToString("yyyyMMddHHmmss") + fileInfo.Extension);
            fileInfo.CopyTo(filePath);
            file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Read);
            wb = new HSSFWorkbook(file);
            int num = 0;
            foreach (DataTable dt in dts)
            {
               
                var sheet = wb.GetSheetAt(num);
                num++;
                //标题
                var row = sheet.GetRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.GetCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                //数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = sheet.GetRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row.GetCell(j).SetCellValue(dt.Rows[i][j].ToString());

                    }
                }
            }
            file.Close();
        }


        /// <summary>
        /// 导出图片
        /// </summary>
        /// <param name="filePath">excel文件</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="img">文件路径</param>
        /// <param name="mGOffet">设置项</param>
        public void ExportIMG(string filePath,string sheetName, string img,IMGOffet mGOffet )
        {
            HSSFWorkbook wb;
            FileStream file;
            file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Read);
            wb = new HSSFWorkbook(file);
           ISheet sheet= wb.GetSheet(sheetName);
            if(sheet==null)
            {
                wb.CreateSheet(sheetName);
            }
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            SetPic(wb, patriarch, img, sheet,mGOffet);

        }
        private void SetPic(HSSFWorkbook workbook, HSSFPatriarch patriarch, string path, ISheet sheet, IMGOffet mGOffet)
        {
            if (string.IsNullOrEmpty(path)) return;
            byte[] bytes = System.IO.File.ReadAllBytes(path);
            int pictureIdx = workbook.AddPicture(bytes, PictureType.JPEG);
            // 插图片的位置  HSSFClientAnchor（dx1,dy1,dx2,dy2,col1,row1,col2,row2) 后面再作解释
            HSSFClientAnchor anchor = new HSSFClientAnchor(mGOffet.StartOffsetX, mGOffet.StartOffsetY, mGOffet.EndOffsetX, mGOffet.EndOffsetY, mGOffet.StartCellCol, mGOffet.StartCellRow, mGOffet.EndCellCol, mGOffet.EndCellRow);
            //把图片插到相应的位置
            HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
        }
    }
}
