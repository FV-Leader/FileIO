using System;
using System.Collections.Generic;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace FileIO
{
    public class ExcelTListHelper
    {
        public static List<List<string>> ExcelToList(string filePath,string sheetName = "Sheet1")
        {
            //初始化
            IWorkbook wb;
            ISheet sheet;
            List<List<string>> ExcelList = new();
            FileStream file = null;

            try
            {
                //获取拓展名
                string extension = Path.GetExtension(filePath);
                file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                //获取相应工作簿
                if (extension.Equals(".xls"))
                {
                    wb = new HSSFWorkbook(file);
                }
                else
                {
                    wb = new XSSFWorkbook(file);
                }

                //获取工作表
                if(sheetName.Equals("Sheet1"))
                {
                    sheet = wb.GetSheetAt(0);
                }
                else
                {
                    sheet = wb.GetSheet(sheetName);
                }

                //读取到List中
                for(int i = 0;i<= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    List<string> CellList = new();
                    for (int j =0;j< row.LastCellNum; j++)
                    {
                        string cell = row.GetCell(j).ToString();
                        CellList.Add(cell);
                    }
                    ExcelList.Add(CellList);
                }
            }
            catch(Exception e)
            {
                if (file != null)
                {
                    file.Close();
                }
                Console.WriteLine("导入报错：" + e.Message);
                return null;
            }
            return ExcelList;
        }

        public static void ListToExcel(string filepath, List<List<string>> list, string sheetName = "Sheet1")
        {
            //初始化
            FileStream file = null;
            XSSFWorkbook wb;
            ISheet sheet;

            try
            {
                //创建文件流
                file = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                //创建工作簿
                wb = new XSSFWorkbook();
                //创建工作表
                sheet = wb.CreateSheet(sheetName);
                //List写入EXCEL
                for (int i = 0; i < list.Count; i++)
                {
                    IRow row = sheet.CreateRow(i);
                    for (int j = 0; j < list[i].Count; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        cell.SetCellValue(list[i][j]);
                    }
                }
                wb.Write(file);
            }
            catch (Exception e)
            {
                Console.WriteLine("导出报错：" + e.Message);
                if (file != null)
                {
                    file.Close();
                }
            }

        }


    }
}
