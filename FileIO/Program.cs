using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.Util;
using System;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using FileIO;
using System.Collections.Generic;

namespace FileIO
{
     public class Program
    {
        static void Main(string[] args)
        {
            string inPath = @"C:\Users\Administrator\Desktop\in.xlsx";
            List<List<string>> list = ExcelTListHelper.ExcelToList(inPath);

            string outPath = @"C:\Users\Administrator\Desktop\out.xlsx";
            ExcelTListHelper.ListToExcel(outPath, list);

        }

    }
}
