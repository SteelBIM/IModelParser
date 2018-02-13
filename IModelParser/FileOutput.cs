using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace IModelParser
{
    class FileOutput
    {
        public void outputTxt(string filepath, Dictionary<string, Dictionary<string, string>> result)
        {
            StringBuilder sb = new StringBuilder();
            foreach(string key in result.Keys)
            {
                sb.AppendLine(key);
                foreach(string propName in result[key].Keys)
                {
                    sb.AppendLine(propName.PadRight(25) + ":" + result[key][propName].PadRight(15));
                }
                sb.AppendLine();
            }
            StreamWriter sw = new StreamWriter(filepath, false);
            sw.Write(sb.ToString());
            sw.Flush();
            sw.Close();
            MessageBox.Show("导出完成！");
        }
        
        public void outputExcel(string filepath, Dictionary<string, Dictionary<string, string>> result)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new ApplicationClass();
            Workbooks workBooks = excelApp.Workbooks;
            Workbook workBook = workBooks.Add(true);
            Sheets sheets = workBook.Sheets;
            Worksheet workSheet = (Worksheet)sheets.get_Item(1);
            workSheet.Cells.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            int rowIndex = 1;
            foreach(string eID in result.Keys)
            {
                workSheet.get_Range(workSheet.Cells[rowIndex, 1], workSheet.Cells[rowIndex, 2]).MergeCells = true;
                workSheet.get_Range(workSheet.Cells[rowIndex, 1], workSheet.Cells[rowIndex, 1]).Interior.Color = Color.AliceBlue;
                workSheet.Cells[rowIndex++, 1] = eID;
                foreach(string propName in result[eID].Keys)
                {
                    workSheet.Cells[rowIndex, 1] = propName;
                    workSheet.Cells[rowIndex++, 2] = result[eID][propName];
                }
            }

            workSheet.Cells.Columns.AutoFit();
            if(filepath != null && filepath != "")
            {
                try
                {
                    workBook.Saved = true;
                    workBook.SaveCopyAs(filepath);
                }
                catch(Exception e)
                {
                    MessageBox.Show("保存失败:\t" + e.StackTrace);
                }
            }
            MessageBox.Show("导出完成！");
        }

        public void regex(string filepath, string target)
        {
            List<string> list = new List<string>();
            DirectoryInfo dir = new DirectoryInfo(filepath);
            foreach(FileInfo f in dir.GetFiles())
            {
                FileStream fs = new FileStream(filepath + f.Name, FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(fs, Encoding.Default);
                string content = reader.ReadToEnd();
                reader.Close();
                fs.Close();
                if(content.Contains(target))
                {
                    list.Add(f.Name);
                }
            }

            if(list.Count >= 0)
            {
                string content = "";
                foreach(string str in list)
                {
                    content += str + "\n";
                }
                MessageBox.Show(content,"结果");
            }
            else
            {
                MessageBox.Show("无");
            }
        }
    }
}
