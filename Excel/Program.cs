using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application excelApp = new Application();
            excelApp.Visible = false;

            Workbook workbook = excelApp.Workbooks.Open(args[0]);
            excelApp.Run("Update"); // Запускаем макрос

            // Сохраняем и закрываем файл дубликата таблицы расчета Excel, сама таблица расчета не сохраняется
            if (File.Exists(args[1]))
                File.Delete(args[1]);
            workbook.SaveAs (args[1], Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close(false);
            excelApp.Quit();

            //Освобождаем ресурсы
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}