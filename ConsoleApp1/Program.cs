using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
namespace ConsoleApp1
{
    class Program
    {
        static string semester = "";
        public struct Course
        {
            string dept_id { get; set; }
            string dept_name { get; set; }
            string course_id { get; set; }
            string class_id { get; set; }
            string class_name { get; set; }
            string teacher_id { get; set; }
            string teacher_name { get; set; }
            public string original { get; set; }
            public string rename { get; set; }
            public Course(IRow row)
            {
                dept_id = row.GetCell(2).StringCellValue + row.GetCell(3).NumericCellValue;
                dept_name = row.GetCell(4).StringCellValue;
                course_id = row.GetCell(5).StringCellValue;
                class_id = row.GetCell(6).StringCellValue;
                class_name = row.GetCell(7).StringCellValue;
                teacher_id = row.GetCell(9).StringCellValue;
                teacher_name = row.GetCell(10).StringCellValue;
                original = semester + teacher_id + course_id + class_id;
                rename = dept_id + "-" + teacher_name + "-" + class_name;
                //Check illegal char in new filename and replace
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    rename = rename.Replace(c, ' ');
                }
            }
        }
        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "改檔名程式";
            Run();
        }
        static void Run()
        {
            while (semester.Length != 4)
            {
                Console.Write("1.輸入學期代碼(EX:1091):");
                semester = Console.ReadLine();
            }
            LoadData();
        }
        static void LoadData()
        {
            Console.WriteLine("\n2.選擇開課資料");
            try
            {
                var dialog = new OpenFileDialog();
                dialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
                dialog.Title = "選擇開課資料";
                var isValid = dialog.ShowDialog();
                if (isValid.Equals(DialogResult.OK))
                {
                    IWorkbook workBook;
                    using (var fs = new FileStream(dialog.FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        string dir = Path.GetDirectoryName(dialog.FileName);
                        Console.WriteLine("\n選取的檔案資料夾 | " + dir);
                        workBook = new XSSFWorkbook(fs);
                        var list = readXls(workBook);
                        createDir(dir);
                        changeFilename(list, dir);
                    }

                }
                else
                {
                    MessageBox.Show("未選擇檔案，程式已結束。", "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[錯誤]"+e.Message);
                Console.ReadLine();
            }
        }
        static List<Course> readXls(IWorkbook workBook)
        {
            var courseList = new List<Course>();
            var sheet = workBook.GetSheetAt(0);
            Console.WriteLine("選取的工作表 | " + sheet.SheetName);
            for (var i = 2; i < sheet.LastRowNum; i++)
            {
                var item = sheet.GetRow(i);
                courseList.Add(new Course(item));
            }
            return courseList;
        }
        static void createDir(string path)
        {
            string[] newDir = { path + @"\Excel\", path + @"\PDF\" };
            Console.WriteLine("\n3.建立資料夾");
            try
            {
                foreach (string newPath in newDir)
                {
                    if (Directory.Exists(newPath))
                    {
                        Console.WriteLine("That path exists already {0}.", newPath);
                        continue;
                    }
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(newPath);
                    Console.WriteLine("The directory was created successfully:{0}", newPath);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                throw;
            }
            finally
            {
                //Console.ReadLine();
            }
        }
        static void changeFilename(List<Course> list, string path)
        {
            Console.WriteLine("\n4.變更檔名");
            string[] extension = { ".xls", ".pdf" };
            string[] targetPath = { path + @"\Excel", path + @"\PDF" };
            int success = 0;
            int failed = 0;
            for (int i = 0; i < extension.Length; i++)
            {
                foreach (Course course in list)
                {
                    string sourceFile = System.IO.Path.Combine(path, course.original + extension[i]);
                    string destFile = System.IO.Path.Combine(targetPath[i], course.rename + extension[i]);
                    try
                    {
                        System.IO.File.Copy(sourceFile, destFile, true);
                        //Console.WriteLine("[更改檔名成功]" + sourceFile + " | " + destFile);
                        success++;
                    }
                    catch (FileNotFoundException e)
                    {
                        failed++;
                        Console.WriteLine("[找不到檔案]" + e.FileName);
                    }

                }
            }
            Console.WriteLine("\n改檔名完成，{0}成功，{1}失敗。", success.ToString(), failed.ToString());
            Console.ReadLine();



        }
    }
}




