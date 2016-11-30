using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OtdrTable
{
    static class Program
    {
        static Boolean Once;
        static private String filepath;
        static private String[,] inputInfo;
        static public XlsxInfo[] xlsxInfo;

        static void Main(string[] args)
        {
            Conex.InitConsole();
            try {
                // 指令后帶文件名 只執行一次
                Once = args.Length != 0;
                Conex.DebugLine("Once: {0}", Once);
                Conex.DebugLine("args: {0}", Once ? String.Join(" ", args) : "null");
                do {
                    if (Once) filepath = args[0];
                    else {
                        Conex.Write("Input filename: ");
                        filepath = Conex.ReadLine();
                        if (filepath.ToUpper() == "EXIT") return;
                    }
                    //Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Conex.WriteLine("Filepath: {0}", filepath);
                    if (!(readFile() && Getparamete())) {
                        if (Once) return;
                        else continue;
                    }

                    Mark();
                } while (!Once);
            }
            catch {
                Conex.ErrorLine("Unknown error.");
            }
            finally {
                Conex.SaveLog();
            }
            return;
        }
        
        static private void Mark() {
            Console.CursorVisible = false;
            Int32 CursorTop = Console.CursorTop + 1;
            ShowInfo();
            Int32 ContextCursorTop = Console.CursorTop;
            Parallel.For(0,xlsxInfo.Length,i => new Xlsx().ExportingXlsx(xlsxInfo[i], CursorTop + i));
            Console.SetCursorPosition(0, ContextCursorTop);
            Conex.InfoLine("    All done.\n");
            Console.CursorVisible = true;
        }


        static private Boolean readFile() {
            try {
                switch (filepath.Split('.').Last().ToUpper()) {
                    case "TXT":
                        inputInfo = Xlsx.ReadTxt(filepath);
                        break;
                    case "XLS":
                    case "XLSX":
                        inputInfo = Xlsx.ReadXlsx(filepath);
                        break;
                    default:
                        Conex.WarnLine("    Unknown file format, try as 'txt'.");
                        inputInfo = Xlsx.ReadTxt(filepath);
                        break;
                }
            }
            catch {
                Conex.ErrorLine("    Failed to read file.");
                return false;
            }
            Conex.InfoLine("    Read successfully.");
            return true;
        }

        static public Boolean Getparamete() {
            try {
                // 0 = xlsxName; 1 = imgName; 2 = chainLength; 3 = overallLength; 4 = Dot B y; 
                xlsxInfo = new XlsxInfo[inputInfo.GetLength(0)];

                Int32 s = 1; // 規避重名
                Random random = new Random();
                // Processing incoming
                for (Int32 i = 0; i < xlsxInfo.Length; i++) {
                    Int32 x = i;
                    xlsxInfo[i].projectName = inputInfo[i, 0];
                    xlsxInfo[i].gyts = Int32.Parse(inputInfo[i, 2].Remove(inputInfo[i, 2].Length - 1, 1).Remove(0, 5));
                    if (i > 0 && inputInfo[i - s, 0] == inputInfo[i, 0]) {
                        s++;
                        xlsxInfo[i].xlsxName = inputInfo[i, 0] + s.ToString("D3") + "曲线图.xlsx";
                    }
                    else {
                        xlsxInfo[i].xlsxName = inputInfo[i, 0] + "曲线图.xlsx";
                        if (s != 1) {
                            xlsxInfo[i - s].xlsxName = inputInfo[i - s, 0] + "001曲线图.xlsx";
                            s = 1;
                        }
                    }
                    xlsxInfo[i].chainLength = Double.Parse(inputInfo[i, 1]) * 1000;

                    // hypothesis
                    xlsxInfo[i].imgName = inputInfo[i, 0];
                    xlsxInfo[i].overallLength = Math.Max(1020, xlsxInfo[i].chainLength * random.Next(200, 250) / 100);
                    xlsxInfo[i].By = Convert.ToDouble(random.Next(12000, 17000)) / 1000;
                }
            }
            catch {
                Conex.ErrorLine("    Analytic failure");
                return false;
            }
            Conex.InfoLine("    Analytical success.");
            return true;
        }
        static public void ShowInfo() {
            Int32 gytsSum = 0;
            for (Int32 i = 0; i < xlsxInfo.Length; i++)
                gytsSum += xlsxInfo[i].gyts;
            Conex.DebugLine("Project sum: {0}; Gyts Sum: {1}.", xlsxInfo.Length, gytsSum);
            for (Int32 i = 0; i < xlsxInfo.Length; i++)
                Conex.WriteLine("{0,2}. Wait   0/{1,-3} {2,10:F2} {3,10:F2} {4,6:F3} {5}",
                    (i + 1).ToString(),
                    xlsxInfo[i].gyts ,
                    xlsxInfo[i].chainLength,
                    xlsxInfo[i].overallLength, 
                    xlsxInfo[i].By, 
                    xlsxInfo[i].imgName);
        }
    }
}
