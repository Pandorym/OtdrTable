using NPOI.SS.UserModel;
using OtdrTable.NPOI_Ex;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtdrTable {
    static class Input {
        static private Random random = new Random();

        static public XlsxInfo[] Getparamete(String filepath) {
            String[,] inputInfo = readFile(filepath);
            if (inputInfo == null) return null;
            
            XlsxInfo[] projectInfo = new XlsxInfo[inputInfo.GetLength(0)];
            try {
                Int32 s = 1; // 規避重名
                // Processing incoming
                for (Int32 i = 0; i < projectInfo.Length; i++) {
                    Int32 x = i;
                    projectInfo[i].projectName = inputInfo[i, 0];
                    projectInfo[i].gyts = Int32.Parse(inputInfo[i, 2].Remove(inputInfo[i, 2].Length - 1, 1).Remove(0, 5));
                    if (i > 0 && inputInfo[i - s, 0] == inputInfo[i, 0]) {
                        s++;
                        projectInfo[i].xlsxName = inputInfo[i, 0] + s.ToString("D3") + "曲线图.xlsx";
                    }
                    else {
                        projectInfo[i].xlsxName = inputInfo[i, 0] + "曲线图.xlsx";
                        if (s != 1) {
                            projectInfo[i - s].xlsxName = inputInfo[i - s, 0] + "001曲线图.xlsx";
                            s = 1;
                        }
                    }
                    projectInfo[i].chainLength = Double.Parse(inputInfo[i, 1]) * 1000;

                    // hypothesis
                    projectInfo[i].imgName = inputInfo[i, 0];
                    projectInfo[i].overallLength = Math.Max(1020, projectInfo[i].chainLength * random.Next(200, 250) / 100);
                    projectInfo[i].By = Convert.ToDouble(random.Next(12000, 17000)) / 1000;
                }
            }
            catch {
                Conex.ErrorLine("    Analytic failure");
                return null;
            }
            Conex.InfoLine("    Analytical success.");
            return projectInfo;
        }

        static public String[,] readFile(String filepath) {
            String[,] inputInfo = null;
            String suffix = filepath.Split('.').Last().ToUpper();
            try {
                switch (suffix) {
                    case "TXT":
                        inputInfo = ReadTxt(filepath);
                        break;
                    case "XLS":
                    case "XLSX":
                        inputInfo = ReadXlsx(filepath);
                        break;
                    default:
                        Conex.WarnLine("    Unknown file format, try as 'TXT'.");
                        inputInfo = ReadTxt(filepath);
                        break;
                }
            }
            catch {
                Conex.ErrorLine("    Failed to read file.");
                return null;
            }
            Conex.InfoLine("    Read successfully. {0}",suffix);
            return inputInfo;
        }

        static private String[,] ReadXlsx(String filePath) {
            ISheet Sheet;
            using (FileStream fs = File.OpenRead(filePath))
                Sheet = WorkbookFactory.Create(fs).GetSheet("工程量明细表");
            Int32 rowCount = Sheet.LastRowNum - 5 + 1;
            String[,] str = new String[rowCount, 3];
            String notEmpty = null;
            for (Int32 row = 3, Index = 0; row < rowCount + 3; Index++, row++) {
                str[Index, 0] = Sheet.GetCellValue(row, 2) ?? notEmpty;
                str[Index, 1] = Sheet.GetCellValue(row, 4);
                str[Index, 2] = Sheet.GetCellValue(row, 3);
                notEmpty = str[Index, 0];
            }
            return str;
        }

        static private String[,] ReadTxt(String filePath) {
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            String[] Line = sr.ReadToEnd().Split('\n');
            String[,] str = new String[Line.Length,3];
            for (Int32 i = 0; i < Line.Length; i++) {
                if (Line[i][Line[i].Length - 1] == '\r') Line[i] = Line[i].Substring(0, Line[i].Length - 1);
                str[i, 0] = Line[i].Split(' ')[0];
                str[i, 2] = Line[i].Split(' ')[1];
                str[i, 1] = Line[i].Split(' ')[2];
            }
            return str;
        }
    }
}
