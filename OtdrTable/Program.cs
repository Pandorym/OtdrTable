using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OtdrTable {
    static class Program {
        static public Option Options;
        

        static void Main(string[] args) {
            Conex.InitConsole();
            CloseDisposal.BindingHandler();

            Options = GetOptions(args, "");
            ShowOptions();
            
            do {
                try {
                    String inputFilepath = GetFilepath_And_DisposeInput();
                    XlsxInfo[] projectInfo = Input.Getparamete(inputFilepath);

                    Mark(projectInfo); // Mark(Input.Getparamete(GetFilepath_And_DisposeInput()));
                }
                catch (Exception e) {
                    if (e.Message == "_EXIT") break;
                    if (e.Message == "_CONTINUE") continue;
                    Conex.ErrorLine("Unknown error.");
                }
            } while (!Options.Once);

            CloseDisposal.Execute();
            return;
        }

        static private void Mark(XlsxInfo[] projectInfo) {
            Conex.CursorVisible = false;
            Int32 CursorTop = Console.CursorTop + 1;
            Output.ShowInfo(projectInfo);
            Conex.SaveCursorPosition();
            Int32 start_TC = Environment.TickCount;
            Parallel.For(0, projectInfo.Length, i => new Output().ExportingXlsx(projectInfo[i], CursorTop + i));
            Conex.RevertCurserPosition();
            Conex.Info("    All done.");
            Conex.DebugLine(" Time consuming: {0} ms\n", Environment.TickCount - start_TC);
            Conex.CursorVisible = true;
        }

        static private void ShowOptions() {
            Conex.DebugLine("Once: {0}", Options.Once);
            Conex.DebugLine("args: {0}", String.Join(" ", Options.args));
            Conex.DebugLine("MarkGytsMode: {0}", Options.MarkGytsMode);
        }

        // 處理用戶在命令行中的輸入。
        // 直到輸入正確的文件路徑，才返回。
        // Re: 如果輸入的是「輸入文件路徑」，將返回文件路徑。
        static private String GetFilepath_And_DisposeInput() {
            String inputFilepath = null;
            do {
                Conex.Write("Input filename: ");

                String input = Options.Once ? String.Join(" ",Options.args) : Conex.ReadLine();
                if (input.ToUpper() == "EXIT" || input.Trim() == "") {
                    Conex.WriteLine("_So long.");
                    throw new Exception("_EXIT");
                }
                try {
                    inputFilepath = Path.GetFullPath(input);
                }
                catch {
                    Conex.ErrorLine("    Invalid filename.");
                }
            } while (inputFilepath == null && !Options.Once);
            if (inputFilepath == null) throw new Exception("_EXIT"); //  Options.Once == true
            Conex.WriteLine("Filepath: {0}", inputFilepath);
            return inputFilepath;
        }

        static private Option GetOptions(String[] args, String OptionFilePath) {
            Option Options = new Option() {
                Once = args.Length != 0,
                args = args
            };
            return Options;
        }
    }
}
