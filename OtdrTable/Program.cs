using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OtdrTable
{
    static class Program
    {
        static Boolean Once;
        static public XlsxInfo[] projectInfo;

        static void Main(string[] args)
        {
            Conex.InitConsole();
            CloseDisposal.BindingHandler();
            Once = args.Length != 0; // 指令后帶文件名 只執行一次
            Conex.DebugLine("Once: {0}", Once);
            Conex.DebugLine("args: {0}", Once ? String.Join(" ", args) : "null");

            String input;
            try {
                do {
                    if (Once) input = args[0];
                    else {
                        Conex.Write("Input filename: ");
                        input = Conex.ReadLine();
                        if (input.ToUpper() == "EXIT" || input.Trim() == "") {
                            Conex.WriteLine("So long.");
                            return;
                        }
                    }
                    Conex.WriteLine("Filepath: {0}", input);
                    if ((projectInfo = Input.Getparamete(input)) == null) {
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
                CloseDisposal.Execute();
            }
            return;
        }
        
        static private void Mark() {
            Conex.CursorVisible = false;
            Int32 CursorTop = Console.CursorTop + 1;
            Output.ShowInfo(projectInfo) ;
            Conex.SaveCursorPosition();
            Int32 start_TC = Environment.TickCount;
            Parallel.For(0,projectInfo.Length,i => new Output().ExportingXlsx(projectInfo[i], CursorTop + i));
            Conex.RevertCurserPosition();
            Conex.Info("    All done.");
            Conex.DebugLine(" Time consuming: {0} ms\n", Environment.TickCount - start_TC);
            Conex.CursorVisible = true;
        }
    }
}
