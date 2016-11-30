using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OtdrTable {
    static class Conex {
        const String LogPath = @"C:\Log\OtdrTeble\";

        static StreamWriter Log;

        static Boolean Record = true;

        static ConsoleColor ForegroundColor = Console.ForegroundColor;
        static ConsoleColor BackgroundColor = Console.BackgroundColor;
        
        static public void InitConsole() {
            Console.Title = "Otdr Table";
            Console.WindowWidth = 120;
            new DirectoryInfo(LogPath).Create();
            Log = new StreamWriter(LogPath + "log_" + (DateTime.Now.ToUniversalTime().Ticks - 621355968000000000) / 10000000);
        }

        static public void SaveLog() {
            Log.Close();
        }

        static public String ReadLine() {
            String str = Console.ReadLine();
            Log.WriteLine(str);
            return str;
        }

        static public void Write(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.Write(ourput);
            if (Record) Log.Write(ourput);
        }

        static public void WriteLine(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.WriteLine(ourput);
            if (Record) Log.WriteLine(ourput);
        }

        static public void Debug(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.Write(ourput);
        }

        static public void DebugLine(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.WriteLine(ourput);
        }

        static public void Info(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.Write(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.Write(ourput);
        }

        static public void InfoLine(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.WriteLine(ourput);
        }

        static public void Warn(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.Write(ourput);
        }

        static public void WarnLine(String message,params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.WriteLine(ourput);
        }

        static public void Error(String message, params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.Write(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.Write(ourput);
        }

        static public void ErrorLine(String message,params Object[] arg) {
            String ourput = String.Format(message, arg);
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine(ourput);
            Console.ForegroundColor = ForegroundColor;
            if (Record) Log.WriteLine(ourput);
        }
    }
}
