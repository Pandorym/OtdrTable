using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace OtdrTable {
    class ConexStream : FileStream {
        public ConexStream(String path, FileMode mode):base(path, mode) {
        }

        public void Write(String message, params Object[] arg) {
            Byte[] buffer = Encoding.Unicode.GetBytes(String.Format(message,arg));
            this.Write(buffer, 0, buffer.Length);
        }

        public void WriteLine(String message, params Object[] arg) {
            this.Write(message + "\r\n", arg);
        }

        public void SetCursorPosition(Int32 left, Int32 top) {
            this.Seek(0, SeekOrigin.Begin);
            Byte[] buffer = new Byte[this.Length];
            this.Read(buffer, 0, Convert.ToInt32(this.Length));
            String[] Lines = Encoding.Unicode.GetString(buffer).Split('\n');
            
            Int32 RNsum = 0;
            for (Int32 lineNum = 0; lineNum < top; lineNum++) {
                Int32 size = Size(Lines[lineNum]);
                if (size > Console.WindowWidth) top -= RNsum = (size - 1) / Console.WindowWidth;
                else RNsum = 0;
            }
            Int32 position = left + Console.WindowWidth * RNsum;
            for (Int32 lineNum = 0; lineNum < top; lineNum++)
                position += Lines[lineNum].Length + 1;
            this.Seek(position * 2, SeekOrigin.Begin);
        }

        public Int32 Size(String str) {
            Int32 size = 0;
            for (Int32 index = 0; index < str.Length; index++)
                if (IsChinese(str[index])) size += 2;
                else size++;
            return size;
        }

        public Boolean IsChinese(Char c) {
            return (c > 0x4e00 && c < 0x9fa5);
        }
    }
}
