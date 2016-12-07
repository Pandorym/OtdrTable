using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OtdrTable {
    static class CloseDisposal {
        [DllImport("kernel32.dll")]
        static private extern Boolean SetConsoleCtrlHandler(ConsoleCtrlDelegate HandlerRoutine, Boolean Add);

        private delegate Boolean ConsoleCtrlDelegate(Int32 dwCtrlType);
        private const Int32 CTRL_CLOSE_EVENT = 2;

        static private ConsoleCtrlDelegate newDelegate = new ConsoleCtrlDelegate(HandlerRoutine);

        static public void BindingHandler() {
            SetConsoleCtrlHandler(newDelegate, true);
        }

        static private Boolean HandlerRoutine(Int32 CtrlType) {
            switch (CtrlType) {
                case CTRL_CLOSE_EVENT:
                    Execute();
                    break;
            }
            return false;
        }

        static public void Execute() {
            Conex.SaveLog();
        }
    }
}
