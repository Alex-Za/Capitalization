using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Capitalization.Adittional_Classes
{
    class ConsoleMessage
    {
        public delegate void MessageHandler(string message);
        public event MessageHandler ErrorNotification;
        public event MessageHandler MessageNotification;

        public void ErrorMessageTriger(string message)
        {
            ErrorNotification?.Invoke(message);
        }
        public void MessageTriger(string message)
        {
            MessageNotification?.Invoke(message);
        }
    }
}
