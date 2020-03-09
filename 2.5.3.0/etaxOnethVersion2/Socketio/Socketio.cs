using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quobject.SocketIoClientDotNet.Client;
using System.IO;

namespace ETAXStartup.Socketio
{
    class Socketio
    {
        private Socket socket;
        public string ip = "http://localhost:8000";
        public void connectSocketio()
        {
            this.socket = IO.Socket(ip);
        }
    }
}
