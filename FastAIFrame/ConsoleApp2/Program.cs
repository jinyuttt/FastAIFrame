using FastStream.socket;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            NettyConfig.Port = 6666;
            NettyServer server = new NettyServer();
            server.RunServerAsync();
            while (true)
            {
                var data = server.GetNetSocket();
                Console.WriteLine(NettyConfig.NettyEncod.GetString(data.data));

            }
           
        }
    }
}
