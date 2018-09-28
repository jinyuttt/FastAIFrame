using FastStream.socket;
using System;

namespace ConsoleApp3
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
