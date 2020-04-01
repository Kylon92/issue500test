using ComShared;
using System;

namespace Com_Consumer
{
    class Program
    {
        
        //C# To C# works fine.
        //However C# Server to VBA Consumer crashes
        static void Main(string[] args)
        {
            Type comType = Type.GetTypeFromProgID("ComServer.SQL");
            Object comObj = Activator.CreateInstance(comType);
            ISQL isql = (ISQL)comObj;
            isql.Crash();
        }
    }
}
