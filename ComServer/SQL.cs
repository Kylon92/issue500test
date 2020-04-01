using ComShared;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ComServer
{
    internal sealed class ContractGuids
    {
        public const String TestGUID = "C6B61646-D342-4B5E-9D7A-D7A8776A453A";
    }
    
    [ComVisible(true)]
    [Guid(ContractGuids.TestGUID)]
    public class SQL : ISQL
    {
        //dotnet publish -r win-x64 -c Debug --self-contained false --framework netcoreapp3.1 "ComServer\ComServer.csproj"
        //regsvr32.exe "%YOURPUBLISHFOLDER%\ComServer.comhost.dll"
        //publish first, then use regsvr32 to register it to COM
        //I recommend testing with VBA/Excel, do note that if your Excel is 64bit you need to publish and reg with win-x64 otherwise as win-x86 for 32Bit Excel.
        //C# COM to C# COM works fine because you'll have a Assembly.GetEntryAssembly() however not calling from C#, like from VBA or VB6 will cause the library to crash
        public void Crash()
        {
            //keep this line to attach a debugger from VBA to Visual Studio
            System.Diagnostics.Debugger.Launch();
            try
            {
                //you may need to change your Database Connection
                SqlConnection Connection = new SqlConnection("Server=(localdb)\\mssqllocaldb;Database=master;Trusted_Connection=True");
                Connection.Open();
            }
            catch(Exception ex)
            {

            }
        }
    }
}
