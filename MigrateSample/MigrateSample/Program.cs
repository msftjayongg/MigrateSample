using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateSample
{
    /// this sample app shows how to:
    /// 1. Load a local notebook
    /// 2. Load a remote notebook
    /// 3. copy contents of the local notebook to the remote notebook
    /// 4. close the local notebook
    /// 
    class Program
    {
        static void Main(string[] args)
        {
            var program = new Program();

            program.DoIt(args);
        }

        private void DoIt(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: MigrateSample <localfile> <remoteurl>");
                return;
            }
            

        }
    }
}
