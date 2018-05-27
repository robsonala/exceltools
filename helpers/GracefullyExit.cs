using System;
using System.Runtime.Loader;

namespace exceltools.helpers
{
    public class GracefullyExit
    {
		/*public bool KeepRunning { get; private set; }
        
        public GracefullyExit()
		{
			this.KeepRunning = true;
   
            Console.CancelKeyPress += delegate (object sender, ConsoleCancelEventArgs e) {
                e.Cancel = true;
				this.KeepRunning = false;

				Console.WriteLine("Received stop signal in Console.CancelKeyPress");
            };
            
			AssemblyLoadContext.Default.Unloading += delegate (AssemblyLoadContext obj) {
				this.KeepRunning = false;

				Console.WriteLine("Received stop signal in AssemblyLoadContext.Default.Unloading");
            };
        }*/
    }
}
