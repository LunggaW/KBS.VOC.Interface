using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace KBS.RANCH.VOC.INTERFACE.TOCANCELLATION
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new ServiceTOCancellation() 
			};
            ServiceBase.Run(ServicesToRun);
        }
    }
}
