/*Post Build Instructions
 * 
 * Copy the DLL file to: C:\Program Files\Corepoint Health\Custom Objects
 * Open a command prompt as administator
 * Change directory to the Custom Objects folder
 * Run the following command to register the DLL:
 * %systemroot%\Microsoft.NET\Framework64\v4.0.30319\regasm IInvokeTest.dll /tlb /nologo
 */

using System;
using System.Runtime.InteropServices;
using CorepointHealth.Common.Interfaces.Gear.ItemInvoke;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IInvokeTest
{

    public class Action
    {
        public string Name { get; set; }
        public string Parameters { get; set; }

    }

    /* Guid instructions
     * The following Guid identifies the DLL and should be replaced in any new com object you create.
     * New Guids can be generated here: https://www.guidgen.com/
     */

    [Guid("732d9f97-c7ca-40f9-9404-2d08aeac1dee")]

    public class CorepointCom : IInvoke
    {
        /* Corepoint IInvoke
         * Corepoint calls the Invoke function to interact with an external Dll.
         * 
         * "sourceArgument is passed in from the "Source Operand" field 
         * "options" is passed in by the "Options" field
         * "resultData" is returned to Corepoint in the "Destination Operand" variable
         * "error" is returned in the "Status" variable
         */
    public void Invoke(ref string sourceOperand, ref string options, ref string destinationOperand, ref string status)
        {
            Action action = JsonConvert.DeserializeObject<Action>(options);

            switch (action.Name)
            {
                case "Test":
                    // Just a test response
                    destinationOperand = "The test was successful.";
                    status = "Success";
                    return;
                default:
                    break;
            }

        }

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey key;
            key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(
              "CLSID\\" + t.GUID.ToString("B") + "\\Implemented Categories\\{740B2584-57D6-46CB-A85D-B2D255115A97}");

            if (key != null)
            {
                key.Close();
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(
              "CLSID\\" + t.GUID.ToString("B") + "\\Implemented Categories\\{740B2584-57D6-46CB-A85D-B2D255115A97}", false);
        }
    }

}
