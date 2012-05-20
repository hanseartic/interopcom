Installation
============

* Download the source-code
* Install the COM interop [Assemblies](http://interopcom.codeplex.com/releases/view/51518)
  * via the [installer](http://www.codeplex.com/Download?ProjectName=interopcom&DownloadId=250419) installer **OR**
  * Create the interop object on your machine
     * Open the downloaded solution.
     * Open the "ComPort" project's properties and check "Register for COM interop" in the build-tab.
     * Build (**needs elevated rights the first time for COM-registration** -- so I did **not** check in, with the "Register for COM interop" checked). If you get an error about not being able to register the SerialPort.dll you may have forgotten to open the solution as administrator.
* Open and build the SLSerialPort project.
* Create a new Silverlight app to consume the COM object.
  * Open the properties of the new project
    * In the "Silverlight" tab check "Enable running application out of the browser"
    * Open "Out-of-Browser Settings ..."
    * Near the bottom of the new dialog check "Require elevated trust when running outside of the browser"
  * Add a reference to SLSerialPort project or resulting assembly.
  * Reference the serial port as in native .NET:

```c#
using System.IO.Ports;
…
namespace SerialPortEnabledApp {
  class SerialPortConnector {
    SerialPort port = new SerialPort("COM1");
    …
  } 
}
```

