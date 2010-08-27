using System.Runtime.InteropServices;

namespace InteropComObjects.IO.Ports {

    public delegate void DataReceivedHandler(object sender, int eventType);
    public delegate void ErrorReceivedHandler(object sender, int eventType);
    public delegate void PinChangedHandler(object sender, int eventType);

    [Guid("54F111C7-0A29-4E5A-88F4-82C09E18E192"),
     InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISerialPortEvents {
        void DataReceived(object sender, int eventType);
        void ErrorReceived(object sender, int eventType);
        void PinChanged(object sender, int eventType);
    }
}