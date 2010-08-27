namespace System.IO.Ports {
    public delegate void SerialDataReceivedEventHandler(object sender, SerialDataReceivedEventArgs e);

    public class SerialDataReceivedEventArgs : EventArgs {
        public SerialData EventType;
    }
}