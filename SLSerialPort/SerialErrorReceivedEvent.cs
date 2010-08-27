namespace System.IO.Ports {
    public delegate void SerialErrorReceivedEventHandler(object sender, SerialErrorReceivedEventArgs e);

    public class SerialErrorReceivedEventArgs : EventArgs {
        public SerialData EventType;
    }
}