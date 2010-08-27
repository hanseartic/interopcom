namespace System.IO.Ports {
    public delegate void SerialPinChangedEventHandler(object sender, SerialPinChangedEventArgs e);

    public class SerialPinChangedEventArgs : EventArgs {
        public SerialData EventType;
    }
}