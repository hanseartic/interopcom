using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace InteropComObjects {
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IComPort {
        String Device { get; set; }
        int LastError { get; }

        void Sleep(int length);
        string ReadString();
        int GetDeviceCount();
        object GetDevice(int deviceNumber);
        String GetErrorDescription(int errorId);

        #region SerialPort-Wrapper
        #region Static members
        /// <summary>Gets an array of serial port names for the current computer.
        /// </summary>
        /// <returns>The names of the currently available serial ports.</returns>
        String[] GetPortNames();
        #endregion

        #region Instance Members
        #region Mehods
        /// <summary>Opens a new serial port connection.
        /// </summary>
        void Open();
        /// <summary>Closes the port connection, sets the
        /// <see cref="System.IO.Ports.SerialPort.IsOpen"/> Property to false,
        ///  and disposes of the internal <see cref="System.IO.Stream"/> object.
        /// </summary>
        void Close();
        /// <summary>Discards data from the serial driver's receive buffer.
        /// </summary>
        void DiscardInBuffer();
        /// <summary>Discards data from the serial driver's transmit buffer.
        /// </summary>
        void DiscardOutBuffer();
        int Read(byte[] buffer, int offset, int count);
        int Read(char[] buffer, int offset, int count);
        int ReadByte();
        int ReadChar();
        String ReadExisting();
        String ReadLine();
        String ReadTo(String value);
        void Write(byte[] buffer, int offset, int count);
        void Write(char[] buffer, int offset, int count);
        void Write(String value);
        void WriteLine(String value);
        #endregion Methods

        #region Properties
        /// <summary>Gets the underlying <see cref="System.IO.Stream"/> object 
        /// for a <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        Stream BaseStream { get; }
        /// <summary>Gets or sets the signal baud rate.
        /// </summary>
        int BaudRate { get; set; }
        /// <summary>Gets or sets the break signal state.
        /// </summary>
        bool BreakState { get; set; }
        /// <summary>Gets the number of bytes of data in the receive buffer.
        /// </summary>
        int BytesToRead { get; }
        /// <summary>Gets the number of bytes of data in the send buffer.
        /// </summary>
        int BytesToWrite { get; }
        /// <summary>Gets the state of the Carrier Detect Line for the port.
        /// </summary>
        bool CDHolding { get; }
        /// <summary>Gets the state of the Clear-to-Send line.
        /// </summary>
        bool CtsHolding { get; }
        /// <summary>Gets or sets the standard length of data bits per byte.
        /// </summary>
        int DataBits { get; set; }
        /// <summary>Gets or sets a value indicating whether null bytes are ignored
        /// when transmitted between the port and the receive buffer.
        /// </summary>
        bool DiscardNull { get; set; }
        /// <summary>Gets the state of the Data Set Ready (DSR) Signal.
        /// </summary>
        bool DsrHolding { get; }
        /// <summary>Gets or sets a value that enables the Data Terminal Ready
        /// (DTR) signal during serial communication.
        /// </summary>
        bool DtrEnable { get; set; }
        /// <summary>Gets or sets the byte encoding for pre- and post-transmission
        /// conversion of text.
        /// </summary>
        Encoding Encoding { get; set; }
        /// <summary>Gets or sets the handshaking protocol for serial port transmission of data.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>
        /// <value>XOnXOff</value>,
        /// <value>XOnXOff</value>,
        /// <value>RequestToSend</value>,
        /// <value>RequestToSendXOnXOff</value>
        /// </remarks>
        String Handshake { get; set; }
        /// <summary>Gets a value indicating the open or closed status of the
        /// <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        bool IsOpen { get; }
        /// <summary>Gets or sets the value used to interpret the end of a call to the 
        /// <see cref="System.IO.Ports.SerialPort.ReadLine()"/> and 
        /// <see cref="System.IO.Ports.SerialPort.WriteLine(System.String)"/> methods.
        /// </summary>
        string NewLine { get; set; }
        /// <summary>Gets or sets the parity-checking protocol.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>,
        /// <value>Odd</value>,
        /// <value>Even</value>,
        /// <value>Mark</value>,
        /// <value>Space</value>
        /// </remarks>
        String Parity { get; set; }
        /// <summary>Gets or sets the byte that replaces invalid bytes in a data stream
        /// when a parity error occurs.
        /// </summary>
        byte ParityReplace { get; set; }
        /// <summary>Gets or sets the port for communications, including but 
        /// not limited to all available COM ports.
        /// </summary>
        string PortName { get; set; }
        /// <summary>Gets or sets the size of the
        /// <see cref="System.IO.Ports.SerialPort"/> input buffer.
        /// </summary>
        int ReadBufferSize { get; set; }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a read operation does not finish.
        /// </summary>
        int ReadTimeout { get; set; }
        int ReceivedBytesThreshold { get; set; }
        /// <summary>Gest or sets a value indicating wheter the Request to Send
        /// (RTS) signal is enabled during serial communication.
        /// </summary>
        bool RtsEnable { get; set; }
        /// <summary>Gets or sets the standard number of stopbits per byte.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>,
        /// <value>One</value>,
        /// <value>Two</value>,
        /// <value>OnePointFive</value>
        /// </remarks>
        String StopBits { get; set; }
        /// <summary>Gets or sets the size of the serial port output buffer.
        /// </summary>
        int WriteBufferSize { get; set; }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a write operation does not finish.
        /// </summary>
        int WriteTimeout { get; set; }
        #endregion Properties
        #endregion Insatance Members
        #endregion SerialPort-Wrapper

    }

    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ComPort : IComPort {

        private SerialPort selectedPort;
        private String[] portNames;
        private static Dictionary<int, string> errors;
        private Queue<string> inQueue;
        //private Queue<byte> outQueue;

        public ComPort() {
            errors = new Dictionary<int, string>();
            inQueue = new Queue<string>();
        }

        #region SerialPort-Wrapper
        // The static members
        #region Static Members
        /// <summary>Gets an array of serial port names for the current computer.
        /// </summary>
        /// <returns>The names of the currently available serial ports.</returns>
        public String[] GetPortNames() {
            return SerialPort.GetPortNames();
        }
        #endregion

        // Instance-accessible members
        #region Instance Members

        #region Methods
        /// <summary>Opens a new serial port connection.
        /// </summary>
        public void Open() {
            var worker = new BackgroundWorker();
            worker.DoWork += OpenAsyncWorker;
            worker.RunWorkerAsync();
        }
        /// <summary>Closes the port connection, sets the
        /// <see cref="System.IO.Ports.SerialPort.IsOpen"/> Property to false,
        ///  and disposes of the internal <see cref="System.IO.Stream"/> object.
        /// </summary>
        public void Close() {
            if (selectedPort != null && selectedPort.IsOpen) {
                selectedPort.Close();
                selectedPort = null;
            }
        }
        /// <summary>Discards data from the serial driver's receive buffer.
        /// </summary>
        public void DiscardInBuffer() {
            if (null != selectedPort) selectedPort.DiscardInBuffer();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Discards data from the serial driver's transmit buffer.
        /// </summary>
        public void DiscardOutBuffer() {
            if (null != selectedPort) selectedPort.DiscardOutBuffer();
            throw new InvalidOperationException("No active port.");
        }
        public int Read(byte[] buffer, int offset, int count) {
            if (null != selectedPort) return selectedPort.Read(buffer, offset, count);
            throw new InvalidOperationException("No active port.");
        }
        public int Read(char[] buffer, int offset, int count) {
            if (null != selectedPort) return selectedPort.Read(buffer, offset, count);
            throw new InvalidOperationException("No active port.");
        }
        public int ReadByte() {
            if (null != selectedPort) return selectedPort.ReadByte();
            throw new InvalidOperationException("No active port.");
        }
        public int ReadChar() {
            if (null != selectedPort) return selectedPort.ReadChar();
            throw new InvalidOperationException("No active port.");
        }
        public string ReadExisting() {
            if (null != selectedPort) return selectedPort.ReadExisting();
            throw new InvalidOperationException("No active port.");
        }
        public string ReadLine() {
            if (null != selectedPort) return selectedPort.ReadLine();
            throw new InvalidOperationException("No active port.");
        }
        public string ReadTo(string value) {
            if (null != selectedPort) return selectedPort.ReadLine();
            throw new InvalidOperationException("No active port.");
        }
        public void Write(byte[] buffer, int offset, int count) {
            if (null != selectedPort) selectedPort.Write(buffer, offset, count);
            throw new InvalidOperationException("No active port.");
        }
        public void Write(char[] buffer, int offset, int count) {
            if (null != selectedPort) selectedPort.Write(buffer, offset, count);
            throw new InvalidOperationException("No active port.");
        }
        public void Write(string value) {
            if (null != selectedPort) selectedPort.Write(value);
            throw new InvalidOperationException("No active port.");
        }
        public void WriteLine(string value) {
            if (null != selectedPort) selectedPort.WriteLine(value);
            throw new InvalidOperationException("No active port.");
        }
        #endregion

        #region Properties
        /// <summary>Gets the underlying <see cref="System.IO.Stream"/> object 
        /// for a <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        public Stream BaseStream {
            get {
                if (null != selectedPort) return selectedPort.BaseStream;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the signal baud rate.
        /// </summary>
        public int BaudRate {
            get {
                if (null != selectedPort) return selectedPort.BaudRate;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (null != selectedPort) selectedPort.BaudRate = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the break signal state.
        /// </summary>
        public bool BreakState {
            get {
                if (null != selectedPort) return selectedPort.BreakState;
                throw new InvalidOperationException("No active port");
            }
            set {
                if (null != selectedPort) selectedPort.BreakState = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets the number of bytes of data in the receive buffer.
        /// </summary>
        public int BytesToRead {
            get {
                if (null != selectedPort) return selectedPort.BytesToRead;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets the number of bytes of data in the send buffer.
        /// </summary>
        public int BytesToWrite {
            get {
                if (null != selectedPort) return selectedPort.BytesToWrite;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets the state of the Carrier Detect Line for the port.
        /// </summary>
        public bool CDHolding {
            get {
                if (null != selectedPort) return selectedPort.CDHolding;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets the state of the Clear-to-Send line.
        /// </summary>
        public bool CtsHolding {
            get {
                if (null != selectedPort) return selectedPort.CtsHolding;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the standard length of data bits per byte.
        /// </summary>
        public int DataBits {
            get {
                if (selectedPort != null) return selectedPort.DataBits;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.DataBits = value;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets a value indicating whether null bytes are ignored
        /// when transmitted between the port and the receive buffer.
        /// </summary>
        public bool DiscardNull {
            get {
                if (selectedPort != null) return selectedPort.DiscardNull;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.DiscardNull = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets the state of the Data Set Ready (DSR) Signal.
        /// </summary>
        public bool DsrHolding {
            get {
                if (selectedPort != null) return selectedPort.DsrHolding;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets a value that enables the Data Terminal Ready
        /// (DTR) signal during serial communication.
        /// </summary>
        public bool DtrEnable {
            get {
                if (selectedPort != null) return selectedPort.DtrEnable;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.DtrEnable = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the byte encoding for pre- and post-transmission
        /// conversion of text.
        /// </summary>
        public Encoding Encoding {
            get {
                if (selectedPort != null) return selectedPort.Encoding;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.Encoding = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the handshaking protocol for serial port transmission of data.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>
        /// <value>XOnXOff</value>,
        /// <value>XOnXOff</value>,
        /// <value>RequestToSend</value>,
        /// <value>RequestToSendXOnXOff</value>
        /// </remarks>
        public String Handshake {
            get {
                if (selectedPort != null) return selectedPort.Handshake.ToString();
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) {
                    Handshake result;
                    Enum.TryParse<Handshake>(value, out result);
                    selectedPort.Handshake = result;
                } else
                    throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets a value indicating the open or closed status of the
        /// <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        public bool IsOpen {
            get {
                if (selectedPort != null) return selectedPort.IsOpen;
                throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the value used to interpret the end of a call to the 
        /// <see cref="System.IO.Ports.SerialPort.ReadLine()"/> and 
        /// <see cref="System.IO.Ports.SerialPort.WriteLine(System.String)"/> methods.
        /// </summary>
        public string NewLine {
            get {
                if (selectedPort != null) return selectedPort.NewLine;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.NewLine = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the parity-checking protocol.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>,
        /// <value>Odd</value>,
        /// <value>Even</value>,
        /// <value>Mark</value>,
        /// <value>Space</value>
        /// </remarks>
        public String Parity {
            get {
                if (selectedPort != null) return selectedPort.Parity.ToString();
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) {
                    Parity result;
                    Enum.TryParse(value, out result);
                    selectedPort.Parity = result;
                } else
                    throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the byte that replaces invalid bytes in a
        /// data stream when a parity error occurs.
        /// </summary>
        public byte ParityReplace {
            get {
                if (selectedPort != null) return selectedPort.ParityReplace;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.ParityReplace = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the port for communications, including but 
        /// not limited to all available COM ports.
        /// </summary>
        public string PortName {
            get {
                if (selectedPort != null) return selectedPort.PortName;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.PortName = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the size of the
        /// <see cref="System.IO.Ports.SerialPort"/> input buffer.
        /// </summary>
        public int ReadBufferSize {
            get {
                if (selectedPort != null) return selectedPort.ReadBufferSize;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.ReadBufferSize = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a read operation does not finish.
        /// </summary>
        public int ReadTimeout {
            get {
                if (selectedPort != null) return selectedPort.ReadTimeout;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.ReadTimeout = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the number of bytes in the internal input buffer before 
        /// </summary>
        public int ReceivedBytesThreshold {
            get {
                if (selectedPort != null) return selectedPort.ReceivedBytesThreshold;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.ReceivedBytesThreshold = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gest or sets a value indicating wheter the Request to Send
        /// (RTS) signal is enabled during serial communication.
        /// </summary>
        public bool RtsEnable {
            get {
                if (selectedPort != null) return selectedPort.RtsEnable;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.RtsEnable = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the standard number of stopbits per byte.
        /// </summary>
        /// <remarks>Valid values are:
        /// <value>None</value>,
        /// <value>One</value>,
        /// <value>Two</value>,
        /// <value>OnePointFive</value>
        /// </remarks>
        public String StopBits {
            get {
                if (selectedPort != null) return selectedPort.StopBits.ToString();
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) {
                    StopBits result;
                    Enum.TryParse<StopBits>(value, out result);
                    selectedPort.StopBits = result;
                } else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the size of the
        /// <see cref="System.IO.Ports.SerialPort"/> output buffer.
        /// </summary>
        public int WriteBufferSize {
            get {
                if (selectedPort != null) return selectedPort.WriteBufferSize;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.WriteBufferSize = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a write operation does not finish.
        /// </summary>
        public int WriteTimeout {
            get {
                if (selectedPort != null) return selectedPort.WriteTimeout;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (selectedPort != null) selectedPort.WriteTimeout = value;
                else throw new InvalidOperationException("No active port.");
            }
        }
        #endregion

        #endregion Instance Members
        #endregion SerialPort-Wrapper

        public int GetDeviceCount() {
            portNames = SerialPort.GetPortNames();
            return portNames.Length;
        }

        public object GetDevice(int deviceNumber) {
            //SecurityPermissionFlag UnmanagedCode = SecurityPermissionFlag.UnmanagedCode;
            return portNames[deviceNumber];
        }

        private void OnSerialPortDataReceived(object sender, SerialDataReceivedEventArgs e) {
            try {
                while (true) {
                    try {
                        inQueue.Enqueue(((SerialPort)sender).ReadLine());
                    } catch (TimeoutException) { break; } // no more data at the moment
                }
                ((SerialPort)sender).DiscardInBuffer();
                //LineRead(inQueue.Peek(), null);
                //var serialPortLineReadHandlers = SerialPortLineRead.GetInvocationList();
                //foreach (var serialPortLineReadHandler in serialPortLineReadHandlers) {
                //    serialPortLineReadHandler.DynamicInvoke(inQueue.Peek());
                //}
            } catch (NullReferenceException) {      // no subscriber
            } catch (IOException) {                 // temporary problem with port
            } catch (IndexOutOfRangeException) {    // temporary problem with port
            } catch (ArgumentOutOfRangeException) { // temporary problem with port
            }
        }

        private void OpenAsyncWorker(object sender, DoWorkEventArgs e) {
            if (selectedPort != null && selectedPort.IsOpen) {
                selectedPort.Close();
            }
            Exception ex = new Exception();
            try {
                selectedPort = new SerialPort(Device);
                selectedPort.DataReceived += OnSerialPortDataReceived;
                selectedPort.Open();
            } catch (UnauthorizedAccessException cex) {
                ex = cex;
                errors.Add(LastError + 1, ex.Message);
            } catch (IOException cex) {
                ex = cex;
                errors.Add(LastError + 1, ex.Message);
            } catch (ArgumentOutOfRangeException cex) {
                ex = cex;
                errors.Add(LastError + 1, ex.Message);
            } catch (ArgumentException cex) {
                ex = cex;
                errors.Add(LastError + 1, ex.Message);
            } catch (InvalidOperationException cex) {
                ex = cex;
                errors.Add(LastError + 1, ex.Message);
            }
            throw new InvalidOperationException("Opening failed. See inner exception for details.", ex);
        }

        private readonly AutoResetEvent sleepAutoResetEvent = new AutoResetEvent(false);
        public void Sleep(int length) {
            //Timer sleepTimer = new Timer(tick, null, 0, length);
            sleepAutoResetEvent.WaitOne(length);
        }
        public string ReadString() {
            try {
                return inQueue.Dequeue();
            } catch {
                return string.Empty;
            }
        }

        public String Device { get; set; }

        public int LastError {
            get { return errors.Count - 1; }
        }

        public String GetErrorDescription(int errorId) {
            String result = string.Empty;
            errors.TryGetValue(errorId, out result);
            return result;
        }
    }
}
