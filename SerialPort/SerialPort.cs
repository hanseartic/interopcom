using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace InteropComObjects.IO.Ports {

    [Guid("30E24601-6E65-4FAF-9999-2B135F0B512F"),
    ClassInterface(ClassInterfaceType.None),
    ComSourceInterfaces(typeof(ISerialPortEvents))]
    public class SerialPort : ISerialPort {

        private readonly System.IO.Ports.SerialPort selectedPort;
        private String[] portNames;
        private static Dictionary<int, string> errors;

        public SerialPort() {
            selectedPort = new System.IO.Ports.SerialPort();
            errors = new Dictionary<int, string>();
        }

        #region SerialPort-Wrapper
        // The (normally) static members
        #region Static Members
        /// <summary>Gets an array of serial port names for the current computer.
        /// </summary>
        /// <returns>The names of the currently available serial ports.</returns>
        public String[] GetPortNames() {
            return System.IO.Ports.SerialPort.GetPortNames();
        }
        #endregion

        // Instance-accessible members
        #region Instance Members

        #region Events
        public event DataReceivedHandler DataReceived;
        public event ErrorReceivedHandler ErrorReceived;
        public event PinChangedHandler PinChanged;
        #endregion

        #region Methods
        /// <summary>Opens a new serial port connection.
        /// </summary>
        public void Open() {
            if (null != selectedPort) {
                if (selectedPort.IsOpen)
                    selectedPort.Close();

                selectedPort.DataReceived += ((sender, e) => {
                    try {
                        DataReceived(sender, (int)(e.EventType));
                    } catch (NullReferenceException) { } // no subscription
                });
                selectedPort.ErrorReceived += ((sender, e) => {
                    try {
                        ErrorReceived(sender, (int)(e.EventType));
                    } catch (NullReferenceException) { } // no subscription
                });
                selectedPort.PinChanged += ((sender, e) => {
                    try {
                        PinChanged(sender, (int)(e.EventType));
                    } catch (NullReferenceException) { } // no subscription
                });
                // pass the possible exceptions up transparently - no handling here
                selectedPort.Open();
            }
        }

        /// <summary>Closes the port connection, sets the
        /// <see cref="System.IO.Ports.SerialPort.IsOpen"/> Property to false,
        ///  and disposes of the internal <see cref="System.IO.Stream"/> object.
        /// </summary>
        public void Close() {
            if (selectedPort != null && selectedPort.IsOpen) {
                selectedPort.Close();
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
        /// <summary>Synchronously reads one byte from the <see cref="System.IO.Ports.SerialPort"/>
        /// input buffer.
        /// </summary>
        /// <returns>The byte, cast to an <see cref="Int32"/>, or -1 if the end of the stream has been read.</returns>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="TimeoutException">
        /// The operation did not complete before the time-out period ended.
        /// - or -
        /// No byte was read.
        /// </exception>
        /// /// <remarks>
        /// This method reads one byte.
        /// 
        /// Use caution when using ReadByte and <see cref="ReadChar"/> together. Switching between reading bytes and reading characters can cause
        /// extra data to be read and/or other unintended behavior. If it is necessary to switch between reading text and reading 
        /// binary data from the stream, select a protocol that carefully defines the boundary between text and binary data, such as
        /// manually reading bytes and decoding the data.
        /// </remarks>
        public int ReadByte() {
            if (null != selectedPort) return selectedPort.ReadByte();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Synchronously reads one character from the <see cref="System.IO.Ports.SerialPort"/>
        /// input buffer.
        /// </summary>
        /// <returns>The character that was read.</returns>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="TimeoutException">
        /// The operation did not complete before the time-out period ended.
        /// - or -
        /// No character was available in the allotted time-out period.
        /// </exception>
        /// <remarks>
        /// This method reads one complete character based on the encoding.
        /// 
        /// Use caution when using <see cref="ReadByte"/> and ReadChar together. Switching between reading bytes and reading characters can cause
        /// extra data to be read and/or other unintended behavior. If it is necessary to switch between reading text and reading 
        /// binary data from the stream, select a protocol that carefully defines the boundary between text and binary data, such as
        /// manually reading bytes and decoding the data.
        /// </remarks>
        public int ReadChar() {
            if (null != selectedPort) return selectedPort.ReadChar();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Reads all immediately available bytes, based on the encoding, in both the stream and the input buffer of the SerialPort object.
        /// </summary>
        /// <returns>The contents of the stream and the input buffer of the SerialPort object.</returns>
        public string ReadExisting() {
            if (null != selectedPort) return selectedPort.ReadExisting();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Reads up to the <see cref="System.IO.Ports.SerialPort.NewLine"/>
        /// value in the input buffer.
        /// </summary>
        /// <returns>The contents of the input buffer up to the first occurrence of a
        /// <see cref="System.IO.Ports.SerialPort.NewLine"/> value.</returns>
        /// <exception cref="InvalidOperationException">The specified port is not open. </exception>
        /// <exception cref="TimeoutException">
        /// The operation did not complete before the time-out period ended.
        /// - or -
        /// No bytes were read.
        /// </exception>
        /// <exception cref="InvalidOperationException">No connection to a serial-port.</exception>
        public string ReadLine() {
            if (null != selectedPort) return selectedPort.ReadLine();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Reads a string up to the specified value in the input buffer.
        /// </summary>
        /// <param name="value">A value that indicates where the read operation stops.</param>
        /// <returns>The contents of the input buffer up to the specified value.</returns>
        /// <exception cref="ArgumentException">The length of the value parameter is 0.</exception>
        /// <exception cref="ArgumentNullException">The value parameter is null.</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open. </exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        /// <exception cref="InvalidOperationException">No connection to a serial-port.</exception>
        public string ReadTo(string value) {
            if (null != selectedPort) return selectedPort.ReadLine();
            throw new InvalidOperationException("No active port.");
        }
        /// <summary>Writes a specified number of bytes to the serial port using data from a buffer.
        /// </summary>
        /// <param name="buffer">The byte array that contains the data to write to the port. </param>
        /// <param name="offset">The zero-based byte offset in the buffer parameter at which to begin copying bytes to the port. </param>
        /// <param name="count">The number of bytes to write. </param>
        /// <exception cref="ArgumentNullException">The buffer passed is null.</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentOutOfRangeException">The offset or count parameters are outside a valid region of the buffer being passed. Either offset or count is less than zero.</exception>
        /// <exception cref="ArgumentException">offset plus count is greater than the length of the buffer.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        public void Write(byte[] buffer, int offset, int count) {
            if (null != selectedPort) selectedPort.Write(buffer, offset, count);
            else throw new InvalidOperationException("No active port.");
        }
        /// <summary>Writes a specified number of characters to the serial port using data from a buffer.
        /// </summary>
        /// <param name="buffer">The character array that contains the data to write to the port.</param>
        /// <param name="offset">The zero-based byte offset in the buffer parameter at which to begin copying bytes to the port.</param>
        /// <param name="count">The number of characters to write.</param>
        /// <exception cref="ArgumentNullException">The buffer passed is null.</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentOutOfRangeException">The offset or count parameters are outside a valid region of the buffer being passed. Either offset or count is less than zero.</exception>
        /// <exception cref="ArgumentException">offset plus count is greater than the length of the buffer.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        public void Write(char[] buffer, int offset, int count) {
            if (null != selectedPort) selectedPort.Write(buffer, offset, count);
            else throw new InvalidOperationException("No active port.");
        }
        /// <summary>Writes the specified string to the serial port.
        /// </summary>
        /// <param name="value">The string for output.</param>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentNullException">The string passed is null.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        public void Write(string value) {
            if (null != selectedPort) selectedPort.Write(value);
            else throw new InvalidOperationException("No active port.");
        }
        /// <summary>Writes the specified string and the NewLine value to the output buffer.
        /// </summary>
        /// <param name="value">The string to write to the output buffer.</param>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentNullException">The string passed is null.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        public void WriteLine(string value) {
            if (null != selectedPort) selectedPort.WriteLine(value);
            else throw new InvalidOperationException("No active port.");
        }
        #endregion

        #region Properties
        /// <summary>Gets the underlying <see cref="System.IO.Stream"/> object 
        /// for a <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        public object BaseStream {
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
                if (null != selectedPort) return selectedPort.PortName;
                throw new InvalidOperationException("No active port.");
            }
            set {
                if (null != selectedPort) selectedPort.PortName = value;
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
            portNames = System.IO.Ports.SerialPort.GetPortNames();
            return portNames.Length;
        }

        public object GetDevice(int deviceNumber) {
            return portNames[deviceNumber];
        }

        private readonly AutoResetEvent sleepAutoResetEvent = new AutoResetEvent(false);
        public void Sleep(int length) {
            sleepAutoResetEvent.WaitOne(length);
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