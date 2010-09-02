using System.Runtime.InteropServices.Automation;

namespace System.IO.Ports {
    
    /// <summary>Makes serial ports accessible to SL4 application through
    /// COM interop functions and provides a nearly complete wrapper
    /// to the .NET System.IO.Ports.SerialPort class.
    /// </summary>
    public class SerialPort {

        private dynamic serialPort;

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class.
        /// </summary>
        public SerialPort() {
            RegisterCom();
        }
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class using the specified port name.
        /// </summary>
        /// <param name="portName">The port to use (for example, COM1)</param>
        public SerialPort(string portName) : this() {
            PortName = portName;
        }
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class using the specified port name and baud rate.
        /// </summary>
        /// <param name="portName">The port to use (for example, COM1)</param>
        /// <param name="baudRate">The baud rate</param>
        public SerialPort(string portName, int baudRate) : this() {
            PortName = portName;
            BaudRate = baudRate;
        }
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class using the specified port name, baud rate, and parity bit.
        /// </summary>
        /// <param name="portName">The port to use (for example, COM1)</param>
        /// <param name="baudRate">The baud rate</param>
        /// <param name="parity">One of the <see cref="Parity"/> values</param>
        public SerialPort(string portName, int baudRate, Parity parity) : this() {
            PortName = portName;
            BaudRate = baudRate;
            Parity = parity;
        }
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class using the specified port name, baud rate, parity bit, and data bits.
        /// </summary>
        /// <param name="portName">The port to use (for example, COM1)</param>
        /// <param name="baudRate">The baud rate</param>
        /// <param name="parity">One of the <see cref="Parity"/> values</param>
        /// <param name="dataBits">The data bits value</param>
        public SerialPort(string portName, int baudRate, Parity parity, int dataBits) :this() {
            PortName = portName;
            BaudRate = baudRate;
            Parity = parity;
            DataBits = dataBits;
        }
        /// <summary>Initializes a new instance of the <see cref="SerialPort"/> class using the specified port name, baud rate, parity bit, data bits, and stop bit.
        /// </summary>
        /// <param name="portName">The port to use (for example, COM1)</param>
        /// <param name="baudRate">The baud rate</param>
        /// <param name="parity">One of the <see cref="Parity"/> values</param>
        /// <param name="dataBits">The data bits value</param>
        /// <param name="stopBits">One of the <see cref="StopBits"/> values</param>
        public SerialPort(string portName, int baudRate, Parity parity, int dataBits, StopBits stopBits) :this() {
            PortName = portName;
            BaudRate = baudRate;
            Parity = parity;
            DataBits = dataBits;
            StopBits = stopBits;
        }
        #endregion

        /// <summary>Accesses the COM interop object for serial port communication
        /// </summary>
        private void RegisterCom() {
            try {
                serialPort = AutomationFactory.GetObject("InteropComObjects.IO.Ports.SerialPort");
            } catch (Exception) { // not yet running
                try {
                    serialPort = AutomationFactory.CreateObject("InteropComObjects.IO.Ports.SerialPort");
                } catch (Exception) { // not found
                    throw new InteropException();
                }
            }
        }

        #region Events
        /// <summary>Represents the method that will handle the data received event of a <see cref="SerialPort"/> object.
        /// </summary>
        public event SerialDataReceivedEventHandler DataReceived;
        /// <summary>Represents the method that handles the error event of a <see cref="SerialPort"/> object.
        /// </summary>
        public event SerialErrorReceivedEventHandler ErrorReceived;
        /// <summary>Represents the method that will handle the serial pin changed event of a <see cref="SerialPort"/> object.
        /// </summary>
        public event SerialPinChangedEventHandler PinChanged;

        /// <summary>Represents the method that will handle the data received event of a dynamic COM-interoped <see cref="SerialPort"/> object.
        /// </summary>
        private AutomationEvent serialDataReceived;
        /// <summary>Represents the method that handles the error event of a dynamic COM-interoped <see cref="SerialPort"/> object.
        /// </summary>
        private AutomationEvent serialErrorReceived;
        /// <summary>Represents the method that will handle the serial pin changed event of a dynamic COM-interoped <see cref="SerialPort"/> object.
        /// </summary>
        private AutomationEvent serialPinChanged;

        /// <summary>Passes the <see cref="serialDataReceived"/> up through the <see cref="DataReceived"/> event.
        /// </summary>
        /// <param name="sender">The sender of the automation event</param>
        /// <param name="e">The event-args contain the original <see cref="SerialDataReceivedEventArgs"/> as param-array</param>
        private void OnSerialDataReceived(object sender, AutomationEventArgs e) {
            try {
                var args = new SerialDataReceivedEventArgs { EventType = (SerialData)(e.Arguments[1]) };
                DataReceived(this, args);
            } catch (NullReferenceException) {} // no subscriber
        }
        /// <summary>Passes the <see cref="serialErrorReceived"/> up through the <see cref="ErrorReceived"/> event.
        /// </summary>
        /// <param name="sender">The sender of the automation event</param>
        /// <param name="e">The event-args contain the original <see cref="SerialErrorReceivedEventArgs"/> as param-array</param>
        private void OnSerialErrorReceived(object sender, AutomationEventArgs e) {
            try {
                var args = new SerialErrorReceivedEventArgs { EventType = (SerialData)(e.Arguments[1]) };
                ErrorReceived(e.Arguments[0], args);
            } catch (NullReferenceException) { } // no subscriber
        }
        /// <summary>Passes the <see cref="serialPinChanged"/> up through the <see cref="PinChanged"/> event.
        /// </summary>
        /// <param name="sender">The sender of the automation event</param>
        /// <param name="e">The event-args contain the original <see cref="SerialPinChangedEventArgs"/> as param-array</param>
        private void OnSerialPinChanged(object sender, AutomationEventArgs e) {
            try {
                var args = new SerialPinChangedEventArgs { EventType = (SerialData)(e.Arguments[1]) };
                PinChanged(e.Arguments[0], args);
            } catch (NullReferenceException) { } // no subscriber
        }
        #endregion

        /// <summary>Gets an array of serial port names for the current computer.
        /// </summary>
        /// <returns>The names of the currently available serial ports.</returns>
        public string[] GetPortNames() {
            if (null != serialPort) return serialPort.GetPortNames();
            throw new InteropException();
        }
        /// <summary>Opens a new serial port connection.
        /// </summary>
        /// <exception cref="UnauthorizedAccessException">
        /// The exception that is thrown when the operating system denies access
        /// because of an I/O error or a specific type of security error.
        /// </exception>
        /// <exception cref="IOException">The exception that is thrown when an
        /// I/O error occurs</exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// The exception that is thrown when the value of an argument is outside
        /// the allowable range of values as defined by the invoked method.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// The exception that is thrown when one of the arguments provided to a
        /// method is not valid.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// The exception that is thrown when a method call is invalid for the
        /// object's current state.
        /// </exception>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public void Open() {
            if (null != serialPort) {
                serialDataReceived = AutomationFactory.GetEvent(serialPort, "DataReceived");
                serialErrorReceived = AutomationFactory.GetEvent(serialPort, "ErrorReceived");
                serialPinChanged = AutomationFactory.GetEvent(serialPort, "PinChanged");
                serialDataReceived.EventRaised += OnSerialDataReceived;
                serialErrorReceived.EventRaised += OnSerialErrorReceived;
                serialPinChanged.EventRaised += OnSerialPinChanged;
                serialPort.Open();
            }
            else throw new InteropException();
        }
        /// <summary>Closes the port connection, sets the
        /// <see cref="System.IO.Ports.SerialPort.IsOpen"/> Property to false,
        ///  and disposes of the internal <see cref="System.IO.Stream"/> object.
        /// </summary>
        public void Close() {
            if (IsOpen) {
                serialDataReceived.EventRaised -= OnSerialDataReceived;
                serialErrorReceived.EventRaised -= OnSerialErrorReceived;
                serialPinChanged.EventRaised -= OnSerialPinChanged;
                serialDataReceived = null;
                serialErrorReceived = null;
                serialPinChanged = null;
                serialPort.Close();
            }
        }
        /// <summary>Discards data from the serial driver's receive buffer.
        /// </summary>
        public void DiscardInBuffer() {
            if (null != serialPort) serialPort.DiscardInBuffer();
        }
        /// <summary>Discards data from the serial driver's transmit buffer.
        /// </summary>
        public void DiscardOutBuffer() {
            if (null != serialPort) serialPort.DiscardOutBuffer();
        }
        /// <summary>Reads a number of bytes from the <see cref="System.IO.Ports.SerialPort"/>
        /// input buffer and writes those bytes into a byte array at the specified offset.
        /// </summary>
        /// <param name="buffer">The byte array to write the input to</param>
        /// <param name="offset">The offset in the buffer array to begin writing</param>
        /// <param name="count">The number of bytes to read. </param>
        /// <returns>The number of bytes read.</returns>
        /// <exception cref="ArgumentNullException">The buffer passed is null</exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// The offset or count parameters are outside a valid region of the buffer being passed.
        /// Either offset or count is less than zero</exception>
        /// <exception cref="ArgumentException">
        /// Offset plus count is greater than the length of the buffer.
        /// - or -
        /// Count is 1 and there is a surrogate character in the buffer.
        /// </exception>
        /// <exception cref="TimeoutException">No bytes were available to read</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public int Read(byte[] buffer, int offset, int count) {
            if (null != serialPort) {
                try {
                    if (null == buffer)
                        throw new ArgumentNullException("buffer", "The buffer passed is null.");
                    if (count + offset > buffer.Length)
                        throw new ArgumentException("Offset plus count is greater than the length of the buffer.");
                    if (count < 0 || offset < 0) {
                        var pname = count < 0 ? "count" : "offset";
                        throw new ArgumentOutOfRangeException(pname,
                                                              pname.Substring(0, 1).ToUpper() + pname.Substring(1) +
                                                              " is less than zero.");
                    }

                    int bytesRead;
                    for (bytesRead = 0; bytesRead < count; bytesRead++) {
                        buffer[bytesRead + offset] = serialPort.ReadByte();
                    }
                    return bytesRead;
                } catch (InvalidOperationException ie) {
                    if (ie.Message == "No active port.")
                        throw new InteropException();
                    throw; // pass up could be: port not open
                }
            }
            throw new InteropException();
        }
        /// <summary>Reads a number of characters from the <see cref="System.IO.Ports.SerialPort"/>
        ///  input buffer and writes them into an array of characters at a given offset.
        /// </summary>
        /// <param name="buffer">The character array to write the input to</param>
        /// <param name="offset">The offset in the buffer array to begin writing</param>
        /// <param name="count">The number of characters to read</param>
        /// <returns>The number of characters read.</returns>
        /// <exception cref="ArgumentNullException">The buffer passed is null</exception>
        /// <exception cref="ArgumentException">
        /// Offset plus count is greater than the length of the buffer.
        /// - or -
        /// Count is 1 and there is a surrogate character in the buffer.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// The offset or count parameters are outside a valid region of the buffer being passed.
        /// Either offset or count is less than zero</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="TimeoutException">No characters were available to read</exception>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public int Read(char[] buffer, int offset, int count) {
            if (null != serialPort) {
                try {
                    if (null == buffer)
                        throw new ArgumentNullException("buffer", "The buffer passed is null.");
                    if (count + offset > buffer.Length)
                        throw new ArgumentException("Offset plus count is greater than the length of the buffer.");
                    if (count < 0 || offset < 0) {
                        var pname = count < 0 ? "count" : "offset";
                        throw new ArgumentOutOfRangeException(pname,
                                                              pname.Substring(0, 1).ToUpper() + pname.Substring(1) +
                                                              " is less than zero.");
                    }

                    int bytesRead;
                    for (bytesRead = 0; bytesRead < count; bytesRead++) {
                        buffer[bytesRead + offset] = (char)serialPort.ReadChar();
                    }
                    if (count == 1 && Char.IsSurrogate(buffer[offset]))
                        throw new ArgumentException("Count is 1 and there is a surrogate character in the buffer.");
                    
                    return bytesRead;
                } catch (InvalidOperationException ie) {
                    if (ie.Message == "No active port.")
                        throw new InteropException();
                    throw; // pass up could be: port not open
                }
            }
            throw new InteropException();
        }
        /// <summary>Synchronously reads one byte from the <see cref="System.IO.Ports.SerialPort"/>
        /// input buffer.
        /// </summary>
        /// <returns>The byte, cast to an <see cref="Int32"/>, or -1 if the end of the stream has been read.</returns>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
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
            try {
                if (null != serialPort) return serialPort.ReadByte();
            } catch (InvalidOperationException ie) {
                if (ie.Message == "No active port.")
                    throw new InteropException();    
                throw;
            }
            throw new InteropException();
        }
        /// <summary>Synchronously reads one character from the <see cref="System.IO.Ports.SerialPort"/>
        /// input buffer.
        /// </summary>
        /// <returns>The character that was read.</returns>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
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
            try {
                if (null != serialPort) return serialPort.ReadChar();
            } catch (InvalidOperationException ie) {
                if (ie.Message == "No active port.")
                    throw new InteropException();
                throw;
            }
            throw new InteropException();
        }
        /// <summary>Reads all immediately available bytes, based on the encoding, in both the stream and the input buffer of the SerialPort object.
        /// </summary>
        /// <returns>The contents of the stream and the input buffer of the SerialPort object.</returns>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public string ReadExisting() {
            if (null != serialPort) return serialPort.ReadExisting();
            throw new InteropException();
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
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public string ReadLine() {
            try {
                if (null != serialPort) {
                    return serialPort.ReadLine();
                }
            } catch (InvalidOperationException ie) {
                if (ie.Message == "No active port.")
                throw new InteropException();
            }
            throw new InteropException();
        }
        /// <summary>Reads a string up to the specified value in the input buffer.
        /// </summary>
        /// <param name="value">A value that indicates where the read operation stops</param>
        /// <returns>The contents of the input buffer up to the specified value.</returns>
        /// <exception cref="ArgumentException">The length of the value parameter is 0</exception>
        /// <exception cref="ArgumentNullException">The value parameter is null</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open. </exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended</exception>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public string ReadTo(string value) {
            try {
                if (null != serialPort) {
                    return serialPort.ReadTo(value);
                }
            } catch (InvalidOperationException ie) {
                if (ie.Message == "No active port.")
                    throw new InteropException();
            }
            throw new InteropException();
        }
        /// <summary>Writes a specified number of bytes to the serial port using data from a buffer.
        /// </summary>
        /// <param name="buffer">The byte array that contains the data to write to the port</param>
        /// <param name="offset">The zero-based byte offset in the buffer parameter at which to begin copying bytes to the port</param>
        /// <param name="count">The number of bytes to write</param>
        /// <exception cref="ArgumentNullException">The buffer passed is null</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="ArgumentOutOfRangeException">The offset or count parameters are outside a valid region of the buffer being passed. Either offset or count is less than zero</exception>
        /// <exception cref="ArgumentException">offset plus count is greater than the length of the buffer</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended</exception>
        public void Write(byte[] buffer, int offset, int count) {
            if (null != serialPort) serialPort.Write(buffer, offset, count);
            else throw new InteropException();
        }
        /// <summary>Writes a specified number of characters to the serial port using data from a buffer.
        /// </summary>
        /// <param name="buffer">The character array that contains the data to write to the port</param>
        /// <param name="offset">The zero-based byte offset in the buffer parameter at which to begin copying bytes to the port</param>
        /// <param name="count">The number of characters to write</param>
        /// <exception cref="ArgumentNullException">The buffer passed is null</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="ArgumentOutOfRangeException">The offset or count parameters are outside a valid region of the buffer being passed. Either offset or count is less than zero</exception>
        /// <exception cref="ArgumentException">offset plus count is greater than the length of the buffer</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended</exception>
        public void Write(char[] buffer, int offset, int count) {       
            if (null == buffer)
                throw new ArgumentNullException("buffer", "The buffer passed is null.");
            if (count + offset > buffer.Length)
                throw new ArgumentException("Offset plus count is greater than the length of the buffer.");
            if (count < 0 || offset < 0) {
                var pname = count < 0 ? "count" : "offset";
                throw new ArgumentOutOfRangeException(pname,
                                                      pname.Substring(0, 1).ToUpper() + pname.Substring(1) +
                                                      " is less than zero.");
            }
            
            if (null != serialPort) {
                var bBuffer = new byte[count];
                int bCounter = 0;
                foreach (var bufferChar in buffer) {
                    if (offset-- > 0) 
                        continue;
                    if (count-- <= 0)
                        break;
                    bBuffer[bCounter++] = Convert.ToByte(bufferChar);
                }
                serialPort.Write(bBuffer, 0, bBuffer.Length);
            }
            else throw new InteropException();
        }
        /// <summary>Writes the specified string to the serial port.
        /// </summary>
        /// <param name="value">The string for output</param>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="ArgumentNullException">The string passed is null</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended</exception>
        public void Write(string value) {
            var buffer = value.ToCharArray();
            Write(buffer, 0, buffer.Length);
        }
        /// <summary>Writes the specified string and the NewLine value to the output buffer.
        /// </summary>
        /// <param name="value">The string to write to the output buffer</param>
        /// <exception cref="InvalidOperationException">The specified port is not open</exception>
        /// <exception cref="ArgumentNullException">The string passed is null</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended</exception>
        public void WriteLine(string value) {
            Write(value + serialPort.NewLine);
        }
        /// <summary>Not Implemented!
        /// Gets the underlying <see cref="System.IO.Stream"/> object 
        /// for a <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        public Stream BaseStream {
            get { throw new NotImplementedException(); }
        }
        /// <summary>Gets or sets the signal baud rate.
        /// </summary>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public int BaudRate {
            get {
                if (null != serialPort) return serialPort.BaudRate;
                throw new InteropException();
            }
            set {
                if (null != serialPort) serialPort.BaudRate = value;
                else throw new InteropException();
            }
        }
        /// <summary>Gets or sets the break signal state.
        /// </summary>
        public bool BreakState {
            get {
                if (null != serialPort) return serialPort.BreakState;
                return false;
            }
            set {
                if (null != serialPort) serialPort.BreakState = value;
            }
        }
        /// <summary>Gets the number of bytes of data in the receive buffer.
        /// </summary>
        public int BytesToRead {
            get {
                if (serialPort != null) return serialPort.BytesToRead;
                return 0;
            }
        }
        /// <summary>Gets the number of bytes of data in the send buffer.
        /// </summary>
        public int BytesToWrite {
            get {
                if (serialPort != null) return serialPort.BytesToWrite;
                return 0;
            }
        }
        /// <summary>Gets the state of the Carrier Detect Line for the port.
        /// </summary>
        public bool CDHolding {
            get {
                if (serialPort != null) return serialPort.CDHolding;
                return false;
            }
        }
        /// <summary>Gets the state of the Clear-to-Send line.
        /// </summary>
        public bool CtsHolding {
            get {
                if (serialPort != null) return serialPort.CtsHolding;
                return false;
            }
        }
        /// <summary>Gets or sets the standard length of data bits per byte.
        /// </summary>
        public int DataBits {
            get {
                if (null != serialPort) return serialPort.DataBits;
                return 0;
            }
            set {
                if (null != serialPort) serialPort.DataBits = value;
            }
        }
        /// <summary>Gets or sets a value indicating whether null bytes are ignored
        /// when transmitted between the port and the receive buffer.
        /// </summary>
        public bool DiscardNull {
            get {
                if (null != serialPort) return serialPort.DiscardNull;
                return false;
            }
            set {
                if (null != serialPort) serialPort.DiscardNull = value;
            }
        }
        /// <summary>Gets the state of the Data Set Ready (DSR) Signal.
        /// </summary>
        public bool DsrHolding {
            get {
                if (null != serialPort) return serialPort.DsrHolding;
                return false;
            }
        }
        /// <summary>Gets or sets a value that enables the Data Terminal Ready
        /// (DTR) signal during serial communication.
        /// </summary>
        public bool DtrEnable {
            get {
                if (null != serialPort) return serialPort.DtrEnable;
                return false;
            }
            set { if (null != serialPort) serialPort.DtrEnable = value; }
        }
        /// <summary>Not Implemented!
        /// Gets or sets the byte encoding for pre- and post-transmission
        /// conversion of text.
        /// </summary>
        public string Encoding {
            get {
                throw new NotImplementedException("Conversion not supported, yet.");
                if (serialPort != null) return serialPort.Encoding;
            }
            set {
                throw new NotImplementedException("Conversion not supported, yet.");
                if (serialPort != null) serialPort.Encoding = value;
            }
        }
        /// <summary>Gets or sets the handshaking protocol for serial port transmission of data.
        /// </summary>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public Handshake Handshake {
            get {
                if (null != serialPort) return Enum.Parse(typeof(Handshake), serialPort.Handshake.ToString(), true);
                throw new InteropException();
            }
            set {
                if (null != serialPort) serialPort.Handshake = value.ToString();
                else throw new InteropException();
            }
        }
        /// <summary>Gets a value indicating the open or closed status of the
        /// <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        public bool IsOpen {
            get {
                if (null != serialPort) return serialPort.IsOpen;
                return false;
            }
        }
        /// <summary>Gets or sets the value used to interpret the end of a call to the 
        /// <see cref="System.IO.Ports.SerialPort.ReadLine()"/> and 
        /// <see cref="System.IO.Ports.SerialPort.WriteLine(System.String)"/> methods.
        /// </summary>
        public string NewLine {
            get {
                if (null != serialPort) return serialPort.NewLine;
                return null;
            }
            set { if (null != serialPort) serialPort.NewLine = value; }
        }
        /// <summary>Gets or sets the parity-checking protocol.
        /// </summary>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public Parity Parity {
             get {
                 if (null != serialPort) return Enum.Parse(typeof(Parity), serialPort.Parity.ToString(), true);
                 throw new InteropException();
            }
            set {
                if (null != serialPort) serialPort.Parity = value.ToString();
                else throw new InteropException();
            }
        }
        /// <summary>Gets or sets the byte that replaces invalid bytes in a data stream
        /// when a parity error occurs.
        /// </summary>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public byte ParityReplace {
            get {
                if (null != serialPort) return serialPort.ParityReplace;
                throw new InteropException();
            }
            set { if (null != serialPort) serialPort.ParityReplace = value; }
        }
        /// <summary>Gets or sets the port for communications, including but 
        /// not limited to all available COM ports.
        /// </summary>
        public string PortName {
            get { 
                if (serialPort != null) return serialPort.PortName;
                return null;
            }
            set { if (serialPort != null) serialPort.PortName = value; }
        }
        /// <summary>Gets or sets the size of the
        /// <see cref="System.IO.Ports.SerialPort"/> input buffer.
        /// </summary>
        public int ReadBufferSize {
            get {
                if (null != serialPort) return serialPort.ReadBufferSize;
                return 0;
            }
            set { if (null != serialPort) serialPort.ReadBufferSize = value; }
        }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a read operation does not finish.
        /// </summary>
        public int ReadTimeout {
            get {
                if (null != serialPort) return serialPort.ReadTimeout;
                return 0;
            }
            set { if (null != serialPort) serialPort.ReadTimeout = value; }
        }
        /// <summary>Gets or sets the number of bytes in the internal input buffer before a DataReceived event occurs.
        /// </summary>
        public int ReceivedBytesThreshold {
            get {
                if (null != serialPort) return serialPort.ReceivedBytesThreshold;
                return 0;
            }
            set { if (null != serialPort) serialPort.ReceivedBytesThreshold = value; }
        }
        /// <summary>Gest or sets a value indicating wheter the Request to Send
        /// (RTS) signal is enabled during serial communication.
        /// </summary>
        public bool RtsEnable {
            get {
                if (null != serialPort) return serialPort.RtsEnable;
                return false;
            }
            set { if (null != serialPort) serialPort.RtsEnable = value; }
        }
        /// <summary>Gets or sets the standard number of stopbits per byte.
        /// </summary>
        /// <exception cref="InteropException">The COM object was not initialized correctly</exception>
        public StopBits StopBits {
            get {
                if (null != serialPort) return Enum.Parse(typeof(StopBits), serialPort.StopBits.ToString(), true);
                throw new InteropException();
            }
            set {
                if (null != serialPort) serialPort.StopBits = value.ToString();
                else throw new InteropException();
            }
        }
        /// <summary>Gets or sets the size of the serial port output buffer.
        /// </summary>
        public int WriteBufferSize {
            get {
                if (null != serialPort) return serialPort.ReadBufferSize;
                return 0;
            }
            set { if (null != serialPort) serialPort.ReadBufferSize = value; }
        }
        /// <summary>Gets or sets the number of milliseconds before a time-out 
        /// occurs when a write operation does not finish.
        /// </summary>
        public int WriteTimeout {
            get {
                if (null != serialPort) return serialPort.ReadBufferSize;
                return 0;
            }
            set { if (null != serialPort) serialPort.ReadBufferSize = value; }
        }
    }
}