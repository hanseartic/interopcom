using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace InteropComObjects.IO.Ports {
    [Guid("77DA781A-5040-48AA-93A2-BC3085FB94F6")]
    public interface ISerialPort {
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
        int ReadByte();
        int ReadChar();
        String ReadExisting();
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
        /// <exception cref="InvalidOperationException">
        /// No connection to a serial-port.
        /// </exception>
        String ReadLine();
        /// <summary>Reads a string up to the specified value in the input buffer.
        /// </summary>
        /// <param name="value">A value that indicates where the read operation stops.</param>
        /// <returns>The contents of the input buffer up to the specified value.</returns>
        /// <exception cref="ArgumentException">The length of the value parameter is 0.</exception>
        /// <exception cref="ArgumentNullException">The value parameter is null.</exception>
        /// <exception cref="InvalidOperationException">The specified port is not open. </exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        /// <exception cref="InvalidOperationException">No connection to a serial-port.</exception>
        String ReadTo(String value);
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
        void Write(byte[] buffer, int offset, int count);
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
        void Write(char[] buffer, int offset, int count);
        /// <summary>Writes the specified string to the serial port.
        /// </summary>
        /// <param name="value">The string for output.</param>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentNullException">The string passed is null.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        void Write(String value);
        /// <summary>Writes the specified string and the NewLine value to the output buffer.
        /// </summary>
        /// <param name="value">The string to write to the output buffer.</param>
        /// <exception cref="InvalidOperationException">The specified port is not open.</exception>
        /// <exception cref="ArgumentNullException">The string passed is null.</exception>
        /// <exception cref="TimeoutException">The operation did not complete before the time-out period ended.</exception>
        void WriteLine(String value);
        #endregion Methods

        #region Properties
        /// <summary>Gets the underlying <see cref="System.IO.Stream"/> object 
        /// for a <see cref="System.IO.Ports.SerialPort"/> object.
        /// </summary>
        object BaseStream { get; }
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
}