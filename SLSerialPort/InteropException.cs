namespace System.IO.Ports {
    public class InteropException : Exception {
        public InteropException() {}
        public InteropException(string message) : base (message) {}
        public InteropException(Exception innerException) : base ("Problem with interop COM. See inner exception for details.", innerException) {}
        public InteropException(string message, Exception innerException) : base (message, innerException) {}
    }
}