# win32-com
Software to access a Windows COM port

The program comProxy aims to simplify I/O through a COM port.
It transmits to the COM port any data it reads from the standard input stream,
and receives data from the COM port and writes it to the standard output stream.
Other software can run the program and use pipes to control its standard input and output streams,
and thus effectively do I/O through the COM port.

comProxy takes the COM port name from a command line argument.
It configures a fixed set of serial port parameters.
