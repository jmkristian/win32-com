/** A Win32 console application, which opens a COM port, copies all the bytes
    from stdin to the COM port and copies all the bytes from the COM port to
    stdout. It copies bytes in both directions concurrently. It logs progress
    and errors to stderr. It exits with a non-zero exit code when some errors
    occur.
*/
/* It's tricky to implement this. There's some guidance in "Serial Communications"
   https://learn.microsoft.com/en-us/previous-versions/ff802693(v=msdn.10)?redirectedfrom=MSDN
   (formerly https://msdn.microsoft.com/en-us/library/ff802693.aspx ).
   It refers to an ambitious MTTTY "example", a copy of which was archived in
   https://github.com/bmo/mttty?tab=readme-ov-file

   Here are some things I learned, using Windows 11 with a USB-to-serial
   adapter designed by FTDI (allegedly).

   To simultaneously read and write a COM port, you must do overlapped I/O.
   Non-overlapped I/O by two threads (a reader and a writer) doesn't work.
   Opening the port twice (for reading and again for writing) doesn't work.

   ReadFile from the COM port returns ERROR_SUCCESS and zero bytes read when
   there's nothing in the input buffer. It never returns ERROR_IO_PENDING,
   in my experience. To know when to read, use WaitCommEvent. When EV_RXCHAR
   arrives, you must immediately read all the available input data (otherwise
   some data may be lost). You must not set timeouts for ReadFile. However,
   setting WriteTotalTimeoutConstant > 0 seems to be necessary (or at least
   sufficient) to prevent WaitCommEvent from always completing immediately
   with no event (which wastes CPU time).

   GetOverlappedResult does *not* reset the event in the OVERLAPPED object,
   at least not when reading or writing. You can and should reset the event.
   I tried setting these events to auto-reset, but that seems to be too
   aggressive. It appears WaitForMultipleObjects resets *all* the auto-reset
   objects, but WAIT_OBJECT_0 + N can only indicate one of them. So you can
   miss an event, if two of them occur concurrently.

   Overlapped input from stdin and overlapped output to stdout don't work.
   You can wait for stdin but not stdout. stdin is signaled when there isn't
   any input available. Reading from stdin blocks until some input arrives.

   To handle stdin and stdout, this code starts two threads, one to read from
   stdin and another to write to stdout. They coordinate with the main thread
   via RingBuffer objects. The reader and writer threads block on I/O and
   Events. The main thread calls buffer methods after being alerted via Events.
 */
#include <Windows.h>
#include <fcntl.h>
#include <stdio.h>

static HANDLE comHandle;
// I/O to and from comHandle is asynchronous:
static OVERLAPPED comEventOverlapped = {0};
static OVERLAPPED comRxOverlapped = {0};
static OVERLAPPED comTxOverlapped = {0};
static DWORD comEventMask = 0;
static DWORD comEventError = ERROR_SUCCESS; // from WaitCommEvent(comHandle)
static DWORD comRxError = ERROR_SUCCESS; // from ReadFile(comHandle)
static DWORD comTxError = ERROR_SUCCESS; // from WriteFile(comHandle)
static BOOL comDone = FALSE;

static const int stdinNumber = _fileno(stdin);
static const int stdoutNumber = _fileno(stdout);
static BOOL stdinDone = FALSE;
static BOOL stdoutDone = FALSE;

static const int INFO = 1;
static const int DEBUG = 2;
static const int TRACE = 3;
static const int MAX_MESSAGE = 300;

static int logLevel = TRACE;
static FILE* logFile = NULL;
static int stampTime(char* message) {
    SYSTEMTIME st;
    GetSystemTime(&st);
    return snprintf(message, MAX_MESSAGE,
                    "[%04d-%02d-%02dT%02d:%02d:%02d.%03dZ] ",
                    st.wYear, st.wMonth, st.wDay,
                    st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
}
static void logInfo(const char* format, ...) {
    va_list args;
    va_start(args, format);
    char message[MAX_MESSAGE + 1];
    if (logLevel >= INFO && logFile != NULL) {
        int start = stampTime(message);
        vsnprintf(message + start, MAX_MESSAGE - start, format, args);
        fprintf(logFile, "%s\n", message);
    }
    va_end(args);
}
static void logDebug(const char* format, ...) {
    va_list args;
    va_start(args, format);
    char message[MAX_MESSAGE + 1];
    if (logLevel >= DEBUG && logFile != NULL) {
        int start = stampTime(message);
        vsnprintf(message + start, MAX_MESSAGE - start, format, args);
        fprintf(logFile, "%s\n", message);
    }
    va_end(args);
}
static void logTrace(const char* format, ...) {
    va_list args;
    va_start(args, format);
    char message[MAX_MESSAGE + 1];
    if (logLevel >= TRACE && logFile != NULL) {
        int start = stampTime(message);
        vsnprintf(message + start, MAX_MESSAGE - start, format, args);
        fprintf(logFile, "%s\n", message);
    }
    va_end(args);
}
static LPSTR errorMessage(DWORD errorCode) {
    LPSTR message = NULL;
    FormatMessage
        (FORMAT_MESSAGE_ALLOCATE_BUFFER | 
         FORMAT_MESSAGE_FROM_SYSTEM | 
         FORMAT_MESSAGE_IGNORE_INSERTS,
         NULL,
         errorCode,
         MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
         (LPSTR)&message,
         0, NULL);
    if (message != NULL) {
        // Trim the trailing control characters:
        int last = strlen(message) - 1;
        while (last > 0 && message[last] < ' ') {
            message[last--] = 0;
        }
    }
    return message;
}
static void logError(const char* from, DWORD err) {
    LPTSTR message = NULL;
    switch(err) {
    case ERROR_SUCCESS:
        logTrace("%s success", from);
        break;
    case ERROR_IO_PENDING:
        logDebug("%s pending", from);
        break;
    default:
        message = errorMessage(err);
        logInfo("%s error %d %s", from, err, (message == NULL) ? "" : message);
        if (message != NULL) LocalFree(message);
    }
}
static void logLastError(const char* from) {
    logError(from, GetLastError());
}
static void logIOResult(const char* from, DWORD err, DWORD count) {
    switch(err) {
    case ERROR_SUCCESS:
        logTrace("%s success %d", from, count);
        break;
    case ERROR_IO_PENDING:
        logTrace("%s pending %d", from, count);
        break;
    case ERROR_IO_INCOMPLETE:
        logInfo("%s incomplete %d", from, count);
        break;
    default:
        logError(from, err);
    }
}

static char asStringBuffer[256];
static const char* asString(const BYTE* from, DWORD length) {
    if (length <= 0 || logLevel < DEBUG) return "";
    DWORD end = length;
    if (end >= sizeof(asStringBuffer)) {
        end = sizeof(asStringBuffer) - 1;
    }
    memmove(asStringBuffer, from, end);
    DWORD c;
    for (c = 0; c < end; ++c) {
        if (asStringBuffer[c] < ' ') {
            asStringBuffer[c] = '.';
        }
    }
    asStringBuffer[end] = 0;
    return asStringBuffer;
}

/** Initialize the COM port. */
static int setComm(HANDLE comHandle) {
    DCB comState = {0};
    comState.DCBlength = sizeof(DCB);
    if (!GetCommState(comHandle, &comState)) {
        logLastError("GetCommState");
        return 2;
    }
    comState.BaudRate = CBR_9600;
    comState.ByteSize = 8;
    comState.Parity   = NOPARITY;
    comState.StopBits = ONESTOPBIT;
    comState.fAbortOnError = FALSE;
    comState.fBinary = TRUE;
    comState.fDsrSensitivity = FALSE;
    comState.fDtrControl = DTR_CONTROL_ENABLE;
    comState.fInX = FALSE;
    comState.fOutX = FALSE;
    comState.fOutxCtsFlow = TRUE;
    comState.fOutxDsrFlow = FALSE;
    comState.fRtsControl = RTS_CONTROL_ENABLE;
    if (!SetCommState(comHandle, &comState)) {
        logLastError("SetCommState");
        return 3;
    }
    COMMTIMEOUTS comTimeouts = {0};
    // Timeouts are not used:
    comTimeouts.ReadIntervalTimeout         = MAXDWORD; // read doesn't time out
    comTimeouts.ReadTotalTimeoutConstant    = 0;
    comTimeouts.ReadTotalTimeoutMultiplier  = 0;
    comTimeouts.WriteTotalTimeoutConstant   = 10; // prevents WaitCommEvent from completing prematurely
    comTimeouts.WriteTotalTimeoutMultiplier = 0;
    if (!SetCommTimeouts(comHandle, &comTimeouts)) {
        logLastError("SetCommTimeouts");
        return 4;
    }
    DWORD comMask = EV_RXCHAR | EV_TXEMPTY | EV_CTS | EV_DSR | EV_RLSD | EV_ERR | EV_RING;
    if (!SetCommMask(comHandle, comMask)) {
        logLastError("SetCommMask");
        return 5;
    }
    return 0;
}

/** A queue of bytes, with limited capacity. One reader and one writer may access it concurrently. */
class RingBuffer {
private:
    DWORD bufferSize;
    BYTE* buffer;
    DWORD dataIndex = 0; // index of the first data byte in buffer
    DWORD spaceIndex = 0; // index of the first empty byte in buffer
    CRITICAL_SECTION section;
    DWORD findData() {
        if (spaceIndex >= dataIndex) {
            return spaceIndex - dataIndex;
        } else {
            return bufferSize - dataIndex;
        }
    }
    DWORD findSpace() {
        if (spaceIndex >= dataIndex) {
            return bufferSize - spaceIndex - (dataIndex == 0 ? 1 : 0);
        } else {
            return dataIndex - spaceIndex - 1;
        }
    }
public:
    // Waitable events:
    const HANDLE notFull = CreateEvent(0, TRUE, TRUE, NULL); // hasSpace
    const HANDLE notEmpty = CreateEvent(0, TRUE, FALSE, NULL); // hasData

    RingBuffer(int capacity) {
        InitializeCriticalSection(&section);
        bufferSize = capacity + 1;
        buffer = new BYTE[bufferSize];
    }
    ~RingBuffer() {
        delete[] buffer;
        DeleteCriticalSection(&section);
    }
    BYTE* data() {
        BYTE* result;
        EnterCriticalSection(&section);
        result = &buffer[dataIndex];
        LeaveCriticalSection(&section);
        return result;
    }
    BYTE* space() {
        BYTE* result;
        EnterCriticalSection(&section);
        result = &buffer[spaceIndex];
        LeaveCriticalSection(&section);
        return result;
    }
    DWORD hasData() { // number of bytes available to remove
        DWORD result;
        EnterCriticalSection(&section);
        result = findData();
        LeaveCriticalSection(&section);
        return result;
    }
    DWORD hasSpace() { // number of bytes that can be added
        DWORD result;
        EnterCriticalSection(&section);
        result = findSpace();
        LeaveCriticalSection(&section);
        return result;
    }
    void addData(DWORD count) {
        if (count > 0) {
            DWORD resetError = ERROR_SUCCESS;
            DWORD setError = ERROR_SUCCESS;
            EnterCriticalSection(&section);
            DWORD toAdd = findSpace();
            if (count > toAdd) {
                logInfo("buffer overrun %d > %d", count, toAdd);
            } else {
                toAdd = count;
            }
            spaceIndex += toAdd;
            if (spaceIndex == bufferSize) {
                spaceIndex = 0;
            }
            //logTrace("SetEvent(notEmpty)");
            if (!SetEvent(notEmpty)) {
                setError = GetLastError();
            }
            if (findSpace() <= 0 && !ResetEvent(notFull)) {
                resetError = GetLastError();
            }
            LeaveCriticalSection(&section);
            if (resetError != ERROR_SUCCESS) logError("ResetEvent RingBuffer.notFull", resetError);
            if (setError != ERROR_SUCCESS) logError("SetEvent RingBuffer.notEmpty", setError);
        }
    }
    void removeData(DWORD count) {
        if (count > 0) {
            DWORD resetError = ERROR_SUCCESS;
            DWORD setError = ERROR_SUCCESS;
            EnterCriticalSection(&section);
            DWORD toRemove = findData();
            if (count > toRemove) {
                logInfo("buffer underrun %d > %d", count, toRemove);
            } else {
                toRemove = count;
            }
            dataIndex += toRemove;
            if (dataIndex == bufferSize) {
                dataIndex = 0;
            }
            if (!SetEvent(notFull)) {
                setError = GetLastError();
            }
            if (findData() <= 0 && !ResetEvent(notEmpty)) {
                resetError = GetLastError();
            }
            LeaveCriticalSection(&section);
            if (resetError != ERROR_SUCCESS) logError("ResetEvent RingBuffer.notEmpty", resetError);
            if (setError != ERROR_SUCCESS) logError("SetEvent RingBuffer.notFull", setError);
        }
    }
};

static RingBuffer rxBuffer(128); // bytes moving from the COM port
static RingBuffer txBuffer(128); // bytes moving to the COM port

static DWORD WINAPI stdinReader(LPVOID parameter) {
    while (TRUE) {
        DWORD toRead = txBuffer.hasSpace();
        if (toRead <= 0) {
            WaitForSingleObject(txBuffer.notFull, INFINITE);
            continue;
        }
        DWORD wasRead = _read(stdinNumber, txBuffer.space(), toRead);
        if (wasRead < 0) {
            perror("_read(stdin)");
            stdinDone = TRUE;
            return errno;
        }
        logDebug("stdin read %d %s", wasRead, asString(txBuffer.space(), wasRead));
        txBuffer.addData(wasRead);
        if (wasRead == 0) {
            stdinDone = TRUE;
            return 0;
        }
    }
}

static DWORD WINAPI stdoutWriter(LPVOID parameter) {
    while (TRUE) {
        DWORD toWrite = rxBuffer.hasData();
        if (toWrite <= 0) {
            WaitForSingleObject(rxBuffer.notEmpty, INFINITE);
            continue;
        }
        DWORD wasWritten = _write(stdoutNumber, rxBuffer.data(), toWrite);
        if (wasWritten < 0) {
            perror("_write(stdout)");
            stdoutDone = TRUE;
            return errno;
        }
        logDebug("stdout wrote %d %s", wasWritten, asString(rxBuffer.data(), wasWritten));
        rxBuffer.removeData(wasWritten);
    }
}

/** Continue reading from comHandle. */
static void comRx() {
    /* There are several possible outcomes:
       - Read enough bytes from comHandle to fill up rxBuffer.
       - Read fewer bytes, in which case we wait for more bytes from comHandle.
       - A read operation is PENDING, in which case we wait until it completes.
       In any case, we don't want signals from rxBuffer.notFull;
       reacting to such signals would be a waste of time. So:
     */
    if (!ResetEvent(rxBuffer.notFull)) {
        logLastError("ResetEvent(rxBuffer.notFull)");
    }
    while (!comDone) {
        BOOL justRead = FALSE;
        BYTE* buffer = rxBuffer.space();
        switch (comRxError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            break;
        default:
            DWORD toRead = rxBuffer.hasSpace();
            if (toRead <= 0) return;
            comRxError = ReadFile(comHandle, buffer, toRead, NULL, &comRxOverlapped)
                ? ERROR_SUCCESS : GetLastError();
            logIOResult("comRx ReadFile", comRxError, toRead);
            justRead = TRUE;
        }
        DWORD wasRead = 0;
        switch (comRxError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            if (justRead) return; // comRx() will be called later.
        case ERROR_SUCCESS:
            comRxError = GetOverlappedResult(comHandle, &comRxOverlapped, &wasRead, FALSE)
                ? ERROR_SUCCESS : GetLastError();
            switch (comRxError) {
            case ERROR_IO_INCOMPLETE:
            case ERROR_IO_PENDING:
                return; // comRx() will be called later.
            case ERROR_SUCCESS:
                logDebug("comRx read %d %s", wasRead, asString(buffer, wasRead));
                if (!ResetEvent(comRxOverlapped.hEvent)) {
                    logLastError("comRx ResetEvent");
                }
                if (wasRead <= 0) return;
                rxBuffer.addData(wasRead);
                break;
            default:
                logError("comRx GetOverlappedResult", comRxError);
                comDone = TRUE;
            }
            break;
        default:
            logError("comRx ReadFile", comRxError);
            comDone = TRUE;
        }
    }
}

/** Continue writing to comHandle. */
static void comTx() {
    /* There are several possible outcomes:
       - Write all the bytes from txBuffer to comHandle.
       - Write fewer bytes, in which case we wait until comHandle can accept more.
       - A write operation is PENDING, in which case we wait until it completes.
       In any case, we don't want signals from txBuffer.notEmpty;
       reacting to such signals would be a waste of time. So:
     */
    if (!ResetEvent(txBuffer.notEmpty)) {
        logLastError("ResetEvent(txBuffer.notEmpty)");
    }
    while (!comDone) {
        BOOL justWrote = FALSE;
        BYTE* buffer = txBuffer.data();
        switch (comTxError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            break;
        default:
            DWORD toWrite = txBuffer.hasData();
            if (toWrite <= 0) return;
            comTxError = WriteFile(comHandle, txBuffer.data(), toWrite, NULL, &comTxOverlapped)
                ? ERROR_SUCCESS : GetLastError();
            logIOResult("comTx WriteFile", comTxError, toWrite);
            justWrote = TRUE;
        }
        DWORD wasWritten = 0;
        switch (comTxError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            if (justWrote) return; // comTx() will be called later.
        case ERROR_SUCCESS:
            comTxError = GetOverlappedResult(comHandle, &comTxOverlapped, &wasWritten, FALSE)
                ? ERROR_SUCCESS : GetLastError();
            logIOResult("comTx GetOverlappedResult", comTxError, wasWritten);
            switch (comTxError) {
            case ERROR_IO_INCOMPLETE:
            case ERROR_IO_PENDING:
                return; // comTx() will be called later.
            case ERROR_SUCCESS:
                logDebug("comTx wrote %d %s", wasWritten, asString(buffer, wasWritten));
                if (!ResetEvent(comTxOverlapped.hEvent)) {
                    logLastError("comTx ResetEvent");
                }
                if (wasWritten <= 0) return;
                txBuffer.removeData(wasWritten);
                break;
            default:
                logError("comTx GetOverlappedResult", comTxError);
                comDone = TRUE;
            }
            break;
        default:
            logError("comTx WriteFile", comTxError);
            comDone = TRUE;
        }
    }
}

/** Continue waiting for COM events. */
static void comEvent() {
    while (!comDone) {
        switch (comEventError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            break;
        default:
            comEventError = WaitCommEvent(comHandle, &comEventMask, &comEventOverlapped)
                ? ERROR_SUCCESS : GetLastError();
            logError("comEvent WaitCommEvent", comEventError);
        }
        switch (comEventError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            DWORD dontCare;
            comEventError = GetOverlappedResult(comHandle, &comEventOverlapped, &dontCare, FALSE)
                ? ERROR_SUCCESS : GetLastError();
            break;
        case ERROR_SUCCESS:
            break;
        default:
            logError("comEvent WaitCommEvent", comEventError);
            comDone = TRUE;
            return;
        }
        switch (comEventError) {
        case ERROR_IO_INCOMPLETE:
        case ERROR_IO_PENDING:
            return; // Try again later, when comEventOverlapped.hEvent is signaled.
        case ERROR_SUCCESS:
            if (!ResetEvent(comEventOverlapped.hEvent)) {
                logLastError("comEvent ResetEvent");
            }
            if (logLevel >= TRACE) {
                logTrace("comEvent%s%s%s%s%s%s%s%s%s",
                         (comEventMask & EV_RXCHAR) ? " RXCHAR" : "",
                         (comEventMask & EV_TXEMPTY) ? " TXEMPTY" : "",
                         (comEventMask & EV_CTS) ? " CTS" : "",
                         (comEventMask & EV_DSR) ? " DSR" : "",
                         (comEventMask & EV_RLSD) ? " RLSD" : "",
                         (comEventMask & EV_BREAK) ? " BREAK" : "",
                         (comEventMask & EV_RXFLAG) ? " RXFLAG" : "",
                         (comEventMask & EV_ERR) ? " ERR" : "",
                         (comEventMask & EV_RING) ? " RING" : "");
            }
            if (comEventMask & EV_RXCHAR) {
                comRx();
            }
            if (comEventMask & EV_TXEMPTY) {
                comTx();
            }
            break;
        default:
            logError("comEvent GetOverlappedResult", comEventError);
            comDone = TRUE;
            return;
        }
    }
}

int main(int argc, char** argv) {
    if (argc < 2) {
        logInfo("usage: %s <COM port name> <log file name>\n", argv[0]);
        return 1;
    }
    if (argc > 2) {
        logFile = fopen(argv[2], "w");
        if (logFile == NULL) {
            fprintf(stderr, "fopen(%s) failed\n", argv[2]); 
            return 2;
        }
    } else {
        logFile = stderr;
    }
    int exitCode = 0;
    if (_setmode(stdinNumber, _O_BINARY) == -1) {
        perror("_setmode(stdin, _O_BINARY");
    }
    if (_setmode(stdoutNumber, _O_BINARY) == -1) {
        perror("_setmode(stdout, _O_BINARY");
    }
    char* comPortName = argv[1];
    logDebug("CreateFile(%s)", comPortName);
    comHandle = CreateFile(comPortName,
                           GENERIC_READ | GENERIC_WRITE,
                           0, // not shared
                           NULL, // no security
                           OPEN_EXISTING,
                           FILE_FLAG_OVERLAPPED,
                           NULL); // template file
    if (comHandle == INVALID_HANDLE_VALUE) {
        DWORD err = GetLastError();
        LPTSTR message = errorMessage(err);
        logInfo("CreateFile(%s) error %d %s", comPortName, err, (message == NULL) ? "" : message);
        if (message != NULL) LocalFree(message);
        return 3;
    }
    setComm(comHandle);
    comEventOverlapped.hEvent = CreateEvent(NULL, TRUE, TRUE, NULL);
    comRxOverlapped.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    comTxOverlapped.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

    HANDLE stdinReaderThread = CreateThread(NULL, 2048, stdinReader, NULL, 0, NULL);
    HANDLE stdoutWriterThread = CreateThread(NULL, 2048, stdoutWriter, NULL, 0, NULL);

    HANDLE waitables[] = {
        comEventOverlapped.hEvent,
        comRxOverlapped.hEvent,
        comTxOverlapped.hEvent,
        rxBuffer.notFull,
        txBuffer.notEmpty,
    };
    while (TRUE) {
        /*  This is unnecessary:
        if (txBuffer.hasData() && comTxError == ERROR_SUCCESS) {
            comTx();
        }
        ... because comTx() will be called when txBuffer.notEmpty is signaled.

        This would frequently be a waste of time:
        if (rxBuffer.hasSpace() && comRxError == ERROR_SUCCESS) {
            comRx();
        }
        ... because it would only read zero bytes from the COM port.
        The COM port indicates no input by returning ERROR_SUCCESS
        and zero bytes read. To avoid this waste of time, we call
        comRx() when WaitCommEvent returns EV_RXCHAR.
        */
        if ((stdoutDone || !rxBuffer.hasData())
            && (comDone || (stdinDone && !txBuffer.hasData()))) {
            break; // exit gracefully
        }
        DWORD waited = WaitForMultipleObjects(5, waitables, FALSE, 2000);
        switch (waited) {
        case WAIT_OBJECT_0 + 0: // COM event
            comEvent();
            continue;
        case WAIT_OBJECT_0 + 3: // rxBuffer
            if (comRxError != ERROR_SUCCESS) continue; // a Read operation is PENDING
        case WAIT_OBJECT_0 + 1: // rx
            comRx();
            continue;
        case WAIT_OBJECT_0 + 4: // txBuffer
            if (comTxError != ERROR_SUCCESS) continue; // a Write operation is PENDING
        case WAIT_OBJECT_0 + 2: // tx
            comTx();
            continue;
        case WAIT_TIMEOUT:
            logTrace("WAIT_TIMEOUT");
            /* A Read may complete immediately with zero bytes read, and
               a Write may complete immediately with zero bytes written.
               Repeating the operation will complete immediately again.
               WaitCommEvent might indicate when to retry, but NOT always.
               So retry periodically:
             */
            if (rxBuffer.hasSpace() && comRxError == ERROR_SUCCESS) {
                logTrace("comRx retry");
                comRx();
            }
            if (txBuffer.hasData() && comTxError == ERROR_SUCCESS) {
                logTrace("comTx retry");
                comTx();
            }
            continue;
        case WAIT_FAILED:
            logLastError("WAIT_FAILED");
            exitCode = 4;
            break;
        default:
            logInfo("WaitForMultipleObjects %x", waited);
            exitCode = 5;
        }
        break; // loop
    }
    if (exitCode == 0 && comDone) exitCode = 6;
    logInfo("Exit code %d %s%stxData %d rxData %d",
            exitCode,
            (comDone ? "comDone " : ""),
            (stdinDone ? "stdinDone " : ""),
            txBuffer.hasData(),
            rxBuffer.hasData());
    CloseHandle(comHandle);
    fclose(logFile);
    return exitCode;
}
