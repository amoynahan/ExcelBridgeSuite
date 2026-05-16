# ExcelBridge named-pipe smoke tests

This update adds the first step of the external worker / named-pipe path.

Build the solution in `Release | x64`, then load the packed add-in from the publish folder.

In Excel, test:

```excel
=CPP_PIPE_STATUS()
```

```excel
=CPP_PIPE_START()
```

```excel
=CPP_PIPE_PING("hello")
```

Expected response:

```text
PONG    hello
```

Then stop the worker:

```excel
=CPP_PIPE_STOP()
```

Expected response:

```text
STOP    OK
```

What this proves:

- Excel can start an external worker executable.
- The add-in can find the worker beside the XLL/publish output.
- Excel can send a command to the worker over a Windows named pipe.
- The worker can send a response back to Excel.

The worker currently supports only `STATUS`, `PING`, and `STOP`. The next step is to add real matrix/object commands to this same protocol.
