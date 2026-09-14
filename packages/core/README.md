# @consultchimps/core

## Random-access byte contracts

`RandomAccessSource` exposes a name, byte size, and bounded asynchronous
`readAt`. `RandomAccessFile` adds writes, truncation, and explicit close. These
contracts let workbook and database operations use filesystem or browser storage
without importing either runtime into the shared operation layer.

## Resource ownership

`OwnedResources` tracks objects with an asynchronous `close()` method. Its
`close()` attempts independent resources together and returns failures with the
owners and original causes. Successful owners are released; failed owners remain
available for another close attempt. Concurrent calls share the current attempt.
Resources added during an attempt remain registered for the next explicit close.

Keep dependent cleanup in the adapter. Closing a database engine and then
removing its file requires two ordered steps; registering both as independent
resources would not preserve that order. The collection does not remove files or
decide whether a failed cleanup changes an operation's result.

See the
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
