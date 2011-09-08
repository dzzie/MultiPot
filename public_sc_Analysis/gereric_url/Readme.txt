
these folders contain raw sample files of shellcode payloads.

the generic_url shellcode handler can handle allof these

it will extract http, ftp, tftp , and ftp echo string commands
and download them as urls.

generic_url will also attempt to locate several styles of 
xor decoder loops in teh shellcode. If found, it will extra
the xor and decode the payload before looking for the url
types listed above.
