REM from https://blog.jayway.com/2014/09/03/creating-self-signed-certificates-with-makecert-exe-for-development/
REM Run this from a VS Command Prompt so pvk2pfx.exe is in the path

makecert.exe ^
-n "CN=CoatsyRootCert,O=coatsy.net,C=Australia" ^
-r ^
-pe ^
-a sha512 ^
-len 4096 ^
-cy authority ^
-sv CoatsyDocSign.pvk ^
CoatsyDocSign.cer

pvk2pfx.exe ^
-pvk CoatsyDocSign.pvk ^
-spc CoatsyDocSign.cer ^
-pfx CoatsyDocSign.pfx ^
-po Pass@word1

pause
