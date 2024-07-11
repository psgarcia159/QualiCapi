Rem --- Desinstala as bibliotecas
REGSVR32 /u /s Boleto2Net.dll
REGSVR32 /u /s BouncyCastle.Cryptography.dll
REGSVR32 /u /s itextsharp.dll
REGSVR32 /u /s Microsoft.Bcl.AsyncInterfaces.dll
REGSVR32 /u /s NReco.PdfGenerator.dll
REGSVR32 /u /s QRCoder.dll
REGSVR32 /u /s System.Buffers.dll
REGSVR32 /u /s System.Memory.dll
REGSVR32 /u /s System.Numerics.Vectors.dll
REGSVR32 /u /s System.Runtime.CompilerServices.Unsafe.dll
REGSVR32 /u /s System.Text.Encodings.Web.dll
REGSVR32 /u /s System.Text.Json.dll
REGSVR32 /u /s System.Threading.Tasks.Extensions.dll
REGSVR32 /u /s System.ValueTuple.dll

Rem --- Reinstala as bibliotecas
REGSVR32 /s Boleto2Net.dll
REGSVR32 /s BouncyCastle.Cryptography.dll
REGSVR32 /s itextsharp.dll
REGSVR32 /s Microsoft.Bcl.AsyncInterfaces.dll
REGSVR32 /s NReco.PdfGenerator.dll
REGSVR32 /s QRCoder.dll
REGSVR32 /s System.Buffers.dll
REGSVR32 /s System.Memory.dll
REGSVR32 /s System.Numerics.Vectors.dll
REGSVR32 /s System.Runtime.CompilerServices.Unsafe.dll
REGSVR32 /s System.Text.Encodings.Web.dll
REGSVR32 /s System.Text.Json.dll
REGSVR32 /s System.Threading.Tasks.Extensions.dll
REGSVR32 /s System.ValueTuple.dll

Rem --- Registra bibliotecas no GAC (????)
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i Boleto2Net.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i BouncyCastle.Cryptography.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i itextsharp.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i Microsoft.Bcl.AsyncInterfaces.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i NReco.PdfGenerator.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i QRCoder.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Buffers.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Memory.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Numerics.Vectors.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Runtime.CompilerServices.Unsafe.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Text.Encodings.Web.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Text.Json.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.Threading.Tasks.Extensions.dll
"c:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i System.ValueTuple.dll
