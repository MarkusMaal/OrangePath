# Sunlight Beta (PPTOS)
This repository stores macro source code, developer reference and macro-disabled PPTX file. Check "RELEASES" for macro-enabled files.

Each .vba file is equivalent to a module used in this PPTOS. DocModule is slide code (e.g. event calls for ActiveX controls), FormModule is not required to run the PPTOS, however it's used by some development macros to make creating and editing applications easier. StdModule contains code for various macros, which may execute automatically or when interacting with the PPTOS.

Note: VBA files are stored in UTF16-LE encoding. To use this encoding, add the following to your ~/.gitconfig or .git/config file:
```
	[diff "localizablestrings"]
	textconv = "iconv -f utf-16 -t utf-8"
```