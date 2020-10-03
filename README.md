# ExcelFormulaExtractor
## Purpose
Helps to extract and reconstruct obfuscated formula code in excel files infected with malware.  This is a script I wrote for a specific piece of malware I was performing static analysis on which can be found at:

https://www.virustotal.com/gui/file/0aae8be1164f7c19c2c7215030b907d1ddefb186b7da77687ddccc4731267474/detection

## What it does
Once you have exported out the excel sheets you want to analyze in CSV format (one per tab) you can run them through the script of which:
* Finds all cells that contain data
* Iterates through each one to reconstruct and expand the referenced cells in in the strings
* Compiles a list of the finalized strings at the end

**Example of an expansion it can perform:**
```
********************************************************
Initial: EsNDtISBBnaCXk=$EI$48049&$U$22638&$EO$25738&$FR$34525&$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
Initial repair: EsNDtISBBnaCXk=$EI$48049&$U$22638&$EO$25738&$FR$34525&$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=S$U$22638&$EO$25738&$FR$34525&$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=Sh$EO$25738&$FR$34525&$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=She$FR$34525&$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=Shel$AP$37942&$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=Shell$B$6338&$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellE$CQ$36635&$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellEx$DS$2454&$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellExe$AW$53450&$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellExec$DA$9220&$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellExecu$FX$55530&$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellExecut$DO$26222&$CV$8941
EsNDtISBBnaCXk=ShellExecute$CV$8941
EsNDtISBBnaCXk=ShellExecuteA
Final: EsNDtISBBnaCXk=ShellExecuteA
```

This ultimately can produce a number of strings and hints as to what the malware is attempting to do to include Win32 calls it makes, directories it creates, executable files it may attempt to write, C&C paths, etc.

Full commentary can be found at:<br /> 
https://www.travismathison.com/posts/Digging-into-obfuscated-excel-formula-code/