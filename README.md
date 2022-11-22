# VBA-imphash
Imphash equivalent for Office files containing VBA macros. 

The Python tool can hash and cluster Office files based on the calculated VBA import hashes. 

The hash is calculated based on the order in which import identifiers are stored in an undocumented Office cache found in Office files with VBA macros. The identifiers are stored following the chronological order in which they were added to the VBA project and, in some cases, the identifiers are kept in the cache even if they are deleted from the VBA project.

The tool or its inner workings could be used for clustering, detection, hunting and potentially forensics.


## How it works
The vba_imphash is computed from "import related" identifiers extracted from the '_VBA_PROJECT' stream. The '_VBA_PROJECT' stream is an [undocumented stream](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ef7087ac-3974-4452-aab2-7dba2214d239) found in Office files containing VBA macros. This is a quote often found in Microsoft Office files documentation and in this case as well: *"PerformanceCache (...) MUST be ignored on read. MUST NOT be present on write."*. 

This stream contains all the identifiers found in the VBA code, such as variable names, function names, APIs and so on. These identifiers seem to be stored in a particular order depending on how the document was generated and, in some cases, seem to persist even if the corresponding VBA code is removed. The '_VBA_PROJECT' stream has been demistified thanks to [Dr. Vesselin Bontchev's](https://bontchev.nlcv.bas.bg/) excellent work in [pcodedmp](https://github.com/bontchev/pcodedmp) from which this project uses some functions to extract said identifiers. 

The tool was named **vba_imphash** as it shares similaritiies to [Mandiant's famous imphash](https://www.mandiant.com/resources/blog/tracking-malware-import-hashing). 
In Mandiant's imphash, the order of imports in the IAT table is influenced by the order of the APIs in the source code and the order in which the corresponding object files are passed to the linker.
The vba_imphash leverages the fact that the order in which identifiers are added to the aforementioned PerformanceCache depends on the order in which the VBA developer adds the identifiers in the VBA code.

From the list of all identifiers extracted from an Office file, the tool only selects those that are "import related". "Import related" identifiers are VBA built-in functions, events, APIs, methods of COM objects commonly used in malicious files and so on. The list of "import related" identifiers is found in the 'import_identifiers.json' file from this project.

The list of identifiers was created based on [Microsoft documentation](https://learn.microsoft.com/en-us/office/vba/api/overview/) and malicious files shared by the open source community.

## Examples
How this works is easier to explain with some examples.

**EXAMPLE 1**

We have the following VBA code:
```
Sub subnr1()
    MyHex = Hex(5)
End Sub

Sub subnr2()
    MyChar = Chr(65)
End Sub

Function f3()
    myOct = Oct(4)
    f3 = myOct
End Function

Sub ShowFileInfo4(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateCreated
    MsgBox s
End Sub

Sub AutoOpen()
    MyNumber = Asc("A")
    MsgBox "test2"
End Sub

```
If a developer copy and pastes this code into a document via the VBA editor offered by the Office suite, from the '_VBA_PROJECT' stream we can extract the following identifiers in this particular order
```
['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'subnr1', 'Document', 'MyHex', 'Hex', 'subnr2', 'MyChar', 'Chr', 'f3', 'myOct', 'Oct', 'ShowFileInfo4', 'filespec', 'fs', 's', 'CreateObject', 'GetFile', 'DateCreated', 'MsgBox', 'AutoOpen', 'MyNumber', 'Asc']
```
Notice that the above identifiers include built in functions, variable names, function names.

Out of these identifiers and based on the predefined list from 'import_identifiers.json', the following identifiers are considered as "import related":
```
['Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project', 'Normal', 'Hex', 'Chr', 'Oct', 'CreateObject', 'GetFile', 'DateCreated', 'MsgBox', 'AutoOpen', 'Asc']
```
In order to compute the vba_imphash the strings from the above list are concantenated via the '-' character and then md5 hashed, obtaining *db5e047c270a87c972dae71ef3fbadc1*.
```
Win16-Win32-Win64-Mac-VBA6-VBA7-Project-Normal-Hex-Chr-Oct-CreateObject-GetFile-DateCreated-MsgBox-AutoOpen-Asc
```

**EXAMPLE 2**
Now let's consider a different scenario. If the order of the functions from the VBA code is scrambled as the following:
```
Sub subnr2()
    MyChar = Chr(65)
End Sub

Sub subnr1()
    MyHex = Hex(5)
End Sub

Sub ShowFileInfo4(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateCreated
    MsgBox s
End Sub

Function f3()
    myOct = Oct(4)
    f3 = myOct
End Function

Sub AutoOpen()
    MyNumber = Asc("A")
    MsgBox "test2"
End Sub

```
If we just copy&paste this code in the VBA editor we get the same identifiers from above, but in a different order and hence a different vba_imphash - *09fb1e28164c8a1745a795697276d840*.

Identifiers:
```
['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'subnr2', 'Document', 'MyChar', 'Chr', 'subnr1', 'MyHex', 'Hex', 'ShowFileInfo4', 'filespec', 'fs', 's', 'CreateObject', 'GetFile', 'DateCreated', 'MsgBox', 'f3', 'myOct', 'Oct', 'AutoOpen', 'MyNumber', 'Asc']
```

"Import identifiers" hashed string:
```
Win16-Win32-Win64-Mac-VBA6-VBA7-Project-Normal-Chr-Hex-CreateObject-GetFile-DateCreated-MsgBox-Oct-AutoOpen-Asc
```

**EXAMPLE 3**

Now for the interesting part, if we keep the same scrambled order as in example 2, but instead of copy&pasting the entire VBA code we add the code sequentially, one function at a time, in the following order: subnr1, subnr2, f3, ShowFileInfo4, AutoOpen, we get the exact same vba_imphash as in example 1, even if the VBA code stored will be identical to example 2.


**EXAMPLE 4**

Perhaps even more interesting is that even deleting part of the VBA macro code only seems to change the order of only a handful of identifiers such as ['Project', 'Normal', 'Office' ,'Document'] that do not affect the "import identifiers" list hence the vba_imphash stays the same.
For this example, one can take the document generated in example 1 delete the 'f3' function, and it will get the exact same vba_imphash as in example 1.

**Other observations**

- Even if other VBA code is added to an already existing document, the order of identifiers already existing is preserved.
- Saving an existing document as a new file leads to the reordering of the identifiers. Saving the file from example 3 as a new document leads to the same vba_imphash as the file from example 2.

## Case study

Samples from [2020-07-20 - WORD DOCS WITH MACROS PUSHING ICEDID (BOKBOT)](https://www.malware-traffic-analysis.net/2020/07/20/index.html)
```
19E84FA825B31A4344D2D992FACD69866001EFD9D7AB9DE2BBA2F8BAD54EFA90       charge_07.20.doc
2446E5FA5A412550FA02B22076D8BAC917D219D027FA867FC60A053133288602       decree.07.20.2020.doc
CC17289F5AF8320473F104037B4E0F431A061FE0D7BA62D6FCE35009CD808017       details 07.20.doc
5F5453BD65E861A7B879AD5C020157CF2473AA35BB90C12A998CD41CD23A8FF2       docs 07.20.20.doc
1FE806AC6B37D4425A40A9FE9E582AB05C9676C8331A2B1F38BC458E966AEAC0       documents-07.20.doc
67F04EFF930119F5A70814D008DD971E8A868E3BCA17756FA07AF254A802D0BA       enjoin.07.20.2020.doc
329064DD4D1880B657600909C4048C97DC4B282A5E9B4CFF06DD117070D708F5       input.07.20.2020.doc
C971C696A37ABE3F8628C3888048096D42873C1E91D778B7DFBA886820058A41       inquiry,07.20.doc
FF9F377F5D6AC1BFCDB7A7B185549CA3761A714CA787456BA656871208EC1B48       legal paper.07.20.doc
D6E51057214555EE71468575F08AFE5128BDA94E6539BBAB8B89FB36121DE10A       official paper-07.20.doc
0250519E459164A9DE842BD9E6BEE64C48C4E8AD6D85E71AED8810D836646332       prescribe -07.20.2020.doc
BFD58B7D1017D89586F7582891DB297B5F75FA867AE282B6ED7329B168895BFA       question,07.20.2020.doc
```

Output from running `vba_imphash.py 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/`
```
Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/charge_07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'Hx', 'autoopen', 'Lr', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'Cd', 'fa', 'g3', 'co', 'Cm', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].
File 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/charge_07.20.doc has the vba imphash 87f2efdb7d55c3e43c83352b2b4e0def from the identifiers ['Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project', 'Normal', 'autoopen', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'Shell', 'URLDownloadToFile', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/decree.07.20.2020.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'D', 'autoopen', 'n', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'Iy', 'jh', 'lq', 'R', 'H', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/details 07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'c6ba3237', 'c11ee1b4', 'bac29bac', 'ActiveWindow', 'DisplayHorizontalScrollBar', 'ea21d096', 'DisplayRulers', 'AutoOpen', 'e538de68', 'd9a57f13', 'aaa', 'ff24bb68', 'dbbb5f26', 'b60d069a', 'c0937b2d', 'e2144357', 'faec0325', 'CreateObject', 'exec', 'd4ed51c9', 'd580001f', 'ce199324', 'f5e2b42b', 'b424a234', 'cc7c5cb6', 'dfe0da05', 'fa6c18a0', 'b3124bf4', 'df22245d', 'b881189a', 'ce6f3d05', 'Selection', 'c9420a37', 'Application', 'ActiveDocument', 'af01ac7d', 'c2ca9be5', 'a06ad8c4', 'Hwnd', 'f309a348', 'Creator', 'b1a8861e', 'c55cbd79', 'ChartDataPointTrack', 'dbd74a61', 'StrConv', 'd3d35dad', 'ec47e177', 'Shapes', 'AlternativeText', 'af64ee14', 'ClickAndTypeParagraphStyle', 'e91f015b', 'Class1', 'd7f29a5a', 'f6bc2a49', 'd597f89b', 'aee0c332', 'f03c3bc8', 'MSXML2', 'XMLHTTP60', 'Send', 'responsebody', 'e7b88c8a', 'StyleAreaWidth', 'df49b5af', 'b5785f4e', 'e4b0d931'].
File 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/details 07.20.doc has the vba imphash e1efaa790b78822931042548c80f16b5 from the identifiers ['Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project', 'Normal', 'ActiveWindow', 'DisplayHorizontalScrollBar', 'DisplayRulers', 'AutoOpen', 'CreateObject', 'exec', 'Selection', 'Application', 'ActiveDocument', 'Hwnd', 'Creator', 'ChartDataPointTrack', 'StrConv', 'Shapes', 'AlternativeText', 'ClickAndTypeParagraphStyle', 'Send', 'responsebody', 'StyleAreaWidth'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/docs 07.20.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'b1bbf9c6', 'a0f97972', 'c2fac190', 'ActiveWindow', 'HorizontalPercentScrolled', 'af174683', 'DocumentMap', 'b97f0223', 'Creator', 'AutoOpen', 'fc644ca1', 'fbd5255e', 'aaa', 'a677e975', 'd88bb23b', 'c2f5932e', 'ba436eea', 'd8d17ded', 'd105f225', 'CreateObject', 'exec', 'c5f5d66e', 'c7b7450c', 'd38b7a0b', 'ad27dc80', 'a8c022cf', 'c2f4a374', 'Application', 'ActiveDocument', 'ConsecutiveHyphensLimit', 'b4ebd87d', 'd5bf097b', 'c35e3263', 'e718db10', 'fa6de948', 'f61e6d84', 'eea221bf', 'ac377da1', 'a4181219', 'View', 'f9feaf2f', 'Visible', 'e27e7537', 'c05d482e', 'a3705320', 'c269e38c', 'AutoSaveOn', 'ba99a350', 'ab6b7b40', 'AttachedTemplate', 'eddea078', 'fcd21d5a', 'c6053011', 'ChartDataPointTrack', 'e40e6765', 'StrConv', 'ed0e8653', 'f578f5c3', 'StyleAreaWidth', 'ab5aea35', 'DisplayScreenTips', 'Shapes', 'AlternativeText', 'e1f640ab', 'Index', 'f6dd9e68', 'd96fb546', 'Class1', 'd9bc0a15', 'DisplayRulers', 'b46c717e', 'fb2c9891', 'e8820aad', 'b8ee83e6', 'c005e001', 'MSXML2', 'XMLHTTP60', 'Send', 'responsebody', 'c9f77080', 'c87feaec', 'Parent', 'bb710c36', 'e353d219', 'af1fdb52'].
File 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/docs 07.20.20.doc has the vba imphash 35ef478dec74fed3ae9746c71eb31ce0 from the identifiers ['Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project', 'Normal', 'ActiveWindow', 'HorizontalPercentScrolled', 'DocumentMap', 'Creator', 'AutoOpen', 'CreateObject', 'exec', 'Application', 'ActiveDocument', 'ConsecutiveHyphensLimit', 'View', 'Visible', 'AutoSaveOn', 'AttachedTemplate', 'ChartDataPointTrack', 'StrConv', 'StyleAreaWidth', 'DisplayScreenTips', 'Shapes', 'AlternativeText', 'Index', 'DisplayRulers', 'Send', 'responsebody', 'Parent'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/\documents-07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'QB', 'autoopen', 'zV', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'bB', 'Ta', 'j8', 'UF', 'gh', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/\enjoin.07.20.2020.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'c6ba3237', 'c11ee1b4', 'b95c28b3', 'ActiveWindow', 'Visible', 'c24a9531', 'Parent', 'AutoOpen', 'da9de17f', 'd9a57f13', 'aaa', 'ff24bb68', 'dbbb5f26', 'aaa1680c', 'c0937b2d', 'e2144357', 'cd52c2ab', 'CreateObject', 'exec', 'd4ed51c9', 'd580001f', 'cd44ba6d', 'DisplayLeftScrollBar', 'f9342700', 'Application', 'ActiveDocument', 'CurrentRsid', 'e393d48f', 'b6ab704d', 'eebd721b', 'fa6c18a0', 'b3124bf4', 'ce31bb28', 'a0b00da4', 'e912391b', 'AutoHyphenation', 'a650ee96', 'DisplayHorizontalScrollBar', 'b06a1bf6', 'ac5c4d59', 'cd9ba679', 'ef4a8c0a', 'WindowState', 'b1a8861e', 'ed95979d', 'ec3fd99a', 'DisplayScreenTips', 'StrConv', 'b6cce54c', 'e08db292', 'Shapes', 'AlternativeText', 'a25c7873', 'd279417a', 'b443b024', 'e6b6341b', 'Class1', 'c5a50d13', 'HorizontalPercentScrolled', 'ab96b2aa', 'ChartDataPointTrack', 'b89ef127', 'fcee294c', 'bb14ae16', 'MSXML2', 'XMLHTTP60', 'Send', 'responsebody', 'c8a3c086', 'Content', 'ee6dd425', 'b5785f4e', 'f2a047cb'].
File 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/enjoin.07.20.2020.doc has the vba imphash 58deb7b6bbd5742ed6f399d2996ced0f from the identifiers ['Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project', 'Normal', 'ActiveWindow', 'Visible', 'Parent', 'AutoOpen', 'CreateObject', 'exec', 'DisplayLeftScrollBar', 'Application', 'ActiveDocument', 'CurrentRsid', 'AutoHyphenation', 'DisplayHorizontalScrollBar', 'WindowState', 'DisplayScreenTips', 'StrConv', 'Shapes', 'AlternativeText', 'HorizontalPercentScrolled', 'ChartDataPointTrack', 'Send', 'responsebody', 'Content'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/input.07.20.2020.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'QB', 'autoopen', 'zV', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'bB', 'Ta', 'j8', 'UF', 'gh', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/inquiry,07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'ha', 'autoopen', 'EE', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'vN', 'kl', 'Ns', 'tF', 'y2', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/legal paper.07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'ND', 'autoopen', 'yR', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'NI', 'dS', 'YN', 'wm', 'hb', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/official paper-07.20.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'AP', 'autoopen', 'au', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'xp', 'Tb', 'ej', 'i0', 'zn', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/prescribe -07.20.2020.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'q1', 'autoopen', 'Si', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'Rx', 'Wt', 'MT', 'yP', 'cE', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].

Computing import hash for OOXML office file 2020-07-20-Word-docs-with-macros-for-IcedID-12-examples/question,07.20.2020.doc.
	[PCODEDMP] All Identifiers = ['Word', 'VBA', 'Win16', 'Win32', 'Win64', 'Mac', 'VBA6', 'VBA7', 'Project1', 'stdole', 'Project', 'ThisDocument', '_Evaluate', 'Normal', 'Office', 'Document', 'Module1', 'nb', 'q2', 'di', 'autoopen', 'RA', 'ActiveDocument', 'CustomXMLParts', 'Count', 'SelectNodes', 'ChildNodes', 'frm', 'download', 'Shell', 'wu', 'bn', 'nb1', 'e4', 'e', 'URLDownloadToFile', 'Wx', 'KZ', 'S6', 'nc', 'X7', 'urlmon', 'UserForm1', 'MSForms', 'UserForm', 'url', 'file'].



****************************************************************************************************
1) Cluster e1efaa790b78822931042548c80f16b5. Len = 1.
Files: ['details 07.20.doc']
2) Cluster 35ef478dec74fed3ae9746c71eb31ce0. Len = 1.
Files: ['docs 07.20.20.doc']
3) Cluster 58deb7b6bbd5742ed6f399d2996ced0f. Len = 1.
Files: ['enjoin.07.20.2020.doc']
4) Cluster 87f2efdb7d55c3e43c83352b2b4e0def. Len = 9.
Files: ['charge_07.20.doc', 'decree.07.20.2020.doc', 'documents-07.20.doc', 'input.07.20.2020.doc', 'inquiry,07.20.doc', 'legal paper.07.20.doc', 'official paper-07.20.doc', 'prescribe -07.20.2020.doc', 'question,07.20.2020.doc']

```

## Setup
Needs Python >= 3.6 as the code uses f-strings. I developed it with Python 3.10. Also uses 7z for extracting OLE stream 'vbaProject.bin' containing macros from [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML) Office files.

**a) Linux**
1) Install 7z
On Ubuntu using apt package manager:
```
sudo apt update
sudo apt install p7zip-full
```
After this you need to be able to run the '7z' command from the terminal.

2) Clone project
```
git clone https://github.com/0x1Avram/vba_imphash.git
```
3) Install dependencies by running this command in the downloaded folder:
```
pip install -r requirements.txt
```
4) Enjoy

**b) Windows**
1) Install 7z

2) Add the installed 7z to Path environment variable

By default 7z gets installed in "C:\Program Files\7-Zip\7z.exe".
After this you need to be able to run the '7z' command from cmd.

3) Clone project
```
git clone https://github.com/0x1Avram/vba_imphash.git
```
4) Install dependencies by running this command in the downloaded folder:
```
pip install -r requirements.txt
```
5) Enjoy


## Usage
Running the 'vba_imphash.py' script without arguments displays detailed information about the command line options:
```
Usage: 
1) For extracting the vba_imphash and displaying the identifiers for a single file: vba_imphash.py file_path
****Example****: vba_imphash.py "details 07.20.doc.old"


2) For clustering files based on the computed vba_imphash:
  a) Without creating the clusters on disk: vba_imphash.py unclustered_files_path
  ****Example****: vba_imphash.py "/home/test/Unclustered files/"

  b) Creating the clusters on disk: vba_imphash.py unclustered_files_path clusters_destination_path
  ****Example****: vba_imphash.py "/home/test/Unclustered files/" "/home/test/Clusters/"


7z needs to be installed and available as a command.
In case the clustering files version of the command line is used, the script creates the following .json files containing relevant information in the current working directory: "vba_imphash_clusters.json", "imphash_identifiers.json", "non_imphash_identifiers.json".
```



## License
The project uses code from [pcodedmp](https://github.com/bontchev/pcodedmp/blob/master/LICENSE) hence the GNU General Public License v3.0.
