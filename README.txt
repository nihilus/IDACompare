
Author:    David Zimmer <dzzie@yahoo.com>
License:   GPL
Copyright: Copyright (C) 2005 iDefense.com, A Verisign Company


  IDACompare_Setup.exe will register dependancies and install full source bundle.

Overview:

 IDACompare is a plugin for IDA which is designed to help you
 line up functions across two separate disassemblies. IDACompare
 also contains a signature scanner, where you can save known functions
 and then scan new disassemblies against them looking for key functions.

 It has tools for sequentially renaming matched functions, as well as porting
 user specified names across disassemblies. 

 This plugin has been designed with Malcode analysis in mind. It should
 work equally well for patch analysis.  

 See readme.chm for more details


Questions:

is there some documentation about the algorithms you used? 
Or can you shortly tell me a bit more about how does it work?

----------------------------

Hello the only documentation on the match logic is within the source code 
itself. It's pretty simple and should be quite readable even to a 
nonprogrammer. The following link will take you directly to one of the 
source lines for some of the match logic

https://github.com/dzzie/IDACompare/blob/master/compare/Form1.frm#L1689

There are two global collections of functions, one for each database. 
Each element is a  class of type Cfunction,

Dim a As New Collection 'of cfunction, all funcs for idb 1
Dim b As New Collection 'of cfunction, all funcs for idb 2
Dim c As CFunction
Dim h As CFunction

Each of these collections is walked over comparing attributes of the 
contained functions trying to find matches

 For Each c In a
    For Each h In b

The Cfunction class is used to parse function attributes and standardize the 
disassembly. Its core is here

https://github.com/dzzie/IDACompare/blob/master/compare/CFunction.cls

The exact CRC method actually works at the standardized disassembler level 
not at the byte level which would not work as offsets change due to recompilation.

There were a lot of modifications 6 to 8 months ago, changes to the C function 
parsing I am not entirely happy with and will likely be reverting. They made 
matching better for close variance, but perform worse in general situations. 
I could also switch between the two based on the results of the exact CRC 
comparison but that may be getting too cute.

The project was originally created in a single weekend, the match logic is relativly 
simple, but it does the brunt of what it needs to do and is easy to modify.

The winmerge plug-in is particularly handy for asm instruction level diffing. 
That came out really well

----------------------------------------------