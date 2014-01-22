
WinMerge 2.x supports COM Plugins such as the VB6 one included in this folder.
I believe the 3.x branch is going to remove this, so you will need a 2.x version.

WinMerge is a free open source diffing utility available from winmerge.org

In order to use this plugin, you will have to:


1) enabled plugins: Plugins -> List -> enable checkbox
2) after the file diff, run it manually Plugins-> Prediffer -> wmIDACompare.dll

Its a pretty agressive asm standardizer. but so far seems decent but you will need to know
how to read asm to make sure you are looking at the same function or not. 

basically it reduces the diffing to just the instruction (and order of instructions)
ignoring all arguments and jump labels. 

It doesnt have to be this agressive, way more can be done to this and will be if need be.

IDACompare will auto-install the plugin to the winmegre plugins directory first time you
goto use it from the Tools menu. The installer will register it on the system in the registry.
