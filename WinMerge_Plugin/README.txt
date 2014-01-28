
WinMerge 2.x supports COM Plugins such as the VB6 one included in this folder.
I believe the 3.x branch is going to remove this, so you will need a 2.x version.

WinMerge is a free open source diffing utility available from winmerge.org

In order to use this plugin, you will have to:

1) enabled plugins: Plugins -> List -> enable checkbox
2) after the file diff, run it manually Plugins-> Prediffer -> wmIDACompare.dll

The plugin contains two different match engines. You configure it through the main
IDACompare config pane. The Debug UI setting will popup a modal form allowing you
to manually apply and edit the transformations before returning it to Winmerge.
(see screen shot)

IDACompare will auto-install the plugin to the winmegre plugins directory first time you
goto use it from the Tools menu. The installer will register it on the system in the registry.
