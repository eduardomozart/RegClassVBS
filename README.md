# Registry Class for VBScript/MSHTA

Read/Write/Delete/Enumerate Keys and Values on Registry in VBScript/MSHTA using WMI.

Have you ever needed to read/write a 64-bit registry key in VBScript? Have you also ever needed to delete a key with sub keys?

The idea is to:

1. Provide x64 registry key support to VBScript\MSHTA.
1. Check if Key or Value exists on Registry. If Value or Key exists, returns it's type.
1. Delete key, sub keys and values.
1. Enumerate sub keys.
1. Enumerate Values from key (and it's type).
1. Create/Set Value on Registry.
1. Create Key on Registry.

Requirements: Windows 2000 or after (95, 98, NT 4 with WMI Core 1.5).

For Windows 2000, apply Hot-fix KB817478 to dump HKEY_USERS properly.

In all cases, WMI must be running for this class to work. On NT systems (2000/XP/Vista), WMI runs as a service. In order for this class to function the WMI service `winmgmts` must be running. Also, DCOM Server Process Launcher `DcomLaunch` must be running.

## Error codes

Most of the functions return error codes. There are standard error codes - negative numbers - that mean the same for all functions:

* -1 invalid Path
* -2 invalid HKey
* -3 os arch mismatch
* -4 permission denied
* Other error codes specific to the functions.

## Building docs

You can build docs using [VBSdoc](http://www.planetcobalt.net/sdb/vbsdoc.shtml).

```
cscript VBSdoc.vbs /a /i:RegClassVBS.vbs /o:docs
cscript VBSdoc.vbs /a /i:examples\ExportKeys.wsf /o:examples\docs
```

## Contributors

- Joe Priestley (2004-2008).
- Kabe. [dump2inf.vbs](http://vega.pgw.jp/~kabe/win/dump2inf.html).
