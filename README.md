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

For Windows 2000 SP4, apply Hotfix [KB817478](http://support.microsoft.com/kb/817478) to replace WMI `stdprov.dll` to dump HKEY_USERS properly.

In all cases, WMI must be running for this class to work. On NT systems (2000/XP/Vista), WMI runs as a service. In order for this class to function the WMI service `winmgmts` must be running. Also, DCOM Server Process Launcher `DcomLaunch` must be running.

## WSH Limitations:

 * Cannot get unexpanded REG_EXPAND_SZ value if valuename includes "\".
 * If the key does not contain any explicit valuenames, the program cannot tell apart
   the key's default value from undefined or REG_NONE.
   The program always emits as default value undefined (FLG_ADDREG_KEYONLY).
 * If the key does not contain any explicit valuenames, and the key itself has REG_EXPAND_SZ
   as the default value, and it does not include any expandable string (%value%),
   the program cannot tell its expandability. Program emits the default value as REG_SZ.
 * Windows 2000, 2003 cannot read REG_QWORD values, as it lacks GetQWORDValue() method.
 * Cannot get REG_RESOURCE_LIST(type 8), REG_FULL_RESOURCE_REQUIREMENTS_LIST(type 10) values.
    (you probably do not want them either)
 * Cannot properly get invalid REG_DWORD values having non-4byte length.
 * On Windows 2000, REG_SZ/REG_MULTI_SZ output could have bogus,memory-leak-ish values 
   due to unknown bug in the system.
   (several occurence when dumping the whole HKEY_LOCAL_MACHINE)

 Note:
 * Dumping full tree of HKLM could take significant amount of time
   with a high CPU load. (HKLM dump of Windows Vista yields ~160MB file)
 * By default, refuses to dump >16kB REG_BINARY. Specify "-s bytes#" to change.
 
## Building docs

You can build docs using [Natural Docs](http://www.naturaldocs.org).

## Contributors

- Joe Priestley (2004-2008).
- Kabe. [dump2inf.vbs](http://vega.pgw.jp/~kabe/win/dump2inf.html).
