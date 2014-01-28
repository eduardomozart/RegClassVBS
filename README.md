Registry Class for VBScript/MSHTA
=================

Read/Write/Delete/Enumerate Keys and Values on Registry in VBScript/MSHTA using WMI.

Have you ever needed to read/write a 64-bit registry key in VBScript? Have you also ever needed to delete a key with subkeys?

The idea is to:

1. Provide x64 registry key support to VBScript\MSHTA.
2. Check if Key or Value exists on Registry. If Value or Key exists, returns it's type.
3. Delete key, subkeys and values.
4. Enumerate subkeys.
5. Enumarate Values from key (and it's type).
6. Create/Set Value on Registry.
7. Create Key on Registry.
