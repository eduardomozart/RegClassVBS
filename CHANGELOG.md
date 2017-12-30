# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [3.0.2] - 2017-12-30
### Added
- Set REG_EXPAND_SZ value example on "examples\Create Demo.vbs".
- Explanation on documentation about setting REG_BINARY values.
- Explanation on documentation about "\\" position (at end) for keys on Exists() function.
- Added [Managing Windows Registry with Scripting](http://www.serverwatch.com/tutorials/article.php/1476861) as Class Reference.

### Fixed
- `EnumKeys()` will return -4 (Permission denied) if key exists but cannot be accessed.
- Fixed `EnumKeys()`, `EnumVals()` and `GetValue()` return codes on documentation: they returns a nonzero value instead of a negative one on failure.

### Changed
- `EnumKeys()`, `EnumVals()`, `SetValue()`, `CreateKey()` and `Delete()` functions will return `-3` (Invalid Path) for invalid/non-existent keys/values and "OS arch mismatch" (`-3`) error code becomes `-5` to provides backward compatibility with JSWare CWMIRegClass.

## [3.0.1] - 2017-12-29
### Changed
- Migrate Docs to Natural Docs.
- Return old behavior from `ConvertType()` function to turn the function more friendly and compatible with older scripts that depends from this class. Example: Now it returns "REG_SZ" (old behavior) instead of WMI hex into EnumVals function.

### Fixed
- Fix 64-bit OS support for ExportKeys example.

## [3.0.0] - 2017-12-28
### Added
- Standardized code indentation.
- Documentation.
- Added REG_QWORD support to SetValue function.
- Added Demos.

### Fixed
- Standardized error codes.
- `EnumKeys()` were returning error code `-3` (OS arch mismatch) to `-1` ("Invalid Path") error code.
- Minor bug fixes.

### Changed
- Changed `ConvertType()` logic - Now it returns Type into WMI hex format. To reproduce the old behavior, you can call `ConvertType()` function (that were made public) that automatically converts the type from hex to string and vice-versa. An example can be found on "examples\GetValue-Demo.wsf".
- `Delete()` function will return `0` (Success) instead of `-4` if key/value does not exist.

## 2.5.0 - 2014-02-05
### Added
- Added support to write REG_QWORD values to registry.

### Changed
- If we try to read/write to a 64-bit key on a 32-bit OS, the code will not redirect it automatically to 32-bit (old behavior). Instead, it will send a error code `-3` (os arch mismatch).
- Now the code returns exactly error code instead of -5 if it is an unknown error for EnumKeys.
- Now the code tries to identify REG_EXPAND_SZ and REG_QWORD values if Typ_ is not specified.
- If SetValue can not create the key, the code will not continue. It will return `CreateKey()` error and Exit Function instead.
