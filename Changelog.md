# Vira 2 Change Log

### 2.25
  
Version number 2.24 is skipped for whatever reason I can't remember
  
- Whitelist and blacklist support.
  
  A whitelisted drive will not be copied, even if it doesn't contain any flag file.
  On the contrary, a blacklisted drive will always be copied regardless of whether it has a flag file.
  
- Changed `ExecuteDrive` from running other programs to running VBScript script.
  
### 2.23
  
- Removed `DriveControlProcess()` as it's useless and incomplete.
  
  However, the function of silently executing program from drive is preserved, and it's still working. It'll even report if execution is successful by return value.
  
- Reverse Vira is suspended.
  
- Camouflaging method changed for one more time.
  
- No longer shows the underscore if the drive has no volume name.
  
### 2.22
  
- Now attempts to do "Reverse Vira" by overwriting drive content with prepared ones.
  
  It seems this feature is somehow faulty, though.
  
- Changed default configuration file location.
  
- Reverted double-click behavior to 2.18 (run directly) for convenience concerns.
  
### 2.21
  
Yeah, versions 2.19 and 2.20 are lost. I don't remember how.
  
- Now it refuses to work if double-clicked directly. It will show some random text instead so it is less suspicious now.
  
- Accepts installation parameters via command line now.
  
  It can also destroy a registry key so others won't be able to deselect `Hide protected operating system files` in Windows Explorer settings.
  
  No longer rejects unknown command line arguments.
  
- Not-ready drives will be (temporarily) skipped now so they don't block the whole program.
  
- You can now obtain some Vira settings and information using own drives, silently.
  
  There are also more options available while harvesting silently.
  
- Shows drive utilization in percentage alongside drive size and data size.
  
- Now loads default config if it can't read local config.
  
- `InstallLocal()` now has a silent mode.
  
- The way it camouflages itself is changed again.
  
### 2.18
  
- Added more variables for silent command execution.
  
- Added parameters for silent harvest.
  
- Now shows version info when command line arguments contains `/version`.
  
- Reintroduced human-readable format for `WriteDriveInfo()`.
  
- Update Administrator check implementation so it is faster.
  
### 2.17
  
- It can now silently execute commands from own drives.
  
- Unknown command line arguments are ignored.
  
- Prevent "Drive not ready" error message by adding drive ready check.
  
- Changed default configuration as well as how it camouflages itself.
  
- Cleaned up some unused command line arguments.
  
- It records drive speed while copying files.
  
- Removed human-readable size format
  
- Minor fixes.
  
### 2.16

- Removed almost all customization and migrated them to a configuration file
  The config file location is the only thing available to change
  
- Added `ReadConfig()` and `WriteConfig()`
  
- Moved `WriteDriveInfo()` to a separate Sub
  
### 2.15

- Added command line arguments. Unknown ones are rejected.
  
- Add self-installation (via command line) and uninstallation
  
- While installing, it can now encrypt itself using [Microsoft Script Encoder](https://msdn.microsoft.com/en-us/library/d14c8zsc(v=vs.84).aspx).
  
### 2.14
  
- The program is now modularized. All functions are packed into Sub's
  
### 2.13
  
- Now tries to hide the destination folder on startup.
  
- Some code cleanup.
  
### 2.12
  
- Added an experimental "Harvest" sub-process
  
### 2.11
  
- Now it scans for drives `D:` to `Z:`. Hard drives are automatically filtered out. Only flash drives will be copied.
  
### 2.10
  
- So sad that versions 2.1 ~ 2.9 are lost. No change log available
  
