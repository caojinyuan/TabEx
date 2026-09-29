# Icon Resources

## Application Icon

[generate_icon.py](generate_icon.py) generates the TabEx TE application icon as
[TabExplorer.ico](TabExplorer.ico) beside the script. To regenerate it from the
repository root (requires Pillow):

```powershell
python icons/generate_icon.py
```

The build script uses `icons/TabExplorer.ico` for the EXE icon and bundles this
directory. No copy of the ICO is needed in the repository root.

## Toolbar Icons

SVG assets are unmodified Lucide icons from version 0.468.0:
https://github.com/lucide-icons/lucide/tree/0.468.0/icons

The upstream license is included in LICENSE. These assets are bundled with the
application and require no network access at runtime.

Git Bash, Command Prompt, Windows PowerShell and TortoiseGit use icons extracted
from their locally installed executables when available. Their artwork is not
redistributed here. The SVG files are used as distinct fallbacks if an executable
or its icon is unavailable. TortoiseGit Commit adds a small check badge to the
native icon to distinguish it from TortoiseGit Log.

The pinned-tab marker remains the original red pushpin (U+1F4CC).