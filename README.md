# Install-PrettyConsole
Configures Windows 10/11 to use the Scott Hanselman "Ultimate PowerShell Prompt...with Windows Terminal."

https://www.hanselman.com/blog/my-ultimate-powershell-prompt-with-oh-my-posh-and-the-windows-terminal

# Instructions

1. Open an elevated PowerShell console (Run as administrator).
   
   a. When available, use Windows Terminal (WT) with the PowerShell profile, not the Windows PowerShell profile.
   
   b. WT with Windows PowerShell is tested and works. 
   
   c. The Windows PowerShell or PowerShell 7 consoles without WT should also work. Though this is untested...
   
3. Run this command to download and save the script to C:\temp. Or change the path if you want, just adjust the commands accordingly.

```PowerShell
$null = mkdir C:\Temp -Force
curl https://git.io/JExjW -OutFile C:\temp\Install-PrettyConsole.ps1
```

4. Execute the script.

```PowerShell
C:\Temp\Install-PrettyConsole.ps1
```

5. If you ran the script in WT + PowerShell (not Windows PowerShell), all you need to do is restart WT and everything should be done.
6. If you ran the script in Windows PowerShell or PowerShell without WT, you need to run WT as administrator, select the PowerShell profile, and rerun the script (command from step 4) inside the PowerShell tab (not Windows PowerShell (blue icon)).

Once you run WT at this point you should see the base Ultimate Console. Enjoy customizing it!

![image](https://user-images.githubusercontent.com/5922742/131871474-45e1239c-e41a-48d5-9a8a-729598a9d071.png)



# Notes

- You will get these WARNING messages when running in PowerShell 7. This is a known issue that I can't get rid of. Just ignore it.

```PowerShell
WARNING: Module Appx is loaded in Windows PowerShell using WinPSCompatSession remoting session; please note that all input and output of commands from this module will be deserialized objects. If you want to load this module into PowerShell please use 'Import-Module -SkipEditionCheck' syntax.
```

- If the console line doesn't look right in PowerShell or Windows PowerShell, then the font may not have been changed. Go to WT settings, select the profile, Appearance tab, change the font to "CaskaydiaCove NF". Save.


# How to run WT as admin

## Windows 11

- Right-click on the Windows button (or Win+X).
- Select "Windows Terminal (Admin)".
- Yes to the prompt.

![image](https://user-images.githubusercontent.com/5922742/131872152-c5d9e5c2-dd48-4f90-9072-5a2f3f7a7bd5.png)

## Windows 10

- Click the Windows button.
- Search for Windows Terminal.

- Right-click the WT icon.
- Run as administrator

-OR- 

- Click the down arrow on the right side of the menu.

![image](https://user-images.githubusercontent.com/5922742/131872840-860adf2f-53bf-4a38-8628-30623ef9528b.png)

- Run as administrator

![image](https://user-images.githubusercontent.com/5922742/131872909-2f1da96e-f333-47b9-8389-2f09ef708081.png)


# What's the difference between PowerShell and Windows PowerShell?

## Windows PowerShell

This is the legacy PowerShell console that is Windows-only. Development of Windows PowerShell has ended in favor of PowerShell, currently version 7.1.4, which is a cross-platform version of PowerShell. Windows PowerShell still works, and works well. It simply lacks some of the newer language features that are a part of PowerShell 7, and of course it is not cross-platform capable.

The Windows Terminal profile named Windows PowerShell opens this legacy version. The icon for this profile is white on blue.

![image](https://user-images.githubusercontent.com/5922742/131873847-c53ecb72-f8da-4dd8-8155-0c7528bdd2e6.png)


## PowerShell

PowerShell, or PowerShell 7, is the current development branch of PowerShell. It is built on .NET Core which is what allows it to be cross-platform capable. There are some compatibility issues between Windows PowerShell and PowerShell 7 (pwsh) that makes pwsh harder to use with older scripts.

It is recommended, however, that all new script developement should be built on pwsh. Cross pwsh/Windows PowerShell compatible scripting is usually non-trivial.

The icon for PowerShell is white on black.

![image](https://user-images.githubusercontent.com/5922742/131874605-aa7df069-ed04-41f3-98ad-e6588a64a0c3.png)
