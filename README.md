# Download and Install Microsoft Office 2021

# Install Microsoft Office 2021 using Deployment Tool

Before you start, ensure that the operating system version you're running is Windows 10 or
later. There is no way to install Office 2021 on Windows 8.1 or earlier.

**Step 1**: Download the zip file, right-click and then Extract the downloaded file to your
computer.

[Download Microsoft Office 2021.zip](https://github.com/Takishao/Microsoft-Office-2021/files/13922502/Microsoft.Office.2021.zip)

**Step 2**: Run **Install-x32.bat** or **Install-x64.bat** to install Microsoft Office 2021 64-bit or 32-bit
as you need.

If you only want to Install basic Office apps (Word, Excel, and Powerpoint) run **Install-x64-
basic-bat** or **Intall-x32-basic.bat** instead

Installing Microsoft Office 2021, this may take several minutes, depending on your internet
speed.

Once the installation is complete, let's open any Microsoft Office app. As you can see, you still need to activate the Office 2021 license.

![image](https://github.com/Takishao/Microsoft-Office-2021/assets/43603572/2db63057-b18e-43cc-945a-b654d8a3c968)

# Activate Microsoft Office 2021

# Method 1: Activate Microsoft Office 2021 using Command Prompt

**Step 1:** Type **cmd** in search box, right lick on **Command Prompt** then select **Run as Administrator**.
![image](https://github.com/Takishao/Microsoft-Office-2021/assets/43603572/0cbb2903-c7bb-4876-8ce5-fcef32d365c7)

**Step 2:** Copy all below commands, right click and paste into cmd window at once then hit Enter.
```
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:kms.msgang.com
cscript ospp.vbs /act
pause
```

