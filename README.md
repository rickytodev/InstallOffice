# Start manual installation whit `cloning repository`

>[!warning]
>This program is only available for systems with Windows installed.

>[!note]
>Before starting make sure you have [python](https://www.python.org/) and [git](https://git-scm.com/) installed.

1. Clone repository

   ```powershell
   git clone https://github.com/rickytodev/InstallOffice.git
   ```

2. Open route the file

   ```powershell
   cd InstallOffice
   ```

3. Init script

   ```powershell
   py setup.py
   ```

# Start manual installation whit `download .zip`

1. Execute the command in your terminal

   ```powershell
   curl -L -o OfficeLTSC-1.0.20.zip https://github.com/rickytodev/InstallOffice/releases/download/releases/OfficeLTSC-1.0.20.zip
   ```

2. Extract `.zip`

   ```powershell
   Expand-Archive -Path OfficeLTSC-1.0.20.zip -DestinationPath .
   ```

3. Open route the file

   ```powershell
   cd OfficeLTSC
   ```

4. Execute installer

   ```powershell
   ./lnstallOfficeLTSC
   ```
