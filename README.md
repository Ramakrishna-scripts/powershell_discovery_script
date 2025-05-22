### Python Commands
**IMPORTANT NOTE: MAKE SURE TO ENTER ONLY FOLDER PATHS AS SEARCHPATH AND OUTPUT PATH, NOT FILE PATH** 
1. **Creating a Virtual Environment**  
   To create a virtual environment in Python, navigate to the directory where your Python script is located and run the following command:

   ```bash
   python -m venv pyscript_env
   ```

2. **Activating the Virtual Environment**  
   To activate the virtual environment, use the following command:

   ```bash
   .\pyscript_env\Scripts\activate
   ```

3. **Installing Dependencies**  
   To install the necessary dependencies, execute:

   ```bash
   pip install -r requirements.txt
   ```

4. **Running the Python Script**  
   To run the Python script, use the following command:

   ```bash
   python python_file_discovery.py "D:\a b" "D:\powershell_stuff"
   ```

   - **D:\a b**: Source location to be scanned.
   - **D:\powershell_stuff**: Location where the output will be generated.
   
   > **Note**: Ensure that paths do not end with a `/` or `\`.

---
 
# PowerShell Script Execution Guide

## Running the PowerShell Scripts

To execute the PowerShell scripts, copy and paste the following commands into your PowerShell terminal:

### 1. ** for FileDiscoveryScan_details_script **
Run this command for `FileDiscoveryScan_details_script.ps1`:

```powershell
 
 .\FileDiscoveryScan_details_script.ps1 -ServerName "MyServer" -ServerFolderPathsToScan "D:\a b" -OutputFolder "D:\powershell_stuff"

```

### 2. **Overview Script**
Run this command for `Overview_Script.ps1`:

```powershell
 
.\Overview_Script.ps1 -ServerName "MyServer" -PathsToScan "D:\Data" -OutputDir "C:\MyReports"

```

 

---
 


