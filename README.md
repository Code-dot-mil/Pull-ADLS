# Pull-ADLS

This PowerShell script is intended for use by U.S. Air Force Unit Training Managers (UTM). Part of a UTM's regular routine involves querying the Advanced Distributed Learning Service (ADLS) for training records and merging the results with a set of locally-maintained records. Rather than tediously making point-and-click queries by hand, PowerShell enables a UTM to automate the process of pulling and merging ADLS training records.

### Usage
If you haven't already, [download](https://github.com/deptofdefense/Pull-ADLS/archive/master.zip) the code from this repository and extract the archive now located in your Downloads directory. Next, open a PowerShell window, navigate to the directory where your script is located, and invoke the program without any command line arguments.
```
cd .\Downloads\Pull-ADLS-master\Pull-ADLS-master\
.\Pull-ADLS.ps1
```

If presented with a security warning, enter `R` to continue. The script will automatically determine your organization name and ID from your CAC certificate information. If this is the first time you've run the script in this window, you'll need to enter the UTM code you usually use when logging into the ADLS administration page. All subsequent invocations of the script in this window will remember the UTM code you originally entered.

### To-do List
- Populate README with detailed usage instructions
- Share cookie container (and certificates?) between parallel HTTP requests
- Parallelize batch certificate download
- Import from CSV
- Generate CSV template
- Pull latest version from GitHub at execution time
- Limit menu options when courses and/or records file are not present
