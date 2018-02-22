# Pull-ADLS

This PowerShell script is intended for use by U.S. Air Force Unit Training Managers (UTM). Part of a UTM's regular routine involves querying the Advanced Distributed Learning Service (ADLS) for training records and merging the results with a set of locally-maintained records. Rather than tediously making point-and-click queries by hand, PowerShell enables a UTM to automate the process of pulling and merging ADLS training records.

### To-do List
- Parallelize HTTP requests for each tracked course
- Human-friendly course completion certificate names
- Convert course completion certificate from HTML to PDF/image
- Import from CSV
- Generate CSV template
