# Remove-ASCVDuplicates

This is my first PS script for the use of others. It's also my first utilizing Windows Forms.

This script has the user select yesterday's list of samples, then today's list of samples. 

From yesterday's list, it keeps all samples 6 days old (yesterday's day 5) in a "previous" index. 
From today's list, it keeps all samples 5 and 6 days old.
Then, it compares the previous index to today's index and keeps any differences, as well as anything 5 days old today. This catches any samples 6 days old that did not appear on yesterday's list.

As this was my first script for other users, it necessitated tight version control and creation of a log file.
