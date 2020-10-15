# VBA + SQL log processor

This was designed for a Windows server where I had no access to any tools other than MS Office. It runs on simple VBA and SQL.

This is a quick hack for automated txt and csv log processing with a couple of manual steps. There is room for improvement in removing manual steps etc. The logs are always in the same format, so the code looks for specific phrases in the file.
Each log represents one batch of 1 configuration and can have up to ~100k rows/lines of text.

You manually select the log output folder from the program this was designed for and the vba runs through all the csv and txt files and gathers information from them.
This saves a lot of time by removing the need to go through each log manually and gathering data for reports. It shows a couple of ways of counting errors in the logs as well as how long the configuration that produced the log took to run.

All the identifying information and data is obfuscated with only example output left.

Running it is just a case of first using ClearDBs(), then  ImportCSV() and ImportLogs() in any order and in the end Counts().
