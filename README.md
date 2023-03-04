Standalone windows command-line program to analyze ACES xml files

Your system may need the 64bit version of MicroSoft's ACE engine if Access is not installed.

example command line call:

ACESinspectorCLI -i "input\ACES file with spaces.xml" -o myOutputDir -t myTempDir -l myLogsDir -v VCdb20230126.accdb -p PCdb20230126.accdb -q Qdb20230126.accdb --delete --verbose


An analysiss spreadsheet will be produced in the specified output directory. Detailed logging is done (oe log file per run, named to match the input ACES file) will be done to the specified logs directory


Return values (int) from the command call
 - 0 successful analysis. Output spreadsheet and log file written 
 - 1 failure - missing command line args
 - 2 failure - local filesystem problems reading input
 - 3 failure - local filesystem problems writing output
 - 4 failure - reference database (vcdb, pcdb or qdb) not found
 - 5 failure - reference database import (vcdb, pcdb or qdb)
 - 6 failure - xml xsd validation



