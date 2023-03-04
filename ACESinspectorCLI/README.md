Standalone windows command-line program to analyze ACES xml files

Your system may need the 64bit version of MicroSoft's ACE engine if Access is not installed.

example command line call:

ACESinspectorCLI -i "input\ACES file with spaces.xml" -o myOutputDir -t myTempDir -l myLogsDir -v VCdb20230126.accdb -p PCdb20230126.accdb -q Qdb20230126.accdb --delete --verbose


Return Vaules (numeric)

 - 0 successful analysis. Output spreadsheet and log file written 
 - 1 failure - missing command line args
 - 2 failure - local filesystem problems reading input
 - 3 failure - local filesystem problems writing output
 - 4 failure - reference database (vcdb, pcdb or qdb) not found
 - 5 failure - reference database import (vcdb, pcdb or qdb)
 - 6 failure - xml xsd validation

