/*
 * Changes
 * 1.0.0.17 (10/6/2023) updated internal ACES schema 4.2 to the latest version of 4.2
 * 1.0.0.16 (7/31/2023) added support for brand/subbrand attributes (at app level) to resolve overlaps
 * 1.0.0.15 (3/5/2023)  added modern (1990+) basevid usage calculation
 * 1.0.0.14 (3/4/2023)  added basevid and qdb usage tabs to spreadsheet
 * 1.0.0.13 (3/4/2023)  defined return values of command call. Refined log output
 * 1.0.0.12 (3/4/2023)  added baseVehicleIDs tab to spreadsheet output
 * 1.0.0.11 (3/3/2023)  changed logging to a dedicated .log file in a logs directory for the given input file
 * 1.0.0.8  (3/1/2023)  fixed logfile functionality
 * 1.0.0.7  (3/1/2023)  fixed logfile functionality when none specified, added error counts to the log output
 * 1.0.0.6  (3/1/2023)  added logfile functionality
 * 1.0.0.5  (3/1/2023)  Round runtime to .1 seconds and display on console output
 * 1.0.0.4  (3/1/2023)  Timespan runtime calulation
 * 1.0.0.3  (2/28/2023) Defaulted to non-verbose console output and no delete of the input ACEC file on completion. Brandcode added to output
 * 1.0.0.0  (2/27/2023) forked code from main ACESinspector (GUI) project at version 1.2.0.48
 * 
 * 
 * 
 * return values
 * 0 successful analysis. Output spreadsheet and log file written 
 * 1 failure - missing command line args
 * 2 failure - local filesystem problems reading input
 * 3 failure - local filesystem problems writing output
 * 4 failure - reference database (vcdb, pcdb or qdb) not found
 * 5 failure - reference database import (vcdb, pcdb or qdb)
 * 6 failure - xml xsd validation
 */


using System.Security.Cryptography;

namespace ACESinspectorCLI
{
    class program
    {
        static string escapeXMLspecialChars(string inputString)
        {
            string outputString = inputString;
            if (!string.IsNullOrEmpty(outputString))
            {
                outputString = outputString.Replace("&", "&amp;");
                outputString = outputString.Replace("<", "&lt;");
                outputString = outputString.Replace(">", "&gt;");
                outputString = outputString.Replace("'", "&apos;");
                outputString = outputString.Replace("\"", "&quot;");
            }
            return outputString;
        }


        static int Main(string[] args)
        {
            DateTime startingDateTime = DateTime.Now;

            if (args.Length == 0)
            {
                Console.WriteLine("Version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
                Console.WriteLine("usage: ACESinspectorCLI -i <ACES xml file> -v <VCdb access file> -p <PCdb access file> -q <Qdb access file> -o <assessment file> -t <temp directory> [-l <logs directory>]");
                Console.WriteLine("\r\n optional switches");
                Console.WriteLine("  --verbose    verbose console output");
                Console.WriteLine("  --delete     delete input ACES file uppon successful analysis");
                return 1; //failure - missing command line args
            }

            bool verbose = false;
            bool deleteACESfileOnSuccess = false;

            string inputFile = "";
            string VCdbFile = "";
            string PCdbFile = "";
            string QdbFile = "";
            string logFile = "";
            string assessmentsPath = "";
            string cachePath = "";

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--verbose") { verbose = true; }
                if (args[i] == "--delete") { deleteACESfileOnSuccess = true; }
                if (args[i] == "-i" && i < (args.Length-1)) {inputFile = args[i + 1]; }
                if (args[i] == "-o" && i < (args.Length - 1)) { assessmentsPath = args[i + 1]; }
                if (args[i] == "-t" && i < (args.Length - 1)) { cachePath = args[i + 1]; }
                if (args[i] == "-v" && i < (args.Length - 1)) { VCdbFile = args[i + 1]; }
                if (args[i] == "-p" && i < (args.Length - 1)) { PCdbFile = args[i + 1]; }
                if (args[i] == "-q" && i < (args.Length - 1)) { QdbFile = args[i + 1]; }
                if (args[i] == "-l" && i < (args.Length - 1)) { logFile = args[i + 1]; }
            }


            if (!File.Exists(inputFile))
            {
                Console.WriteLine("input ACES file ("+ inputFile + ") does not exist");
                return 2; //failure - local filesystem problems reading input
            }

            if (!Directory.Exists(assessmentsPath))
            {
                Console.WriteLine("output directory (" + assessmentsPath + ") does not exist");
                return 3; //failure - local filesystem problems writing output
            }

            if (Directory.Exists(cachePath))
            {// temp directory exists - now see if the AiFragments folder exists and create it if not
                if (!Directory.Exists(cachePath+"\\AiFragments"))
                {
                    try 
                    {
                        Directory.CreateDirectory(cachePath + "\\AiFragments");
                    }
                    catch (Exception ex)
                    { 
                        Console.WriteLine("failed to create AiFragments directory inside temp folder: " + ex.Message); 
                        return 3; //failure - local filesystem problems writing output
                    }
                }
            }
            else 
            {
                Console.WriteLine("temp directory (" + cachePath + ") does not exist");
                return 3; //failure - local filesystem problems writing output
            }

            if (!File.Exists(VCdbFile))
            {
                Console.WriteLine("VCdb Access database file (" + VCdbFile + ") does not exist");
                return 4; //failure - reference database (vcdb, pcdb, qdb) not found
            }

            if (!File.Exists(PCdbFile))
            {
                Console.WriteLine("PCdb Access database file (" + PCdbFile + ") does not exist");
                return 4; //failure - reference database (vcdb, pcdb, qdb) not found
            }

            if (!File.Exists(QdbFile))
            {
                Console.WriteLine("Qdb Access database file (" + QdbFile + ") does not exist");
                return 4; //failure - reference database (vcdb, pcdb, qdb) not found
            }


            if (verbose) { Console.WriteLine("version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()); }

            List<string> cacheFilesToDeleteOnExit = new List<string>();

            bool useAssetsAsFitment = false;
            bool reportAllAppsInProblemGroup = false;
            bool concernForDisparate = false;
            bool respectQdbType = false;
            string macroProblemsDescription = "";
            int threadCount = 20;
            int treeConfigLimit = 1000;

            Dictionary<String, String> noteTranslationDictionary = new Dictionary<string, string>();
            Dictionary<String, QdbQualifier> noteToQdbTransformDictionary = new Dictionary<string, QdbQualifier>();

            ACES aces = new ACES();     // this instance will hold all the data imported from our "primary" ACES xml file
            VCdb vcdb = new VCdb();     // this class will hold all the contents of the the imported VCdb M$Access file - mostly in "Dictionary" type variables for super-fast lookup (way faster than repeatedly querying the underlying Access file)
            PCdb pcdb = new PCdb();
            Qdb qdb = new Qdb();

            // hash the input file - temp fragment files are named including this hash which ensures a fragment that wasn't deleted from a past run can't confuse assessment results
            using (var md5 = MD5.Create())
            {
                try
                {
                    using (var stream = File.OpenRead(inputFile))
                    {
                        aces.fileMD5hash = BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", string.Empty);
                    }
                }
                catch (Exception ex)
                {
                    if (verbose) { Console.WriteLine("error opening input ACES file: " + ex.Message); }
                    return 2; //failure - local filesystem problems reading input
                }
            }


            if (logFile != "")
            {// logs directory was specd. add the input aces filename to it with a .log ext
                logFile = logFile + "\\" + Path.GetFileNameWithoutExtension(inputFile)+"_"+ aces.fileMD5hash + ".log";
                try
                { // write the file (potentially over-write) with one line of content to get the party started
                    File.WriteAllText(logFile, DateTime.Now.ToString() + "\tVersion " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + " started" + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error writing to log file: " + ex.Message);
                    return 3; //failure - local filesystem problems writing output
                }
            }


            if (verbose) { Console.WriteLine("ACES file md5 hash: " + aces.fileMD5hash); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tACES file:" + inputFile + Environment.NewLine); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tACES file md5 hash:" + aces.fileMD5hash + Environment.NewLine); }

            // load reference databases

            if (verbose) { Console.WriteLine("Importing VCdb: " + VCdbFile); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tImporting VCdb: " + VCdbFile + Environment.NewLine); }

            vcdb.clear();
            vcdb.connectLocalOLEDB(VCdbFile);
            vcdb.importOLEDBdata();
            if (!vcdb.importSuccess)
            {
                if (verbose) { Console.WriteLine("  VCdb import failed: " + vcdb.importExceptionMessage); }
                if (logFile != ""){ File.AppendAllText(logFile, DateTime.Now.ToString() + "\tVCdb import failed: " + vcdb.importExceptionMessage + Environment.NewLine); }
                return 5; //failure - reference database import (vcdb, pcdb, qdb) 
            }
            if (verbose) { Console.WriteLine("  Done (version date: " + vcdb.version + ")"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tSuccessful VCdb import (version date: " + vcdb.version + ")" + Environment.NewLine); }


            if (verbose) { Console.WriteLine("Importing PCdb: " + PCdbFile); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tImporting PCdb: " + PCdbFile + Environment.NewLine); }
            pcdb.clear();
            pcdb.connectLocalOLEDB(PCdbFile);
            pcdb.importOLEdb();
            if (!pcdb.importSuccess)
            {
                if (verbose) { Console.WriteLine("PCdb import failed: " + pcdb.importExceptionMessage); }
                return 5; //failure - reference database import (vcdb, pcdb, qdb) 
            }
            if (verbose) { Console.WriteLine("  Done PCdb (version date: " + pcdb.version + ")"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tSuccessful PCdb import (version date: " + pcdb.version + ")" + Environment.NewLine); }


            if (verbose) { Console.WriteLine("Importing Qdb: " + QdbFile); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tImporting Qdb: " + QdbFile + Environment.NewLine); }
            qdb.clear();
            qdb.connectLocalOLEDB(QdbFile);
            qdb.importOLEdb();
            if (!qdb.importSuccess)
            {
                if (verbose) { Console.WriteLine("  Qdb import failed: " + qdb.importExceptionMessage); }
                if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tQdb import failed: " + qdb.importExceptionMessage + Environment.NewLine); }
                return 5; //failure - reference database import (vcdb, pcdb, qdb) 
            }
            if (verbose) { Console.WriteLine("  Done Qdb (version date: " + qdb.version + ")"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tSuccessful Qdb import (version date: " + qdb.version + ")" + Environment.NewLine); }

            aces.clear();

            if (verbose) { Console.WriteLine("Importing ACES xml"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tStarting ACES file import:" + aces.filePath + Environment.NewLine); }

            int importedAppsCount = aces.importXML(inputFile, "", false, false, noteTranslationDictionary, noteToQdbTransformDictionary, vcdb);

            if (importedAppsCount > 0)
            {
                if (verbose) { Console.WriteLine("  Done (" + importedAppsCount.ToString() + " apps)"); }
                if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tImported " + importedAppsCount.ToString() + " apps" + Environment.NewLine); }
            }
            else
            { // something went wrong in the import

                if (verbose) { Console.WriteLine("Import of XML file failed"); }
                if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tACES file import failed. First XSD error:" + aces.xmlValidationErrors.FirstOrDefault().ToString() + Environment.NewLine); }
                return 6; //failure - xml xsd validation
            }


            aces.analysisRunning = true;

            string problemDescription = "";
            int elementPrevalence = 0;
            Dictionary<string, int> fitmentElementPrevalence = new Dictionary<string, int>();


            List<string> partTypeNameList = new List<string>();
            List<string> positionNameList = new List<string>();
            foreach (KeyValuePair<string, int> partEntry in aces.partsAppCounts)
            {
                partTypeNameList.Clear(); foreach (int partTypeId in aces.partsPartTypes[partEntry.Key]) { partTypeNameList.Add(pcdb.niceParttype(partTypeId)); }
                positionNameList.Clear(); foreach (int positionId in aces.partsPositions[partEntry.Key]) { positionNameList.Add(pcdb.nicePosition(positionId)); }
            }


            aces.clearAnalysisResults();
            var taskList = new List<Task>();


            aces.qtyOutlierThreshold = 1;
            aces.qtyOutlierSampleSize = 1000;

            aces.establishFitmentTreeRoots(useAssetsAsFitment); // maybe should move this to be the last step in the import process. It is not threadable, so it must be run in a blocking way before the threadable stuff is run

            int numberOfSections = threadCount;
            if ((numberOfSections * 5) > aces.apps.Count()) { numberOfSections = 1; } // ensure at least 5apps per sections or dont break up

            int sectionSize = Convert.ToInt32(aces.apps.Count() / numberOfSections);
            int sectionNumber = 1;

            aces.individualAnanlysisChunksList.Add(new analysisChunk());
            aces.individualAnanlysisChunksList.Last().appsList = new List<App>();
            aces.individualAnanlysisChunksList.Last().id = 1; aces.individualAnanlysisChunksList.Last().cachefile = cachePath + "\\AiFragments\\" + aces.fileMD5hash;

            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_parttypePositionErrors1.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_QdbErrors1.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_questionableNotes1.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidBasevehicles1.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidVCdbCodes1.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_configurationErrors1.txt");

            for (int i = 0; i < aces.apps.Count(); i++)
            {
                if (aces.individualAnanlysisChunksList.Last().appsList.Count() >= sectionSize)
                {
                    sectionNumber++;
                    aces.individualAnanlysisChunksList.Add(new analysisChunk());
                    aces.individualAnanlysisChunksList.Last().id = sectionNumber;
                    aces.individualAnanlysisChunksList.Last().cachefile = cachePath + "\\AiFragments\\" + aces.fileMD5hash;
                    aces.individualAnanlysisChunksList.Last().appsList = new List<App>();
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_parttypePositionErrors" + sectionNumber.ToString() + ".txt");
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_QdbErrors" + sectionNumber.ToString() + ".txt");
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_questionableNotes" + sectionNumber.ToString() + ".txt");
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidBasevehicles" + sectionNumber.ToString() + ".txt");
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidVCdbCodes" + sectionNumber.ToString() + ".txt");
                    cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_configurationErrors" + sectionNumber.ToString() + ".txt");
                }
                aces.individualAnanlysisChunksList.Last().appsList.Add(aces.apps[i]);
            }

            // run in a parallel (an arbitrary number a chunks) the individaul analysis of the primary ACES apps list

            //if (vcdb.useRemoteDB) { foreach (analysisChunk chunk in aces.individualAnanlysisChunksList) { vcdb.addNewMySQLconnection(); } } // instance one mysql connection for each chunk
            var individialAppAnalysisTask = new Task(() => { Parallel.ForEach(aces.individualAnanlysisChunksList, chunk => { aces.findIndividualAppErrors(chunk, vcdb, pcdb, qdb); }); });
            taskList.Add(individialAppAnalysisTask);
            individialAppAnalysisTask.Start();

            //run a single, sequential thread looking for outlier-ness in the apps list 
            // this can be run concurently with the other analysis threads, it just can't be broken into its own concurent chunks
            aces.outlierAnanlysisChunksList.Add(new analysisChunk());
            aces.outlierAnanlysisChunksList.Last().cachefile = cachePath + "\\AiFragments\\" + aces.fileMD5hash;
            aces.outlierAnanlysisChunksList.Last().appsList = aces.apps;
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_qtyOutliers.txt");
            var outlierAppAnalysisTask = new Task(() => { aces.findIndividualAppOutliers(aces.outlierAnanlysisChunksList.Last(), vcdb, pcdb, qdb); });
            taskList.Add(outlierAppAnalysisTask);
            outlierAppAnalysisTask.Start();

            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_parttypeDisagreements.txt");
            cacheFilesToDeleteOnExit.Add(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_assetProblems.txt");


            numberOfSections = threadCount;
            if ((numberOfSections * 5) > aces.fitmentAnalysisChunksList.Count()) { numberOfSections = 1; } // ensure at least 5 chunkgroups per sections or dont break up

            sectionSize = Convert.ToInt32(aces.fitmentAnalysisChunksList.Count() / numberOfSections);


            sectionNumber = 1;
            aces.fitmentAnalysisChunksGroups.Add(new analysisChunkGroup());
            aces.fitmentAnalysisChunksGroups.Last().chunkList = new List<analysisChunk>();
            aces.fitmentAnalysisChunksGroups.Last().id = sectionNumber;

            for (int i = 0; i < aces.fitmentAnalysisChunksList.Count(); i++)
            {// every chunk represents the apps in one fitment tree (MMY/parttype/position/mfrlabel/asset)
                // chunkgroups represent an arbitrary-sized collection of chunks for the purpose for parrallel multi-tasking (each group is a new task)
                if (aces.fitmentAnalysisChunksGroups.Last().chunkList.Count() >= sectionSize)
                {
                    sectionNumber++;
                    aces.fitmentAnalysisChunksGroups.Add(new analysisChunkGroup());
                    aces.fitmentAnalysisChunksGroups.Last().chunkList = new List<analysisChunk>();
                    aces.fitmentAnalysisChunksGroups.Last().id = sectionNumber;
                }
                aces.fitmentAnalysisChunksGroups.Last().chunkList.Add(aces.fitmentAnalysisChunksList[i]);
            }

            var macroAppAnalysisTask = new Task(() => { Parallel.ForEach(aces.fitmentAnalysisChunksGroups, chunkGroup => { aces.findFitmentLogicProblems(chunkGroup, vcdb, pcdb, qdb, cachePath + @"\ACESinspector-fitment permutations.txt", treeConfigLimit, cachePath, concernForDisparate, respectQdbType); }); });
            taskList.Add(macroAppAnalysisTask);
            macroAppAnalysisTask.Start();




            if (verbose) { Console.WriteLine("Analyzing applications..."); }

            bool stuffStillRunning = true;
            while (stuffStillRunning)
            {
                stuffStillRunning = false;
                foreach (Task t in taskList)
                {
                    if (!t.IsCompleted)
                    {
                        stuffStillRunning = true;
                    }
                }

            }

            if (verbose) { Console.WriteLine("  Done"); }

            //-------------------------------------------- all analysis is done and recorded at this point (all tasks terminated) -------------------------------------------------
            // compile the total warning and errors counts 
            aces.parttypePositionErrorsCount = 0; aces.qdbErrorsCount = 0; aces.questionableNotesCount = 0; aces.basevehicleidsErrorsCount = 0; aces.vcdbCodesErrorsCount = 0; aces.vcdbConfigurationsErrorsCount = 0; aces.parttypeDisagreementCount = 0; aces.qtyOutlierCount = 0;
            foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
            {
                aces.parttypePositionErrorsCount += chunk.parttypePositionErrorsCount;
                aces.qdbErrorsCount += chunk.qdbErrorsCount;
                aces.questionableNotesCount += chunk.questionableNotesCount;
                aces.basevehicleidsErrorsCount += chunk.basevehicleidsErrorsCount;
                aces.vcdbCodesErrorsCount += chunk.vcdbCodesErrorsCount;
                aces.vcdbConfigurationsErrorsCount += chunk.vcdbConfigurationsErrorsCount;
            }

            foreach (analysisChunk chunk in aces.outlierAnanlysisChunksList)
            {// should only be one item in list 
                aces.parttypeDisagreementCount += chunk.parttypeDisagreementErrorsCount;
                aces.qtyOutlierCount += chunk.qtyOutlierCount;
                aces.assetProblemsCount += chunk.assetProblemsCount;
            }



            aces.fitmentLogicProblemsCount = 0;
            int problemGroupNumber = 0;

            for (int i = 0; i < aces.fitmentAnalysisChunksGroups.Count; i++)
            {
                for (int j = 0; j < aces.fitmentAnalysisChunksGroups[i].chunkList.Count; j++)
                {
                    if (aces.fitmentAnalysisChunksGroups[i].chunkList[j].problemAppsList.Count > 0)
                    {
                        aces.fitmentLogicProblemsCount += aces.fitmentAnalysisChunksGroups[i].chunkList[j].problemAppsList.Count;
                        problemGroupNumber++;
                        if (reportAllAppsInProblemGroup)
                        {
                            aces.fitmentProblemGroupsAppLists.Add(problemGroupNumber.ToString(), aces.fitmentAnalysisChunksGroups[i].chunkList[j].appsList);
                        }
                        else
                        {
                            aces.fitmentProblemGroupsAppLists.Add(problemGroupNumber.ToString(), aces.fitmentAnalysisChunksGroups[i].chunkList[j].problemAppsList);
                        }
                        aces.fitmentProblemGroupsBestPermutations.Add(problemGroupNumber.ToString(), aces.fitmentAnalysisChunksGroups[i].chunkList[j].lowestBadnessPermutation);
                    }
                }
            }

            if (verbose) { Console.WriteLine((aces.basevehicleidsErrorsCount + aces.vcdbCodesErrorsCount + aces.vcdbConfigurationsErrorsCount + aces.qdbErrorsCount + aces.parttypePositionErrorsCount).ToString() + " errors"); }
            if (logFile != "") 
            {
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tbasevehicleidsErrorsCount:" + aces.basevehicleidsErrorsCount.ToString() + Environment.NewLine);
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tvcdbCodesErrorsCount:" + aces.vcdbCodesErrorsCount.ToString() + Environment.NewLine);
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tvcdbConfigurationsErrorsCount:" + aces.vcdbConfigurationsErrorsCount + Environment.NewLine);
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tqdbErrorsCount:" + aces.qdbErrorsCount + Environment.NewLine);
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tparttypePositionErrorsCount:" + aces.parttypePositionErrorsCount.ToString() + Environment.NewLine);
            }


            List<string> problemsListTemp = new List<string>();
            if (aces.fitmentLogicProblemsCount > 0) { problemsListTemp.Add(aces.fitmentLogicProblemsCount.ToString() + " logic flaws"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tfitmentLogicProblemsCount:" + aces.fitmentLogicProblemsCount.ToString() + Environment.NewLine); }

            if (aces.qtyOutlierCount > 0) { problemsListTemp.Add(aces.qtyOutlierCount.ToString() + " qty outliers"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tqtyOutlierCount:" + aces.qtyOutlierCount.ToString() + Environment.NewLine); }


            if (aces.parttypeDisagreementCount > 0) { problemsListTemp.Add(aces.parttypeDisagreementCount.ToString() + " type disagreements"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tparttypeDisagreementCount:" + aces.parttypeDisagreementCount.ToString() + Environment.NewLine); }


            if (aces.assetProblemsCount > 0) { problemsListTemp.Add(aces.assetProblemsCount.ToString() + " Asset problems"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tassetProblemsCount:" + aces.assetProblemsCount.ToString() + Environment.NewLine); }


            if (problemsListTemp.Count() == 0) { macroProblemsDescription = "0 problems"; } else { macroProblemsDescription = string.Join(", ", problemsListTemp); }

            if (verbose) { Console.WriteLine(macroProblemsDescription); }
            if (verbose) { Console.WriteLine("writing assessment file"); }

            DateTime endingDateTime = DateTime.Now;
            TimeSpan runTime = endingDateTime - startingDateTime;
            if (verbose) { Console.WriteLine("run time: " + Math.Round(runTime.TotalMilliseconds/1000,1).ToString() + " seconds"); }
            if (logFile != "") { File.AppendAllText(logFile, DateTime.Now.ToString() + "\tAnalysis took " + Math.Round(runTime.TotalMilliseconds / 1000, 1).ToString() + " seconds" + Environment.NewLine); }


            // start building the assessment file

            string validatedAgainstVCdb = ""; if (aces.VcdbVersionDate != vcdb.version) { validatedAgainstVCdb = "analyzed against:" + vcdb.version; }
            string validatedAgainstPCdb = ""; if (aces.PcdbVersionDate != pcdb.version) { validatedAgainstPCdb = "analyzedagainst:" + pcdb.version; }
            string validatedAgainstQdb = ""; if (aces.QdbVersionDate != qdb.version) { validatedAgainstQdb = "analyzed against:" + qdb.version; }
            string excelTabColorXMLtag = "";

            string assessmentFilename = assessmentsPath + "\\" + Path.GetFileNameWithoutExtension(aces.filePath) + "_assessment.xml";

            // calculate basevehilce usage against vcdb's total possible basevehilce offering
            int basevehicleHitcount = 0;
            int modernBasevehicleHitcount = 0;
            int modernBasevehiclesAvail = 0;
            foreach (KeyValuePair<int, BaseVehicle> entry in vcdb.vcdbBasevhicleDict)
            {
                if (Convert.ToInt32(entry.Value.YearId) >= 1990) { modernBasevehiclesAvail++; }

                if (aces.basevidOccurrences.ContainsKey(entry.Key))
                {
                    basevehicleHitcount++;
                    if ( Convert.ToInt32(entry.Value.YearId) >= 1990) // && entry.Value.VehicleTypeName=="Car"
                    {
                        //sw.WriteLine(entry.Key.ToString() + "\t" + entry.Value.MakeName + "\t" + entry.Value.ModelName + "\t" + entry.Value.YearId + "\t" + entry.Value.VehicleTypeName);
                        modernBasevehicleHitcount++;
                    }
                }
            }
            if (logFile != "") 
            {
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tBasevehicle coverage (all years) " + Math.Round(Convert.ToDouble(basevehicleHitcount * 100) / (vcdb.vcdbBasevhicleDict.Count + 1), 1).ToString() + "%  (" + basevehicleHitcount + " used, " + vcdb.vcdbBasevhicleDict.Count + " available)" + Environment.NewLine);
                File.AppendAllText(logFile, DateTime.Now.ToString() + "\tBasevehicle coverage (1990+) " + Math.Round(Convert.ToDouble(modernBasevehicleHitcount * 100) / (modernBasevehiclesAvail + 1), 1).ToString() + "%  (" + modernBasevehicleHitcount + " used, " + modernBasevehiclesAvail + " available)" + Environment.NewLine);
            }

            try
            {
                using (StreamWriter sw = new StreamWriter(assessmentFilename))
                {
                    sw.Write("<?xml version=\"1.0\"?><?mso-application progid=\"Excel.Sheet\"?><Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\"><DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"><Author>ACESinspector</Author><LastAuthor>ACESinspector</LastAuthor><Created>2017-02-20T01:10:23Z</Created><LastSaved>2017-02-20T02:49:36Z</LastSaved><Version>14.00</Version></DocumentProperties><OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\"><AllowPNG/></OfficeDocumentSettings><ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"><WindowHeight>7500</WindowHeight><WindowWidth>15315</WindowWidth><WindowTopX>120</WindowTopX><WindowTopY>150</WindowTopY><TabRatio>785</TabRatio><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook><Styles><Style ss:ID=\"Default\" ss:Name=\"Normal\"><Alignment ss:Vertical=\"Bottom\"/><Borders/><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/><Interior/><NumberFormat/><Protection/></Style><Style ss:ID=\"s62\"><NumberFormat ss:Format=\"Short Date\"/></Style><Style ss:ID=\"s64\" ss:Name=\"Hyperlink\"><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#0000FF\" ss:Underline=\"Single\"/></Style><Style ss:ID=\"s65\"><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/><Interior ss:Color=\"#D9D9D9\" ss:Pattern=\"Solid\"/></Style></Styles><Worksheet ss:Name=\"Stats\"><Table ss:ExpandedColumnCount=\"3\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"170\"/><Column ss:Width=\"171\"/><Column ss:Width=\"144\"/>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Input Filename</Data></Cell><Cell><Data ss:Type=\"String\">" + Path.GetFileName(aces.filePath) + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Title</Data></Cell><Cell><Data ss:Type=\"String\">" + aces.DocumentTitle + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Brand</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + aces.BrandAAIAID + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">ACES version</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + aces.version + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">VCdb version cited</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + aces.VcdbVersionDate + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + validatedAgainstVCdb + "</Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">PCdb version cited</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + aces.PcdbVersionDate + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + validatedAgainstPCdb + "</Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Qdb version cited</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + aces.QdbVersionDate + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + validatedAgainstQdb + "</Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Application count</Data></Cell><Cell><Data ss:Type=\"Number\">" + aces.apps.Count.ToString() + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Unique Part count</Data></Cell><Cell><Data ss:Type=\"Number\">" + aces.partsAppCounts.Count.ToString() + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Unique MfrLabel count</Data></Cell><Cell><Data ss:Type=\"Number\">" + aces.distinctMfrLabels.Count.ToString() + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Unique Parttypes count</Data></Cell><Cell><Data ss:Type=\"Number\">" + aces.distinctPartTypes.Count.ToString() + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");

                    if ((aces.parttypePositionErrorsCount + aces.vcdbCodesErrorsCount + aces.vcdbConfigurationsErrorsCount + aces.basevehicleidsErrorsCount + aces.qdbErrorsCount + aces.fitmentLogicProblemsCount) > 0)
                    {
                        List<string> failureReasons = new List<string>();
                        if (aces.parttypePositionErrorsCount > 0) { failureReasons.Add(aces.parttypePositionErrorsCount.ToString() + " partType-position pairings"); }
                        if (aces.vcdbCodesErrorsCount > 0) { failureReasons.Add(aces.vcdbCodesErrorsCount.ToString() + " invalid VCdb codes"); }
                        if (aces.vcdbConfigurationsErrorsCount > 0) { failureReasons.Add(aces.vcdbConfigurationsErrorsCount.ToString() + " invalid VCdb configs"); }
                        if (aces.basevehicleidsErrorsCount > 0) { failureReasons.Add(aces.basevehicleidsErrorsCount.ToString() + " invalid basevehicles"); }
                        if (aces.qdbErrorsCount > 0) { failureReasons.Add(aces.qdbErrorsCount.ToString() + " Qdb errors"); }
                        if (aces.fitmentLogicProblemsCount > 0) { failureReasons.Add(aces.fitmentLogicProblemsCount.ToString() + " fitment logic problems"); }
                        sw.Write("<Row><Cell><Data ss:Type=\"String\">Result</Data></Cell><Cell><Data ss:Type=\"String\">Fail</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + string.Join(",", failureReasons) + "</Data></Cell></Row>");
                    }
                    else
                    {
                        sw.Write("<Row><Cell><Data ss:Type=\"String\">Result</Data></Cell><Cell><Data ss:Type=\"String\">Pass</Data></Cell></Row>");
                    }

                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Qdb Utilization (%)</Data></Cell><Cell><Data ss:Type=\"Number\">" + aces.QdbUtilizationScore.ToString("0.00") + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">All BaseVehilce Coverage (%)</Data></Cell><Cell><Data ss:Type=\"Number\">" + Math.Round(Convert.ToDouble(basevehicleHitcount * 100) / (vcdb.vcdbBasevhicleDict.Count + 1), 1).ToString("0.00") + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + basevehicleHitcount.ToString() + " used, " + vcdb.vcdbBasevhicleDict.Count.ToString() + " available" + "</Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">1990+ BaseVehilce Coverage (%)</Data></Cell><Cell><Data ss:Type=\"Number\">" + Math.Round(Convert.ToDouble(modernBasevehicleHitcount * 100) / (modernBasevehiclesAvail + 1), 1).ToString("0.00") + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\">" + modernBasevehicleHitcount.ToString() + " used, " + modernBasevehiclesAvail.ToString() + " available" + "</Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Validation tool</Data></Cell><Cell ss:StyleID=\"s64\" ss:HRef=\"https://github.com/zerodbu/ACESinspectorCLI\"><Data ss:Type=\"String\">ACESinspectorCLI version 1.0.0.17</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("<Row><Cell><Data ss:Type=\"String\">Processing Time (Seconds)</Data></Cell><Cell><Data ss:Type=\"Number\">" + Math.Round(runTime.TotalMilliseconds / 1000, 1).ToString() + "</Data></Cell><Cell ss:StyleID=\"s62\"><Data ss:Type=\"String\"></Data></Cell></Row>");
                    sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><Selected/><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");

                    sw.Write("<Worksheet ss:Name=\"Parts\"><Table ss:ExpandedColumnCount=\"4\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Applications Count</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Types</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Positions</Data></Cell></Row>");
                    foreach (KeyValuePair<string, int> partsAppCountEntry in aces.partsAppCounts)
                    {
                        partTypeNameList.Clear(); foreach (int partTypeId in aces.partsPartTypes[partsAppCountEntry.Key]) { partTypeNameList.Add(pcdb.niceParttype(partTypeId)); }
                        positionNameList.Clear(); foreach (int positionId in aces.partsPositions[partsAppCountEntry.Key]) { positionNameList.Add(pcdb.nicePosition(positionId)); }
                        sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(partsAppCountEntry.Key) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + partsAppCountEntry.Value.ToString() + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(string.Join(",", partTypeNameList)) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(String.Join(",", positionNameList)) + "</Data></Cell></Row>");
                    }
                    sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");

                    sw.Write("<Worksheet ss:Name=\"Part Types\"><Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Index=\"2\" ss:AutoFitWidth=\"0\" ss:Width=\"183.75\"/>");
                    foreach (int distinctPartType in aces.distinctPartTypes) { sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + distinctPartType + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(pcdb.niceParttype(distinctPartType)) + "</Data></Cell></Row>"); }
                    sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");

                    if (aces.distinctMfrLabels.Count > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"MfrLabels\"><Table ss:ExpandedColumnCount=\"1\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"151.5\"/>");
                        foreach (string distinctMfrLabel in aces.distinctMfrLabels) { sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(distinctMfrLabel) + "</Data></Cell></Row>"); }
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    // basevid listing
                    sw.Write("<Worksheet ss:Name=\"BaseVehicle Usage\"><Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"84\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"70\"/>");
                    sw.Write("<Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">BaseVehicle ID</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Occurrences</Data></Cell></Row>");
                    foreach (KeyValuePair<int, int> entry in aces.basevidOccurrences)
                    {
                        sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + entry.Key.ToString() + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + entry.Value.ToString() + "</Data></Cell></Row>");
                    }
                    sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");


                    // Qdb listing
                    sw.Write("<Worksheet ss:Name=\"Qdb Usage\"><Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"50\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"70\"/>");
                    sw.Write("<Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Qdb ID</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Occurrences</Data></Cell></Row>");
                    foreach (KeyValuePair<int, int> entry in aces.qdbidOccurrences)
                    {
                        sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + entry.Key.ToString() + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + entry.Value.ToString() + "</Data></Cell></Row>");
                    }
                    sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");



                    if (aces.noteCounts.Count > 0)
                    {
                        if (noteToQdbTransformDictionary.Count() > 0)
                        {// transforms are in play - they will be listed in the third column
                            sw.Write("<Worksheet ss:Name=\"Note Tags\"><Table ss:ExpandedColumnCount=\"3\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"350\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"62.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"350\"/>");
                            sw.Write("<Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Note Text</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Occurrences</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Qdb Transform</Data></Cell></Row>");
                        }
                        else
                        {// no transforms in play - list the notes and counts only
                            sw.Write("<Worksheet ss:Name=\"Note Tags\"><Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"350\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"62.5\"/>");
                            sw.Write("<Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Note Text</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Occurrences</Data></Cell></Row>");
                        }

                        foreach (KeyValuePair<string, int> noteEntry in aces.noteCounts)
                        {
                            if (noteToQdbTransformDictionary.ContainsKey(noteEntry.Key))
                            {
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(noteEntry.Key) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + noteEntry.Value + "</Data></Cell><Cell><Data ss:Type=\"String\">" + qdb.niceQdbQualifier(noteToQdbTransformDictionary[noteEntry.Key].qualifierId, noteToQdbTransformDictionary[noteEntry.Key].qualifierParameters) + "</Data></Cell></Row>");
                            }
                            else
                            {
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(noteEntry.Key) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + noteEntry.Value + "</Data></Cell></Row>");
                            }
                        }
                        excelTabColorXMLtag = "";// "<TabColorIndex>13</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    if (aces.parttypeDisagreementCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Parttype Disagreement\"><Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"45\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"78.75\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Parttypes</Data></Cell></Row>");
                        using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_parttypeDisagreements.txt"))
                        {
                            while (!reader.EndOfStream)
                            {
                                string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + fileds[1] + "</Data></Cell></Row>");
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>13</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }


                    if (aces.qtyOutlierCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Qty Outliers\"><Table ss:ExpandedColumnCount=\"13\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"180\"/><Column ss:Width=\"36\"/><Column ss:Width=\"77\"/><Column ss:Width=\"120\"/><Column ss:Width=\"33\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"46\"/><Column ss:Width=\"120\"/><Column ss:Width=\"180\"/><Column ss:Width=\"180\"/><Column ss:Width=\"180\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Error Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">VCdb-coded attributes</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Qdb-coded qualifiers</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Notes</Data></Cell></Row>");
                        using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_qtyOutliers.txt"))
                        {
                            while (!reader.EndOfStream)
                            {
                                string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[8] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[12]) + "</Data></Cell></Row>");
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>13</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    if (aces.parttypePositionErrorsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"PartType-Position Errors\"><Table ss:ExpandedColumnCount=\"11\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"115\"/><Column ss:Width=\"36\"/><Column ss:Width=\"77\"/><Column ss:Width=\"120\"/><Column ss:Width=\"33\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"46\"/><Column ss:Width=\"120\"/><Column ss:Width=\"180\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Error Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Fitment</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.parttypePositionErrorsCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_parttypePositionErrors" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[0]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[8] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    if (aces.qdbErrorsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Qdb Errors\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"115\"/><Column ss:Width=\"36\"/><Column ss:Width=\"77\"/><Column ss:Width=\"120\"/><Column ss:Width=\"33\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"46\"/><Column ss:Width=\"120\"/><Column ss:Width=\"180\"/><Column ss:Width=\"180\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Error Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">VCdb-coded attributes</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Notes</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.qdbErrorsCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_qdbErrors" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[0]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[8] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }


                    if (aces.questionableNotesCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Questionable Notes\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"115\"/><Column ss:Width=\"36\"/><Column ss:Width=\"77\"/><Column ss:Width=\"120\"/><Column ss:Width=\"33\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"46\"/><Column ss:Width=\"120\"/><Column ss:Width=\"180\"/><Column ss:Width=\"180\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Error Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">VCdb-coded attributes</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Notes</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.questionableNotesCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_questionableNotes" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[0]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[8] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>13</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    //-----------
                    if (aces.assetProblemsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Asset Problems\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"180\"/><Column ss:Width=\"50\"/><Column ss:Width=\"50\"/><Column ss:Width=\"77\"/><Column ss:Width=\"120\"/><Column ss:Width=\"33\"/><Column ss:Width=\"120\"/><Column ss:Width=\"120\"/><Column ss:Width=\"80\"/><Column ss:Width=\"46\"/><Column ss:Width=\"50\"/><Column ss:Width=\"180\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Problem Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Reference</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Fitment</Data></Cell></Row>");
                        using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_assetProblems.txt"))
                        {
                            while (!reader.EndOfStream)
                            {
                                string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[3] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[8]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[9] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell></Row>");
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>13</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }
                    //---------------

                    if (aces.basevehicleidsErrorsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name =\"Invalid Base Vids\"><Table ss:ExpandedColumnCount=\"7\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"45\"/><Column ss:Width=\"77.25\"/><Column ss:Index=\"4\" ss:AutoFitWidth=\"0\" ss:Width=\"96\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"73.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"253.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"371.25\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Invalid BaseVid</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Fitment</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.basevehicleidsErrorsCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidBasevehicles" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[2]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[4] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><Print><ValidPrinterInfo/><HorizontalResolution>600</HorizontalResolution><VerticalResolution>600</VerticalResolution></Print>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    if (aces.vcdbCodesErrorsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Invalid VCdb Codes\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"45\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"54\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"76\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"78.75\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"99.75\"/><Column ss:Width=\"31.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"60\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"112.5\"/><Column ss:Width=\"43.5\"/><Column ss:Width =\"43.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"237\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"319.5\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Reference</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehicle id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">VCdb Attributes</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Notes</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.vcdbCodesErrorsCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_invalidVCdbCodes" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[1]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[2] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[7]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[8] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type =\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }


                    if (aces.vcdbConfigurationsErrorsCount > 0)
                    {
                        sw.Write("<Worksheet ss:Name=\"Invalid VCdb Configs\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:AutoFitWidth=\"0\" ss:Width=\"45\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"45\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"78.75\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"99.75\"/><Column ss:Width=\"31.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"60\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"112.5\"/><Column ss:Width=\"43.5\"/><Column ss:Width =\"43.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"237\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"319.5\"/><Column ss:AutoFitWidth=\"0\" ss:Width=\"319.5\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Base Vehiclce id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">VCdb-coded Attributes</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Qdb-coded Qualifiers</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Notes</Data></Cell></Row>");
                        foreach (analysisChunk chunk in aces.individualAnanlysisChunksList)
                        {
                            if (chunk.vcdbConfigurationsErrorsCount > 0)
                            {
                                try
                                {
                                    using (var reader = new StreamReader(cachePath + "\\AiFragments\\" + aces.fileMD5hash + "_configurationErrors" + chunk.id.ToString() + ".txt"))
                                    {
                                        while (!reader.EndOfStream)
                                        {
                                            string line = reader.ReadLine(); string[] fileds = line.Split('\t');
                                            sw.Write("<Row><Cell><Data ss:Type=\"Number\">" + fileds[0] + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[1] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[2]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[3]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[4]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[5]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[6]) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + fileds[7] + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[8]) + "</Data></Cell><Cell><Data ss:Type =\"String\">" + escapeXMLspecialChars(fileds[9]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[10]) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(fileds[11]) + "</Data></Cell></Row>");
                                        }
                                    }
                                }
                                catch (Exception ex) { }
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }


                    if (aces.fitmentLogicProblemsCount > 0)
                    {
                        problemDescription = "";
                        elementPrevalence = 0;
                        fitmentElementPrevalence.Clear();

                        sw.Write("<Worksheet ss:Name=\"Fitment Logic Problems\"><Table ss:ExpandedColumnCount=\"12\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\"><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Column ss:Width=\"100\"/><Row><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Problem Description</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Group</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">App Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">BaseVehcile Id</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Make</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Model</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Year</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part Type</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Position</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Quantity</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Part</Data></Cell><Cell ss:StyleID=\"s65\"><Data ss:Type=\"String\">Fitment</Data></Cell></Row>");
                        foreach (KeyValuePair<string, List<App>> entry in aces.fitmentProblemGroupsAppLists)
                        {
                            // construct a tree in order to re-discover the problmes with it.
                            aces.fitmentNodeList.Clear();
                            fitmentElementPrevalence.Clear();
                            foreach (string fitmentElement in aces.fitmentProblemGroupsBestPermutations[entry.Key])
                            {
                                fitmentElementPrevalence.Add(fitmentElement, elementPrevalence); elementPrevalence++;
                            }
                            aces.fitmentNodeList.AddRange(aces.buildFitmentTreeFromAppList(entry.Value, fitmentElementPrevalence, -1, false, false, vcdb, qdb));
                            problemDescription = aces.fitmentTreeProblemDescription(aces.fitmentNodeList, qdb, concernForDisparate, respectQdbType);

                            foreach (App app in entry.Value)
                            {
                                sw.Write("<Row><Cell><Data ss:Type=\"String\">" + problemDescription + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(entry.Key) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + app.id.ToString() + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + app.basevehilceid.ToString() + "</Data></Cell><Cell><Data ss:Type=\"String\">" + vcdb.niceMakeOfBasevid(app.basevehilceid) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(vcdb.niceModelOfBasevid(app.basevehilceid)) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + vcdb.niceYearOfBasevid(app.basevehilceid) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(pcdb.niceParttype(app.parttypeid)) + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(pcdb.nicePosition(app.positionid)) + "</Data></Cell><Cell><Data ss:Type=\"Number\">" + app.quantity.ToString() + "</Data></Cell><Cell><Data ss:Type=\"String\">" + app.part + "</Data></Cell><Cell><Data ss:Type=\"String\">" + escapeXMLspecialChars(app.niceFullFitmentString(vcdb, qdb)) + "</Data></Cell></Row>");
                            }
                        }
                        excelTabColorXMLtag = "<TabColorIndex>10</TabColorIndex>";
                        sw.Write("</Table><WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup>" + excelTabColorXMLtag + "<FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>0</ActiveRow></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>");
                    }

                    sw.Write("</Workbook>");
                }
            }
            catch (Exception ex)
            {
                if (verbose) { Console.WriteLine("Assessment file NOT created: " + ex.Message); }
            }


            if (verbose) { Console.WriteLine("deleting temp files"); }

            foreach (string cachefile in cacheFilesToDeleteOnExit)
            {
                try { File.Delete(cachefile); } catch (Exception ex) { }
            }


            if (deleteACESfileOnSuccess)
            {
                if (verbose) { Console.WriteLine("deleting input file"); }
                try
                {
                    File.Delete(inputFile);
                }
                catch
                (
                Exception ex)
                {
                    if (verbose) { Console.WriteLine("failed to delete input file: " + ex.Message); }
                }
            }

            return 0; // successful analysis

        }

    }

}
