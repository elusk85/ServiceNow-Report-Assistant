using ServiceNow_Report_Assistant.WorkBooks;
using System;


namespace ServiceNow_Report_Assistant.UserInterfaces
{
    public static class StartingUserInterface
    {
        public static bool Quit = false;
        public static void CommandLoop()
        {
            while (!Quit)
            {
                Console.WriteLine("What would you like to do?");
                var command = Console.ReadLine().ToLower();
                CommandRoute(command);
            }
        }

        public static void CommandRoute(string command)
        {
            if (command.StartsWith("combine"))
                CombineCommand(command);
            else if (command.StartsWith("format"))
                FormatCommand(command);
            else if (command == "help")
                HelpCommand();
            else if (command == "quit")
                Quit = true;
            else
                Console.WriteLine("{0} was not recognized, please try again.", command);
        }

        public static void CombineCommand(string command)
        {
            var parts = command.Split(' ');
            //if (parts.Length != 5)
            //{
            //    Console.WriteLine("Command not valid, Combine requires the name of both reports, their corresponding paths, and the type of report for formatting.");
            //    return;
            //}
            var name1 = parts[1];
            var path1 = parts[2];
            var name2 = parts[3];
            var path2 = parts[4];
            var type = parts[5].ToLower();
            
            //BasicReport basereport;
            //if (type == "Asset")
            //    basereport = new AssetReport(name, path);
            //else if (type == "CAB")
            //    basereport = new CABReport(name, path);
            //else
            //{
            //    Console.WriteLine("{0} is not a supported type of report, please try again.", type);
            //    return;
            //}
            Console.WriteLine("Created report {0}.", name1);
            //AssistantUserInterface.CommandLoop(gradeBook);
        }

        public static void FormatCommand(string command)
        {
            var parts = command.Split(' ');
            if (parts.Length != 2)
            {
                Console.WriteLine("Command not valid, Format requires a name and type for the report.");
                return;
            }
        }

        public static void HelpCommand()
        {
            Console.WriteLine("Report Assistant accepts the following commands:");
            Console.WriteLine();
            Console.WriteLine("Combine 'Name1' 'Path1 'Name2' 'Path2' 'Type' - Combines 2 or more reports. The type field determines the formatting of the result.");
            Console.WriteLine();
            Console.WriteLine("Format 'Name' - Formats the report with the provided 'Name'.");
            Console.WriteLine();
            Console.WriteLine("Help - Displays all accepted commands.");
            Console.WriteLine();
            Console.WriteLine("Quit - Exits the application");
        }
    }
}
