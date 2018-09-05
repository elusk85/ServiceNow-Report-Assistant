using System;
using ServiceNow_Report_Assistant.UserInterfaces;

namespace ServiceNow_Report_Assistant
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("#=============================================#");
            Console.WriteLine("# Welcome to the ServiceNow Report Assistant! #");
            Console.WriteLine("#=============================================#");
            Console.WriteLine();

            StartingUserInterface.CommandLoop();

            Console.WriteLine("Thank you for using ServiceNow Report Assistant!");
            Console.WriteLine("Have a nice day!");
            Console.Read();
        }
    }
}
