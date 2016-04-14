namespace SqlExport.Console
{
    using System;
    using System.Linq;

    using CommandLine;
    using CommandLine.Text;

    public class Options
    {
        [Option("date", SetName = "one", Required = true)]
        public string Date { get; set; }

        [Option("fromdate", SetName = "two", Required = true)]
        public DateTime FromDate { get; set; }

        [Option("todate", SetName = "two", Required = true)]
        public DateTime ToDate { get; set; }
        
        ////[HelpOption]
        ////public string GetUsage()
        ////{
        ////    var help = new HelpText
        ////    {
        ////        Heading = new HeadingInfo("SqlExport", "<<app version>>"),
        ////        Copyright = new CopyrightInfo("TSS", 2014),
        ////        AdditionalNewLineAfterOption = true,
        ////        AddDashesToOption = true
        ////    };
          
        ////    help.AddOptions(this);

        ////    if (this.LastParserState.Errors.Any())
        ////    {
        ////        var errors = help.RenderParsingErrorsText(this, 2); // indent with two spaces

        ////        if (!string.IsNullOrEmpty(errors))
        ////        {
        ////            help.AddPreOptionsLine(string.Concat(Environment.NewLine, "ERROR(S):"));
        ////            help.AddPreOptionsLine(errors);
        ////        }
        ////    }

        ////    return help;
        ////}
    }
}