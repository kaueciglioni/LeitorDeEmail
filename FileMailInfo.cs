using System;

namespace LeitorDeEmail
{

    public class FileMailInfo
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Date { get; set; }
        public string ExistingAttachment { get; set; }
        public string MessageId { get; set; }
    }
}

