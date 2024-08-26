using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MimeKit;

namespace LeitorDeEmail
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Por favor, forneça o caminho do arquivo mbox.");
                return;
            }

            string filePath = args[0];
            var mailManager = new MailManager();
            mailManager.StartProcess(filePath);
        }
    }

    public class MailManager
    {
        public void StartProcess(string p_filePath)
        {
            var l_files = Directory.GetFiles(p_filePath, "*.*", SearchOption.AllDirectories)
                                   .Where(f => f.EndsWith(".mbox") || f.EndsWith(".eml"));

            foreach (var l_file in l_files)
            {
                try
                {
                    var l_rides = ProcessMailFile(l_file);
                }
                catch (Exception)
                {
                    // Tratamento de erro básico para evitar que o programa pare.
                    Console.WriteLine($"Erro ao processar o arquivo '{l_file}'");
                }
            }
        }

        public List<FileMailInfo> ProcessMailFile(string p_file)
        {
            var l_rides = new List<FileMailInfo>();
            var fileExtension = Path.GetExtension(p_file).ToLower();

            try
            {
                switch (fileExtension)
                {
                    case ".mbox":
                        l_rides.AddRange(ProcessMbox(p_file));
                        break;

                    case ".eml":
                        l_rides.AddRange(ProcessEml(p_file));
                        break;

                    default:
                        Console.WriteLine($"Extensão de arquivo não suportada: {fileExtension}");
                        break;
                }
            }
            catch (Exception)
            {
                Console.WriteLine($"Erro ao processar o arquivo '{p_file}'");
            }

            return l_rides;
        }

        private List<FileMailInfo> ProcessMbox(string p_file)
        {
            var l_rides = new List<FileMailInfo>();
            var l_directory = Path.GetDirectoryName(p_file);
            var l_attachmentsDirectory = Path.Combine(l_directory, "Anexos");

            Directory.CreateDirectory(l_attachmentsDirectory);

            using (var l_stream = File.OpenRead(p_file))
            {
                var l_process = new MimeParser(l_stream, MimeFormat.Mbox);

                while (!l_process.IsEndOfStream)
                {
                    try
                    {
                        var l_messageInfo = l_process.ParseMessage();
                        var l_ride = ProcessAttachment(l_messageInfo, l_attachmentsDirectory);
                        l_rides.Add(l_ride);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine($"Erro ao processar a mensagem no arquivo '{p_file}'");
                    }
                }
            }

            return l_rides;
        }

        private List<FileMailInfo> ProcessEml(string p_file)
        {
            var l_rides = new List<FileMailInfo>();
            var l_directory = Path.GetDirectoryName(p_file);
            var l_attachmentsDirectory = Path.Combine(l_directory, "Anexos");

            Directory.CreateDirectory(l_attachmentsDirectory);

            try
            {
                using (var l_stream = File.OpenRead(p_file))
                {
                    var l_messageInfo = MimeMessage.Load(l_stream);
                    var l_ride = ProcessAttachment(l_messageInfo, l_attachmentsDirectory);
                    l_rides.Add(l_ride);
                }
            }
            catch (Exception)
            {
                Console.WriteLine($"Erro ao carregar a mensagem do arquivo '{p_file}'");
            }

            return l_rides;
        }

        private FileMailInfo ProcessAttachment(MimeMessage p_message, string p_attachmentsDirectory)
        {
            var l_messageId = Guid.NewGuid().ToString();

            try
            {
                foreach (var l_attachment in p_message.Attachments.OfType<MimePart>())
                {
                    var simplifiedFileName = SimplifyFileName(l_attachment.FileName);
                    var l_fileName = $"{l_messageId}_{simplifiedFileName}";
                    var l_attachmentPath = Path.Combine(p_attachmentsDirectory, l_fileName);

                    try
                    {
                        using (var l_fileStream = File.Create(l_attachmentPath))
                        {
                            l_attachment.Content.DecodeTo(l_fileStream);
                        }
                    }
                    catch (Exception)
                    {
                        Console.WriteLine($"Erro ao processar o anexo '{simplifiedFileName}' da mensagem '{p_message.Subject}'");
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine($"Erro ao processar anexos da mensagem '{p_message.Subject}'");
            }

            return new FileMailInfo
            {
                Subject = p_message.Subject,
                Body = p_message.TextBody ?? p_message.HtmlBody,
                From = p_message.From.ToString(),
                To = string.Join(", ", p_message.To.Mailboxes.Select(m => m.Address)),
                Date = p_message.Date.ToString(),
                ExistingAttachment = p_message.Attachments.Any() ? "yes" : "no",
                MessageId = l_messageId
            };
        }

        private string SimplifyFileName(string fileName)
        {
            return new string(fileName.Where(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == '.').ToArray());
        }
    }
}
