using System;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Types;
using System.IO;
using System.Net.Http;

namespace TelegramBot
{
    class Program
    {
        static void Main(string[] args)
        {
            Task.Run(() => MyBot());
            Console.ReadKey();
        }

        static async Task MyBot()
        {
            try
            {
                TelegramBotClient bot = new TelegramBotClient(Token.GetToken());

                int offset = 0;
                while (true)
                {
                    Update[] updates = bot.GetUpdates(offset).Result;
                    foreach (Update update in updates)
                    {
                        offset = update.Id + 1;
                        try
                        {
                            if (update.Message.Document != null)
                            {
                                var document = update.Message.Document;
                                string mimeType = document.MimeType;
                                string fileId = document.FileId;
                                string fileName = document.FileName;

                                var file = await bot.GetFile(fileId);
                                string filePath = file.FilePath;
                                string fileUrl = $"https://api.telegram.org/file/bot{bot.Token}/{filePath}";

                                using (var client = new HttpClient())
                                using (var stream = await client.GetStreamAsync(fileUrl))
                                using (var memory = new MemoryStream())
                                {
                                    await stream.CopyToAsync(memory);
                                    memory.Position = 0;

                                    string tempPdfPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(fileName) + ".pdf");

                                    MemoryStream outputPdf = new MemoryStream();

                                    bool x = true;
                                    if (mimeType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                                    {
                                        var doc = new Aspose.Words.Document(memory);
                                        doc.Save(tempPdfPath, Aspose.Words.SaveFormat.Pdf);
                                    }
                                    else if (mimeType == "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
                                             mimeType == "application/vnd.ms-powerpoint")
                                    {
                                        var presentation = new Aspose.Slides.Presentation(memory);
                                        presentation.Save(tempPdfPath, Aspose.Slides.Export.SaveFormat.Pdf);
                                    }
                                    else
                                    {
                                        await bot.SendMessage(update.Message.Chat.Id, "welcome to bot. \nwe can convert powerpoint or word to pdf.");
                                        x = false;
                                    }

                                    if (x)
                                    {
                                        using (var fileStream = System.IO.File.OpenRead(tempPdfPath))
                                        {
                                            var inputFile = new InputFileStream(fileStream, Path.GetFileName(tempPdfPath));

                                            await bot.SendDocument(
                                                chatId: update.Message.Chat.Id,
                                                document: inputFile
                                            );
                                        }
                                    }
                                    if (System.IO.File.Exists(tempPdfPath))
                                    {
                                        System.IO.File.Delete(tempPdfPath);
                                    }
                                }
                            }
                            else
                            {
                                await bot.SendMessage(update.Message.Chat.Id, "welcome to bot. \nwe can convert powerpoint or word to pdf.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            await bot.SendMessage(update.Message.Chat.Id, "ERROR");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }

        }
    }
}
