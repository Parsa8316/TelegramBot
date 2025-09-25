using System;
using System.Collections.Generic;
using System.Text;

namespace TelegramBot
{
    public static class Token
    {
        private const string token = "yourToken";

        public static string GetToken()
        {
            return token;
        }
    }
}
