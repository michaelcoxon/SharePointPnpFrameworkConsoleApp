using System;
using System.Security;
using PnP.Framework;

namespace SharePointPnpFrameworkPlayground
{
    internal class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            var username = GetString("Enter username:");
            var password = GetSecureString("Enter password:");

            using (var authManager = new AuthenticationManager(username: username, password: password))
            {
                // check login
                try
                {
                    Console.WriteLine("Logging in...");
                    Console.WriteLine();

                    // Test the credentials
                    // TODO: ask again
                    //
                    // we use the .default scope so that we can make sure that the
                    // pnp management shell has been allowed in azure ad.
                    await authManager.GetAccessTokenAsync(new string[] { ".default" });

                    Console.WriteLine("Login successful!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine();

                    Console.Error.WriteLine("Login failed!");
                    Console.Error.WriteLine(ex.Message);
                    return; // die
                }


                /// Get site and then do work....
                try
                {
                    var site = GetString("Enter site URL:");
                    Console.Write("Loading...");

                    // do sharepoint things
                    using (var context = await authManager.GetContextAsync(site))
                    {
                        Console.WriteLine("OK");

                        Console.Write("Loading site...");

                        // queue load an object from remote
                        context.Load(context.Web);

                        // excute load queue as query
                        await context.ExecuteQueryAsync();

                        Console.WriteLine("OK");
                        Console.WriteLine();

                        // dump site title
                        Console.WriteLine($"Site title: '{context.Web.Title}'");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine();

                    Console.Error.WriteLine(ex.Message);
                    return; // die
                }
            }
        }

        //
        // Console helper methods
        //

        private static string GetString(string prompt, string value = default)
        {
            Console.WriteLine(prompt);

            if (value == default)
            {
                value = Console.ReadLine();
            }
            else
            {
                Console.WriteLine(value);
            }

            Console.WriteLine();
            return value;
        }

        private static SecureString GetSecureString(string prompt, string value = default)
        {
            Console.WriteLine(prompt);

            var secureString = new SecureString();
            {
                if (value == default)
                {
                    ConsoleKey key;
                    do
                    {
                        var keyInfo = Console.ReadKey(intercept: true);
                        key = keyInfo.Key;

                        if (key == ConsoleKey.Backspace && secureString.Length > 0)
                        {
                            secureString.RemoveAt(secureString.Length - 1);
                        }
                        else if (!char.IsControl(keyInfo.KeyChar))
                        {
                            secureString.AppendChar(keyInfo.KeyChar);
                        }
                    }
                    while (key != ConsoleKey.Enter);

                    Console.WriteLine();
                }
                else
                {
                    foreach (var c in value)
                    {
                        secureString.AppendChar(c);
                    }
                }
            }

            Console.WriteLine();
            return secureString;
        }
    }
}