using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.IO;


namespace getUsers
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string domainName = "nextend.local";
            string[] propertiesToLoad = { "samaccountname", "givenname", "sn", "mail" };

            using (DirectoryEntry root = new DirectoryEntry("LDAP://" + domainName))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(root))
                {
                    searcher.PageSize = 1000;
                    searcher.Filter = "(objectCategory=person)";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.AddRange(propertiesToLoad);

                    SearchResultCollection results = searcher.FindAll();
                    string csvFilePath = "users.csv";
                    using (StreamWriter writer = new StreamWriter(csvFilePath))
                    {
                        writer.WriteLine("User Name,First Name,Last Name,Email");
                        foreach (SearchResult result in results)
                        {
                            string userName = null;
                            if (result.Properties.Contains("samaccountname") && result.Properties["samaccountname"].Count > 0)
                            {
                                userName = result.Properties["samaccountname"][0].ToString();
                            }

                            string firstName = null;
                            if (result.Properties.Contains("givenname") && result.Properties["givenname"].Count > 0)
                            {
                                firstName = result.Properties["givenname"][0].ToString();
                            }

                            string lastName = null;
                            if (result.Properties.Contains("sn") && result.Properties["sn"].Count > 0)
                            {
                                lastName = result.Properties["sn"][0].ToString();
                            }

                            string email = null;
                            if (result.Properties.Contains("mail") && result.Properties["mail"].Count > 0)
                            {
                                email = result.Properties["mail"][0].ToString();
                            }

                            writer.WriteLine(string.Format("{0},{1},{2},{3}", userName, firstName, lastName, email));
                        }
                    }
                    Console.WriteLine("Results exported to CSV file: " + csvFilePath);
                }    
            }

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();


        }
    }
}
