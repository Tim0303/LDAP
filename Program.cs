using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Linq;
using System.Net;

namespace LDAP
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string account = "";
            string password = "";
            var domainName = "";
            var ldapServer = @"";

            LDAPAuthenticate(ldapServer, domainName, account, password);

            UserAsync(ldapServer, domainName, account, password);

            // 等待使用者輸入任意鍵後結束程式
            Console.WriteLine("等待使用者輸入任意鍵後結束程式");
            Console.ReadKey();
        }

        /// <summary>
        /// LDAP 驗證
        /// </summary>
        /// <param name="ldapServer"></param>
        /// <param name="domain"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static bool LDAPAuthenticate(string ldapServer, string domain, string username, string password)
        {
            var ldapDirectoryIdentifier = new LdapDirectoryIdentifier(ldapServer, 389, true, false);
            // 建立 LDAP 連線
            using (var ldapConnection = new LdapConnection(ldapDirectoryIdentifier))
            {
                ldapConnection.AuthType = AuthType.Negotiate;

                try
                {
                    var networkCredential = new NetworkCredential(username, password, domain);
                    // 嘗試綁定以驗證用戶
                    ldapConnection.Bind(networkCredential);
                    Console.WriteLine($"username password domain 驗證成功");

                    // 取得使用者資訊
                    var searchRequest = new SearchRequest();
                    searchRequest.Filter = $"(sAMAccountName={username})";
                    searchRequest.Scope = System.DirectoryServices.Protocols.SearchScope.Subtree;
                    searchRequest.DistinguishedName = "DC=xxx,DC=xxx,DC=xx";
                    searchRequest.Attributes.Add("distinguishedName");
                    searchRequest.Attributes.Add("mail");
                    searchRequest.Attributes.Add("displayName");
                    searchRequest.Attributes.Add("title");

                    var response = (SearchResponse)ldapConnection.SendRequest(searchRequest);

                    response.Entries.Cast<SearchResultEntry>().ToList().ForEach(entry =>
                    {
                        Console.WriteLine($"distinguishedName: {entry.Attributes["distinguishedName"][0]}");
                        Console.WriteLine($"mail: {entry.Attributes["mail"][0]}");
                        Console.WriteLine($"displayName: {entry.Attributes["displayName"][0]}");
                        Console.WriteLine($"title: {entry.Attributes["title"][0]}");
                    });
                }
                catch (LdapException ex)
                {
                    Console.WriteLine($"username password domain驗證失敗：{ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"未知錯誤：{ex.Message}");
                }

                return true;
            }
        }

        /// <summary>
        /// 同步使用者資訊
        /// </summary>
        /// <param name="ldapServer"></param>
        /// <param name="domain"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public static void UserAsync(string ldapServer, string domain, string username, string password)
        {
            try
            {
                var users = new List<User>();
                var ldapDirectoryIdentifier = new LdapDirectoryIdentifier(ldapServer, 389, true, false);

                using (var ldapConnection = new LdapConnection(ldapDirectoryIdentifier))
                {
                    ldapConnection.AuthType = AuthType.Negotiate;
                    ldapConnection.Credential = new NetworkCredential(username, password, domain);
                    ldapConnection.Bind();
                    string searchBase = "OU=xx,OU=xx,OU=xxx,DC=xxx,DC=xxx,DC=xx";
                    string searchFilter = "(objectClass=user)";
                    SearchRequest searchRequest = new SearchRequest(
                        searchBase,
                        searchFilter,
                        SearchScope.Subtree,
                        new string[] { "distinguishedName", "sAMAccountName", "displayName", "mail", "title" }
                    );
                    SearchResponse searchResponse = (SearchResponse)ldapConnection.SendRequest(searchRequest);

                    // 處理結果
                    foreach (SearchResultEntry entry in searchResponse.Entries)
                    {
                        string accountName = entry.Attributes["sAMAccountName"]?[0]?.ToString();
                        string displayName = entry.Attributes["displayName"]?[0]?.ToString();
                        string email = entry.Attributes["mail"]?[0]?.ToString();
                        string distinguishedName = entry.Attributes["distinguishedName"]?[0]?.ToString();
                        string title = entry.Attributes["title"]?[0]?.ToString();

                        // Console.WriteLine($"帳號：{accountName}, 姓名：{displayName}, Email：{email}, distinguishedName:{distinguishedName}, title:{title}");

                        // 預處理distinguishedName 把OU的值取出
                        var ouValues = distinguishedName.Split(',')
                            .Where(part => part.StartsWith("OU="))
                            .Select(part => part.Substring(3))
                            .ToList();

                        users.Add(new User
                        {
                            Id = accountName,
                            DisplayName = displayName,
                            Email = email,
                            Org = ouValues[1],
                            JobTitle = title,
                            Region = ouValues[0],
                        });

                        // 在這裡同步到DB或其他邏輯
                    }
                }

                // 輸出結果
                users.ForEach(user =>
                {
                    Console.WriteLine($"帳號：{user.Id}, 姓名：{user.DisplayName}, Email：{user.Email}, 組織：{user.Org}, 職稱：{user.JobTitle}, 區域：{user.Region}");
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"錯誤: {ex.Message}");
            }
        }
    }

    public class User
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string Org { get; set; }
        public string JobTitle { get; set; }
        public string Region { get; set; }
    }
}