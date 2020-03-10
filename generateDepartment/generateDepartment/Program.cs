
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace generateDepartment
{
    public class Program
    {

       
       static GraphServiceClient connectToGraphAPI(string tId, string cId, string cSecret)
        {

            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var confidentialClient = ConfidentialClientApplicationBuilder
          .Create(cId)
          .WithAuthority($"https://login.microsoftonline.com/murphyoil.onmicrosoft.com/v2.0")
          .WithClientSecret(cSecret)
          .Build();

            GraphServiceClient gsClient =
            new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                var authResult = await confidentialClient
        .AcquireTokenForClient(scopes)
        .ExecuteAsync();

                // Add the access token in the Authorization header of the API request.
                requestMessage.Headers.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            })
            );


            return gsClient;
        }



        static async System.Threading.Tasks.Task Main(string[] args)
        {

            GraphServiceClient graphServiceClient1 = connectToGraphAPI("murphyoil.onmicrosoft.com", "e528d78a-7851-4202-9473-e89542531c19", "H1H[t6rp78sB=v/dUvvicbCDrvyrtHk]");
            var usersForDep = await graphServiceClient1.Users.Request().Select(e => new
              {
                    e.DisplayName,
                    e.UserPrincipalName,
                    e.Department,
                e.BusinessPhones,
                e.MobilePhone
            }).GetAsync();

            string departmentListFilePath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory.ToString(), "unique1DepartmentList.txt");

            string uniqueDepartmentListFilePath= System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory.ToString(), "uniqueFinalDeptList.txt");
            List<string> departmentlist = new List<string>();

            //users next page to get all users
            do
            {
                foreach (User user in usersForDep)

                {
                    //var Peoples = await graphServiceClient.Users[user.Id].People.Request().GetAsync();

                   if (user.Department != null && user.UserPrincipalName != null)
                    {
                        //combinedText = combinedText + "Department: " + People.Department + linebreaker;

                        try
                        {
                            departmentlist.Add(user.Department.ToString());
                            System.IO.File.AppendAllText("uniqueDept.txt", user.Department.ToString() + Environment.NewLine);
                        }

                        catch (Exception ex)
                        {

                            Console.WriteLine(ex.Message);
                        }


                    }

                }

            } while (usersForDep.NextPageRequest != null && (usersForDep = await usersForDep.NextPageRequest.GetAsync()).Count > 0);

            // List<string> allLinesText = System.IO.File.ReadAllLines(departmentListFilePath).ToList();





            List<string> uniqueDeptList = departmentlist.Distinct().ToList();

            foreach (string line in uniqueDeptList)

            {

                System.IO.File.AppendAllText("uniqueFinal.txt", line + Environment.NewLine);

            }




        }


     


    }
}