using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GraphClient
{
    class Program
    {
        static String client_id = "";
        static String tenant_id = "";
        static String client_secret = "";
        static String[] scopes = { "https://graph.microsoft.com/.default" };
        static String output_path = "c:\\temp\\SignIns_Report.csv";

        static IConfidentialClientApplication _pca = null;
        private static IConfidentialClientApplication Pca
        {
            get
            {
                if(_pca == null )
                {
                    _pca = ConfidentialClientApplicationBuilder
                        .Create( client_id )
                        .WithClientSecret( client_secret )
                        .WithTenantId( tenant_id )
                        .Build();
                }
                return _pca;
            }
        }

        static ClientCredentialProvider _authProvider = null;
        private static ClientCredentialProvider AuthProvider
        {
            get
            {
                if(_authProvider == null )
                {
                    _authProvider = new ClientCredentialProvider( Pca );
                }
                return _authProvider;
            }
        }

        static GraphServiceClient _graphClient = null;
        private static GraphServiceClient GraphClient
        {
            get
            {
                if(_graphClient == null )
                {
                    _graphClient = new GraphServiceClient( AuthProvider );
                }
                return _graphClient;
            }
        }

        static void Main(string[] args)
        {
            Get_Users_SignIn_Logs().Wait();

            //Create_B2CUser( "Joe Smoe", "Joe", "Smoe", "joesmoe", "joesmoe@rayheld.com" ).Wait();

            //Set_AppRoles().Wait();

            Console.WriteLine( $"\nPress any key to close..." );
            Console.ReadKey();
        }

        static async Task Get_Me()
        {
            User user = null;
            user = await GraphClient.Me.Request().GetAsync();
            Console.WriteLine( $"Display Name = {user.DisplayName}\nEmail = {user.Mail}" );
        }

        static async Task Get_Users_SignIn_Logs( )
        {
            List<User> user_pages = new List<User>();
            IGraphServiceUsersCollectionPage user_page = null;

            try
            {
                user_page = await GraphClient.Users
                    .Request()
                    //.Top ( 1 )
                    //.Filter("mail eq 'user@testtenant.com'") --> To Get just one user
                    .GetAsync();

                if ( user_page != null )
                {
                    user_pages.AddRange( user_page );
                    while ( user_page.NextPageRequest != null ) {
                        user_page = await user_page.NextPageRequest.GetAsync();
                        user_pages.AddRange( user_page );
                    }
                }
            } catch ( Exception ex )
            {
                Console.WriteLine( $"Exception: {ex.Message}" );
            }

            Console.WriteLine( $"---------------------------\nTotal users {user_pages.Count}" );

            List<SignIn> user_signins = new List<SignIn>();

            foreach ( User u in user_pages )
            {
                Console.WriteLine( $"Getting signins for user: {u.UserPrincipalName}" );

                IAuditLogRootSignInsCollectionPage signin_page;
                try
                {
                    Int32 signin_count = 0;
                    signin_page = await GraphClient.AuditLogs.SignIns
                        .Request()
                        // needs filter for the user
                        .Filter( $"userId eq '{ u.Id }'" )
                        //.Top( 1 )
                        .WithMaxRetry(5) // default is 3 -- this is graph client handling 429 errors internally
                        .GetAsync();
                    if (signin_page != null )
                    {
                        signin_count += signin_page.Count;
                        user_signins.AddRange( signin_page );
                        while ( signin_page.NextPageRequest != null )
                        {
                            signin_page = await signin_page.NextPageRequest.GetAsync();
                            user_signins.AddRange( signin_page );
                        }

                        Console.WriteLine( $"Total of {signin_count} signins found for user." );
                    }
                } catch (Exception ex )
                {
                    Console.WriteLine( $"Exception: {ex.Message}" );
                }

             
            }

            Console.WriteLine( "Saving report..." );
            Output_SignIn_Report_To_CSV( user_signins, output_path );
            
        }

        static void Output_SignIn_Report_To_CSV(List<SignIn> all_signins, string filePath )
        {
            using ( StreamWriter file = new StreamWriter( filePath ) )
            {
                Type type = all_signins[0].GetType();
                PropertyInfo[] properties = type.GetProperties();

                StringBuilder sb = new StringBuilder();

                foreach (PropertyInfo property in properties )
                {
                    if(sb.Length > 0 )
                    {
                        sb.Append( "," );
                    }
                    sb.Append($"\"{property.Name.ToString()}\"" );
                }
                file.WriteLine( sb.ToString() );

                foreach(SignIn log in all_signins )
                {
                    sb = new StringBuilder();
                    foreach (PropertyInfo property in properties )
                    {
                        if (sb.Length > 0 )
                        {
                            sb.Append( "," );
                        }
                        if(property.GetValue(log, null) != null )
                        {
                            sb.Append( $"\"{Convert_SignInLog_GraphType( property, log )}\"" );
                        } else
                        {
                            sb.Append( $"\"\"" );
                        }
                    }
                    file.WriteLine( sb.ToString() );
                }
            }
        }

        static String Convert_SignInLog_GraphType(PropertyInfo property, SignIn log )
        {
            StringBuilder sb = new StringBuilder();
            // need to perform the graph action to retrieve the value
            switch ( property.PropertyType.FullName )
            {
                case "Microsoft.Graph.SignInStatus":
                    if ( log.Status != null )
                    {
                        return "Status:" +
                            log.Status.ErrorCode != null ? log.Status.ErrorCode.ToString() : string.Empty;
                    }
                    else
                    {
                        return string.Empty;
                    }

                case "Microsoft.Graph.DeviceDetail":
                    if ( log.DeviceDetail != null )
                    {
                        return "Id:" + log.DeviceDetail.DeviceId != null ? log.DeviceDetail.DeviceId.ToString() : string.Empty;
                    } else
                    {
                        return string.Empty;
                    }
                case "Microsoft.Graph.SignInLocation":
                    if ( log.Location != null )
                    {
                        sb.Append (( log.Location.City != null ? log.Location.City.ToString() : string.Empty ) + ", " );
                        sb.Append (( log.Location.State != null ? log.Location.State.ToString() : string.Empty ) + " " );
                        sb.Append ( log.Location.CountryOrRegion != null ? log.Location.CountryOrRegion.ToString() : string.Empty );
                        return sb.ToString();
                    } else
                    {
                        return string.Empty;
                    }
                default:
                    return property.GetValue( log, null ) != null ? property.GetValue( log, null ).ToString() : string.Empty;
            }

        }

    }
}
