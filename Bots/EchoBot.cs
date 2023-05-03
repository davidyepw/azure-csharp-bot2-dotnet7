// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
//
// Generated with EchoBot .NET Template version v4.17.1

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.Web.Administration;

namespace EchoBot.Bots
{
    public class EchoBot : ActivityHandler
    {
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            Console.WriteLine("**************on message start*********");
            if(turnContext.Activity.Text.Contains("listarAppPools"))
            {
            String remoteIISServer = "";
            remoteIISServer = turnContext.Activity.Text.Substring(turnContext.Activity.Text.IndexOf(" ")+1);
            Console.WriteLine(remoteIISServer);
            InitialSessionState initial = InitialSessionState.CreateDefault();
            initial.ExecutionPolicy = 0;
            //Server: C:\Windows\System32\WindowsPowerShell\v1.0\Modules\IISAdministration
            //initial.ImportPSModule(new string[] {"C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\Modules\\IISAdministration\\IISAdministration.psd1"} );

            //Desktop: C:\Program Files\WindowsPowerShell\Modules\IISAdministration\1.1.0.0
            //initial.ImportPSModule(new string[] {"C:\\Program Files\\WindowsPowerShell\\Modules\\IISAdministration\\1.1.0.0\\IISAdministration.psd1"} );
            Runspace runspace = RunspaceFactory.CreateRunspace(initial);
            runspace.Open(); 
            PowerShell ps = PowerShell.Create();
            ps.Runspace = runspace;

            
            ps.Commands.AddCommand("invoke-command")
                .AddParameter("ComputerName", remoteIISServer)
                .AddParameter("ScriptBlock", ScriptBlock.Create("Get-IISAppPool"));
                       
            
            /*
            ps.Commands.AddCommand("get-process");
            */
            var cadena = "";
            foreach(PSObject item in ps.Invoke()){
                Console.WriteLine("inside foreach");
                Console.WriteLine("item value:" +item.Members["Name"].Value.ToString());
                cadena += item.Members["Name"].Value.ToString() + " " + item.Members["State"].Value.ToString() + "\n\n";
            }
            var replyText = cadena;
            await turnContext.SendActivityAsync(MessageFactory.Text(replyText, replyText), cancellationToken);
            }
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            
            
            var welcomeText = "Comandos disponibles: \r\n listarAppPools SERVERNAME \r\n listarWebsites SERVERNAME";
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(welcomeText, welcomeText), cancellationToken);
                }
            }
        }
    }
}
