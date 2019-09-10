import { AzureFunction, Context } from "@azure/functions"
import * as shell from "node-powershell";
import { IQueueObj } from "../QueueTrigger";
import { SocketHelper } from "../Helper/helper";
import * as io from "socket.io-client";

const activityFunction: AzureFunction = async function (context: Context, queueObj: IQueueObj): Promise<string> {

    const inputs = queueObj;
    const service = new SocketHelper();
    const socket: SocketIOClient.Socket = io('https://spsexpress.azurewebsites.net/');
    
    const room: string = inputs.isBot ? 'bot' : inputs.siteUrl;
    const conversationId: string = inputs.conversationId ? inputs.conversationId : '';
    const ps = new shell({
        executionPolicy: 'Bypass',
        noProfile: true
    });

    try {

        // Emit update
        service.emitUpdate(false, "Converting to Hub Site", room, socket, conversationId);

        // Import module
        ps.addCommand('Import-Module C:/Users/sagren/dev/sps/azure/FunctionApp/SharePointPnPPowerShellOnline/3.12.1908.1/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        // ps.addCommand('Import-Module D:/Home/site/wwwroot/modules/SharePointPnPPowerShellOnline/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        ps.addCommand('$progressPreference = "silentlyContinue"');

        // Connect to site
        ps.addCommand(`$encpassword = convertto-securestring -String ${process.env.adminPass} -AsPlainText -Force`);
        ps.addCommand(`$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ${process.env.adminUser}, $encpassword`)
        ps.addCommand(`Connect-PnPOnline -credentials $cred -Url ${inputs.siteUrl}`);
        ps.addCommand('$site = Get-PnPSite');
        let output = await ps.invoke();
        
        // Register as an Hub Site
        ps.addCommand('Register-PnPHubSite -Site $site');
        output = await ps.invoke();        

        // Emit update
        service.emitUpdate(false, "Successfully converted to Hub Site", room, socket, conversationId);

        return "Converted site to Hub Site";

    } catch (err) {
        context.log(err.message);
        await ps.dispose();

        return err.message;
    }

};

export default activityFunction;
