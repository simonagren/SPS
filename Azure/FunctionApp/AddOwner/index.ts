import { AzureFunction, Context } from "@azure/functions"
import * as shell from "node-powershell";
import * as io from "socket.io-client";
import { SocketHelper } from "../Helper/helper";
import { IQueueObj } from "../QueueTrigger";


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
        service.emitUpdate(false, "Adding owner to owners group", room, socket, conversationId);

        // Import module
        ps.addCommand('Import-Module C:/Users/sagren/dev/sps/azure/FunctionApp/SharePointPnPPowerShellOnline/3.12.1908.1/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');        
        // ps.addCommand('Import-Module D:/Home/site/wwwroot/modules/SharePointPnPPowerShellOnline/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        ps.addCommand('$progressPreference = "silentlyContinue"');
        
        // Connect to site
        ps.addCommand(`Connect-PnPOnline -AppId ${process.env.spId} -AppSecret ${process.env.spSecret} -Url ${inputs.siteUrl}`);
        let output = await ps.invoke();
        
        // Get owner-group and add owner
        ps.addCommand('$ownergroup = Get-PnPGroup | where {$_.Title -like "*Owners*"}');
        ps.addCommand(`Add-PnPUserToGroup -LoginName "${inputs.siteOwner}" -Identity $ownergroup.Id`);
        
        // Execute PS
        output = await ps.invoke();
        context.log(output);
        
        // Emit update
        service.emitUpdate(false, "Successfully added owner", room, socket, conversationId);

        return "Added owner";

    } catch (err) {
        context.log(err.message);
        await ps.dispose();

        return err.mesage;
    }

};

export default activityFunction;
