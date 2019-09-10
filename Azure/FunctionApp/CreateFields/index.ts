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
        service.emitUpdate(false, "Creating fields", room, socket, conversationId);

        // Import module
        ps.addCommand('Import-Module C:/Users/sagren/dev/sps/azure/FunctionApp/SharePointPnPPowerShellOnline/3.12.1908.1/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');        
        // ps.addCommand('Import-Module D:/Home/site/wwwroot/modules/SharePointPnPPowerShellOnline/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        ps.addCommand('$progressPreference = "silentlyContinue"');
        
        // Connect to site
        ps.addCommand(`Connect-PnPOnline -AppId ${process.env.spId} -AppSecret ${process.env.spSecret} -Url ${inputs.siteUrl}`);
        ps.addCommand('$Site = Get-PnPSite');
        let output = await ps.invoke();
        
        // Apply provisioning template
        ps.addCommand(`Apply-PnPProvisioningTemplate -Path C:/Users/sagren/dev/sps/Azure/FunctionApp/${inputs.siteType}-fields.xml`);
        // ps.addCommand(`Apply-PnPProvisioningTemplate -Path D:/home/site/wwwroot/FunctionApp/${inputs.siteType}-fields.xml`);
        
        output = await ps.invoke();
        context.log(output);
        
        // Emit update
        service.emitUpdate(false, "Successfully created fields", room, socket, conversationId);

        return "Created fields";

    } catch (err) {
        context.log(err.message);
        await ps.dispose();

        return err.message;
    }

};

export default activityFunction;
