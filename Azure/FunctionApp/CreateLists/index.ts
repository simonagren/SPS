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

        service.emitUpdate(false, "Creating lists", room, socket, conversationId);

        ps.addCommand('Import-Module C:/Users/sagren/Desktop/SharePointPnPPowerShellOnline/3.12.1908.1/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');        
        // ps.addCommand('Import-Module D:/Home/site/wwwroot/modules/SharePointPnPPowerShellOnline/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        ps.addCommand('$progressPreference = "silentlyContinue"');
        ps.addCommand(`Connect-PnPOnline -AppId ${process.env.spId} -AppSecret ${process.env.spSecret} -Url ${inputs.siteUrl}`);
        ps.addCommand('$Site = Get-PnPSite');
        let output = await ps.invoke();
        context.log(`Connected to site ${output}`);
        context.log("Applying provisioning template");

        ps.addCommand(`Apply-PnPProvisioningTemplate -Path C:/Users/sagren/dev/sps/Azure/FunctionApp/${inputs.siteType}-lists.xml`);
        // ps.addCommand(`Apply-PnPProvisioningTemplate -Path D:/home/site/wwwroot/FunctionApp/${inputs.siteType}-lists.xml`);
        
        output = await ps.invoke();
        context.log(output);
        
        service.emitUpdate(false, "Successfully created lists", room, socket, conversationId);

        return "Created lists";

    } catch (err) {
        context.log(err);
        await ps.dispose();

        return `Created lists at: ${context.bindings.siteUrl}`;
    }

};

export default activityFunction;
