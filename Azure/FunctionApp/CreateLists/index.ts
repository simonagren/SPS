import { AzureFunction, Context } from "@azure/functions"
import * as shell from "node-powershell";

const activityFunction: AzureFunction = async function (context: Context): Promise<string> {
    const ps = new shell({
        executionPolicy: 'Bypass',
        noProfile: true
    });

    try {
        
        ps.addCommand('Import-Module D:/Home/site/wwwroot/modules/SharePointPnPPowerShellOnline/SharePointPnPPowerShellOnline.psd1 -WarningAction SilentlyContinue');
        ps.addCommand('$progressPreference = "silentlyContinue"');
        ps.addCommand(`Connect-PnPOnline -AppId ${process.env.spId} -AppSecret ${process.env.spSecret} -Url ${process.env.spTenantUrl}/sites/test`);
        ps.addCommand('$Web = Get-PnPWeb');
        ps.addCommand('Get-PnPProperty -ClientObject $Web -Property Title');
        let output = await ps.invoke();
        context.log(`Connected to site ${output}`);
        context.log("Applying provisioning template");

        ps.addCommand('Apply-PnPProvisioningTemplate -Path test.xml');
        output = await ps.invoke();
        context.log(output);
        context.log("Done!")

        return "Gick bra";

    } catch (err) {
        context.log(err);
        await ps.dispose();

        return `Created lists at: ${context.bindings.siteUrl}`;
    }

};

export default activityFunction;
