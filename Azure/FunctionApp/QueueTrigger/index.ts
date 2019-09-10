import * as df from "durable-functions";
import { AzureFunction, Context } from "@azure/functions";

const queueTrigger: AzureFunction = async function (context: Context, myQueueItem: IQueueObj): Promise<void> {
    
    const client = df.getClient(context);
    
    // start a new orchestration based on site type
    const instanceId = await client.startNew(myQueueItem.siteType, undefined, myQueueItem)

    context.log(`Started orchestration with ID = '${instanceId}'.`);
    context.log('Queue trigger function processed work item', myQueueItem);
};

export default queueTrigger;

export interface IQueueObj {
    siteUrl: string;
    siteTitle?: string;
    siteOwner?: string;
    siteType: string;
    siteAlias?: string;
    siteCreationType?: string;
    conversationId?: string;
    isBot?: boolean;
}