import * as df from "durable-functions";
import { AzureFunction, Context } from "@azure/functions";

const queueTrigger: AzureFunction = async function (context: Context, myQueueItem: QueueObj): Promise<void> {
    
    context.log('Item template: ', myQueueItem.template);
    context.log('Item url: ', myQueueItem.webUrl);
    
    const client = df.getClient(context);
    const instanceId = await client.startNew(myQueueItem.template, undefined, myQueueItem)
    
    context.log(`Started orchestration with ID = '${instanceId}'.`);
    context.log('Queue trigger function processed work item', myQueueItem);
};

export default queueTrigger;

export interface QueueObj {
    webUrl: string;
    template: string;
}