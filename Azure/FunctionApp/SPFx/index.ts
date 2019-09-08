import * as df from "durable-functions"
import { IQueueObj } from "../QueueTrigger";

const orchestrator = df.orchestrator(function* (context) {
    const inputs: IQueueObj = context.bindings.context.input;
    const outputs = [];
    
    // Start provisioning
    outputs.push(yield context.df.callActivity("EmitStart", inputs));

    // Convert to Hub Site
    outputs.push(yield context.df.callActivity("EstablishHubsite", inputs));
    
    // Create Lists
    outputs.push(yield context.df.callActivity("CreateLists", inputs));
    
    // Complete provisioning
    outputs.push(yield context.df.callActivity("EmitComplete", inputs));

    return outputs;
});

export default orchestrator;
