import * as df from "durable-functions"
import { IQueueObj } from "../QueueTrigger";

const orchestrator = df.orchestrator(function* (context) {
    const inputs: IQueueObj = context.bindings.context.input;
    const outputs = [];
    
    // Start provisioning
    outputs.push(yield context.df.callActivity("EmitStart", inputs));

    // Create site
    outputs.push(yield context.df.callActivity("CreateTeamSite", inputs));

    // Add owner to group
    outputs.push(yield context.df.callActivity("AddOwner", inputs));

    // Add fields
    outputs.push(yield context.df.callActivity("CreateFields", inputs));

    // Add content types
    outputs.push(yield context.df.callActivity("CreateCTypes", inputs));

    // Add lists
    outputs.push(yield context.df.callActivity("CreateLists", inputs));

    // Complete provisioning
    outputs.push(yield context.df.callActivity("EmitComplete", inputs));

    return outputs;
});

export default orchestrator;
