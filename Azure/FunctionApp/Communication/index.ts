import * as df from "durable-functions"
import { IQueueObj } from "../QueueTrigger";

const orchestrator = df.orchestrator(function* (context) {
    const inputs: IQueueObj = context.bindings.context.input;
    const outputs = [];

    // Start provisioning
    outputs.push(yield context.df.callActivity("EmitStart", inputs));
    
    // Create site
    outputs.push(yield context.df.callActivity("CreateCommSite", inputs));

    // Add Theme
    outputs.push(yield context.df.callActivity("AddTheme", inputs));

    // Add owner to group
    outputs.push(yield context.df.callActivity("AddOwner", inputs));

    // Add all company
    outputs.push(yield context.df.callActivity("AddAllCompany", inputs));

    // Complete provisioning
    outputs.push(yield context.df.callActivity("EmitComplete", inputs));

    return outputs;
});

export default orchestrator;
