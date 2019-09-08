import * as df from "durable-functions"
import { IQueueObj } from "../QueueTrigger";

const orchestrator = df.orchestrator(function* (context) {
    const inputs: IQueueObj = context.bindings.context.input;
    const outputs = [];

    // Start provisioning
    outputs.push(yield context.df.callActivity("EmitStart", inputs));

    // Create site
    outputs.push(yield context.df.callActivity("CreateTeamSite", inputs));

    // Set logo
    outputs.push(yield context.df.callActivity("SetExternal", inputs));

    // Complete provisioning
    outputs.push(yield context.df.callActivity("EmitComplete", inputs));

    return outputs;
});

export default orchestrator;
