import * as df from "durable-functions"
import * as io from "socket.io-client";
import { IFunctionContext, DurableOrchestrationContext } from "durable-functions/lib/src/classes";
import { Socket } from "net";

const orchestrator = df.orchestrator(function* (context) {
    const outputs = [];
    const siteUrl: string = context.bindings.siteUrl;
    const socket: SocketIOClient.Socket = io('');
    
    // Replace "Hello" with the name of your Durable Activity Function.

    _emitUpdate(false, "Converting to hub site", context, siteUrl, socket);
    outputs.push(yield context.df.callActivity("EstablishHubsite", siteUrl));
    _emitUpdate(false, "Successfully created hub site", context, siteUrl, socket);

    _emitUpdate(false, "Creating lists", context, siteUrl, socket);
    outputs.push(yield context.df.callActivity("CreateLists", context.bindings));
    _emitUpdate(false, "Successfully created lists", context, siteUrl, socket);

    _emitUpdate(false, "Creating new home page", context, siteUrl, socket);
    outputs.push(yield context.df.callActivity("CreateHomePage", context.bindings));
    _emitUpdate(false, "Successfully created home page", context, siteUrl, socket);

    return outputs;
});

const _emitUpdate = (complete: boolean, result: string, context: IFunctionContext, siteUrl: string, socket: SocketIOClient.Socket) => {
    if (!context.df.isReplaying) {
        const status: string = complete == true ? "Complete" : "InProcess";
        const response: EmitObject = { result: result, status: status, siteUrl: siteUrl };
        socket.emit('root', response);
    }
}

export default orchestrator;

export interface EmitObject {
    result: string;
    status: string;
    siteUrl: string;
}