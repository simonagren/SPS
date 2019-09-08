export interface EmitObject {
    result?: string;
    room: string;
    conversationId?: string;
}

export class SocketHelper {
    public emitStart = (room: string, socket: SocketIOClient.Socket, conversationId?: string) => {
            const response: EmitObject = { room: room, conversationId: conversationId ? conversationId : '' };
            socket.emit('start', response);
    }
    
    public emitUpdate = (complete: boolean, result: string, room: string, socket: SocketIOClient.Socket, conversationId?: string) => {
            // const status: string = complete == true ? "Complete" : "InProcess";
            const response: EmitObject = { room: room, result: result, conversationId: conversationId ? conversationId : ''  };
            
            if (complete === true) {
                socket.emit('complete', response);
            } else {
                socket.emit('update', response);
            }
    }
}