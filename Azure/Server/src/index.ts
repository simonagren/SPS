import express from "express";
import http from "http";
import socketio from "socket.io";

const app = express();
const server = new http.Server(app);
const io = socketio(server);
const port = process.env.port || 3000;

server.listen(port, () => {
    // tslint:disable-next-line:no-console
    console.log("listening on *:" + port);
});

app.get("/", (req, res) => {
    res.sendFile(__dirname + "/index.html");
  });

io.on("connection", (socket) => {

    socket.on("room", (room: string) => {
        socket.join(room);
    });

    socket.on("start", (data: IEmitObject) => {
        socket.in(data.room).emit("startProvisioning", data);
    });

    socket.on("update", (data: IEmitObject) => {
        socket.in(data.room).emit("provisioningUpdate", data);
    });

    socket.on("complete", (data: IEmitObject) => {
        socket.in(data.room).emit("provisioningComplete", data);
    });

});

export interface IEmitObject {
    result: string;
    room: string;
    conversationId?: string;
}
