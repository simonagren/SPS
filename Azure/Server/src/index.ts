import express from "express";
import http from "http";
import socketio from "socket.io";

const app = express();
const server = new http.Server(app);
const io = socketio(server);

io.on("connection", (socket) => {

    socket.on("message", (data) => {
        if (data === "update") {
            socket.broadcast.emit("updateFromServer", data);
        } else if (data === "done") {
            socket.broadcast.emit("doneFromServer", data);
        }
    });

});

server.listen(3000, () => {
    // tslint:disable-next-line:no-console
    console.log("listening on *:3000");
});
