import * as React from "react";
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import MessagePanel from './MessagePanel';
import { autobind } from "@uifabric/utilities";

export interface IPanelMessageBarProps {
    siteUrl: string;
    socket: SocketIOClient.Socket;
    
}

export interface IPanelMessageBarState {
    showPanel: boolean;
    currentMessage: any;
    panelMessages: string[];
    showMessageBar: boolean;
    messageBarType: MessageBarType;
}

export default class PanelMessageBar extends React.Component<IPanelMessageBarProps, IPanelMessageBarState>
{
    constructor(props: IPanelMessageBarProps) {
        super(props);

        this.state = {
            showPanel: false,
            currentMessage: null,
            panelMessages: [],
            showMessageBar: false,
            messageBarType: MessageBarType.info,
        }


    }

    public componentDidMount(): void {        
        this._registerSocketIo();
    }

    public render(): React.ReactElement<IPanelMessageBarProps> {
        const { showPanel, currentMessage, panelMessages, showMessageBar, messageBarType } = this.state;

        return (
            <div>
                {showMessageBar &&
                    <div>
                        <MessageBar messageBarType={messageBarType}>
                            {currentMessage}
                        </MessageBar>
                        <MessagePanel
                            messages={panelMessages}
                            showPanel={showPanel}
                            hidePanel={this._hidePanel}
                        />
                    </div>
                }

            </div>


        );
    }

    private _registerSocketIo(): void {
        const { socket } = this.props;
        
        socket.on('startProvisioning', () => {
            this._showMessageBarIntro();
        });

        socket.on('provisioningUpdate', (data) => {
            this._addMessage(data.result);
        });

        socket.on('provisioningComplete', () => {
            this._showMessageBarComplete();
        });
    }

    @autobind
    private _showMessageBarIntro(): void {
        this.setState({
            showMessageBar: true,
            currentMessage: (<div>PnP Templates are being applied. If you want to see the process, press <b style={{ cursor: "pointer", textDecoration: "underline"  }} onClick={this._showPanel}>here</b></div>)
        });
    }

    @autobind
    private _showMessageBarComplete(): void {
        this.setState({
            messageBarType: MessageBarType.success, 
            currentMessage: (<div>Provisioning is complete. If you want to see how the site looks press <b style={{ cursor: "pointer", textDecoration: "underline"  }} onClick={() => location.reload()}>here</b></div>)
         });
    }

    @autobind
    private _addMessage(message: string): void {
        this.setState(prevState => ({
            panelMessages: [...prevState.panelMessages, message]
        }));

    }

    @autobind
    private _showPanel(): void {
        this.setState({ showPanel: true })
    }

    @autobind
    private _hidePanel(): void {
        this.setState({ showPanel: false })
    }


}