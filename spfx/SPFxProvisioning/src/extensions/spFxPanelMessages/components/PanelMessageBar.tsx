import * as React from "react";
import * as ReactDOM from "react-dom";
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { IMessagePanelProps } from './MessagePanel';
import MessagePanel from './MessagePanel';
import { DefaultButton, autobind } from "office-ui-fabric-react";

//import {ISiteArchivedMessageBarProps} from "./ISiteArchivedMessageBar";

export interface IPanelMessageBarProps {
}

export interface IPanelMessageBarState {
    showPanel: boolean;
    currentMessage: any;
    panelMessages: string[];
    showMessageBar: boolean;
}


export default class PanelMessageBar extends React.Component<IPanelMessageBarProps, IPanelMessageBarState>
{
    constructor(props: IPanelMessageBarProps) {
        super(props);

        this.state = {
            showPanel: false,
            currentMessage: "PnP Templates are being applied. If you want to see the process, press here",
            panelMessages: [],
            showMessageBar: false,
        }

        this._registerSocketIo();

    }

    public componentDidMount(): void {
        this._showMessageBarIntro();
        setTimeout(this._addMessage, 5000);
        setTimeout(this._addMessage, 10000);
        setTimeout(this._addMessage, 15000);        

      }

    public render(): React.ReactElement<IPanelMessageBarProps> {
        const { showPanel, currentMessage, panelMessages, showMessageBar } = this.state;

        return (
            <div>
                {showMessageBar &&
                    <div>
                        <MessageBar messageBarType={MessageBarType.info}>
                            {currentMessage}
                            <DefaultButton secondaryText="Opens the Sample Panel" onClick={this._showPanel} text="Open Panel" />
                            <DefaultButton secondaryText="Adds a message" onClick={this._addMessage} text="Add message" />
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
        // const socket = io("url");

        // socket.on('connect', () => {

        //   socket.emit('room', siteUrl);

        //   socket.on('provisioningReady', (data) => {
        //     this._showMessageBar(data);
        //   });

        //   socket.on('provisioningComplete', (data) => {
        //     this._createNotification(data);
        //   });

        // })
    }

    @autobind
    private _showMessageBarIntro(): void {
        this.setState({ 
            showMessageBar: true, 
            // message: "Test" && <DefaultButton secondaryText="Opens the Sample Panel" onClick={this._showPanel} text="Open Panel" />

        })
    }

    @autobind
    private _showMessageBarOutro(): void {
        this.setState({ showPanel: true })
    }

    @autobind
    private _hideMessageBar(): void {
        this.setState({ showPanel: false })
    }

    @autobind
    private _addMessage(): void {

        this.setState(prevState => ({
            panelMessages: [...prevState.panelMessages, "A new message"]
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