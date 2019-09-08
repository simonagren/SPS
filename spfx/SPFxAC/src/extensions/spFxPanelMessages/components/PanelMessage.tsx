import * as React from "react";
import * as ReactDOM from "react-dom";
import { Icon } from "office-ui-fabric-react";


export interface IPanelMessageProps {
    message: string;
}

export interface IPanelMessageState {
    done: boolean;
}


export default class PanelMessage extends React.Component<IPanelMessageProps, IPanelMessageState>
{
    constructor(props: IPanelMessageProps) {
        super(props);

        this.state = {
            done: false
        }
    }

    public render(): React.ReactElement<IPanelMessageProps> {
        const { message } = this.props;
        const { done } = this.state;

        return (
            <div style={{ margin: "10px" }} className="ms-siteScriptProgress-actionHeader">
                <Icon iconName="CheckMark" className="ms-siteScriptProgress-actionIcon ms-siteScriptProgress-successIcon root-559" >
                </Icon>
                <div className="ms-siteScriptProgress-actionTitle ms-siteScriptProgress-activeActionTitle">
                    {message}
                </div>
            </div>


        );
    }

}
