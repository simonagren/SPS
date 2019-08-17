import * as React from "react";
import { Panel, PanelType, Icon } from "office-ui-fabric-react";
import PanelMessage from "./PanelMessage";

export interface IMessagePanelProps {
    showPanel: boolean;
    hidePanel: () => void;
    messages?: string[];
}

export interface IMessagePanelState {
    currentMessages?: string[];
}

export default class MessagePanel extends React.Component<IMessagePanelProps, IMessagePanelState>
{
    constructor(props: IMessagePanelProps) {
        super(props);

        this.state = {
            currentMessages: []
        }

    }

    public componentDidUpdate(prevProps: Readonly<IMessagePanelProps>, prevState: Readonly<IMessagePanelState>, snapShot?: any): void {
        if (this.props.messages === prevProps.messages) {
            return;
        }

        this.setState({ currentMessages: this.props.messages })
    }

    public render(): React.ReactElement<IMessagePanelProps> {
        const { showPanel, hidePanel } = this.props;
        const { currentMessages } = this.state;

        return (
            <div>
                <Panel
                    isOpen={showPanel}
                    type={PanelType.customNear}
                    customWidth={"400px"}
                    onDismiss={hidePanel}
                    headerText="Panel - Small, left-aligned, fixed"
                >
                    {currentMessages &&
                        currentMessages.map(msg =>
                            <PanelMessage message={msg} />
                        )
                    }
                </Panel>
            </div>
        );
    }
}

