import * as React from 'react';
import styles from './SpFxProvisioning.module.scss';
// import { ISpFxProvisioningProps, ISpFxProvisioningState } from '.';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { ListView, SelectionMode } from "@pnp/spfx-controls-react/lib/ListView";
import { sp, Web, Site, WebPart, Items } from '@pnp/sp';
import { Guid } from '@microsoft/sp-core-library';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISpFxProvisioningProps } from './ISpFxProvisioningProps';

export default class SpFxProvisioning extends React.Component<ISpFxProvisioningProps, {}> {
// export default class SpFxProvisioning extends React.Component<ISpFxProvisioningProps, ISpFxProvisioningState> {

  
  constructor(props: ISpFxProvisioningProps) {
    super(props);

    this.state = {
      items: [],
      error: undefined,
      loading: true,
    };
  }

  public componentDidMount(): void {
    if (!this.props.listId) {
      return;
    }

    // this._loadItems();

  }

  public componentDidUpdate(prevProps: Readonly<ISpFxProvisioningProps>, {}, snapShot?: any): void {
  // public componentDidUpdate(prevProps: Readonly<ISpFxProvisioningProps>, prevState: Readonly<ISpFxProvisioningState>, snapShot?: any): void {

    if (this.props.listId === prevProps.listId) {
      return;
    }

    // this._loadItems();
  }
  
  public render(): React.ReactElement<ISpFxProvisioningProps> {
    // const { onConfigure } = this.props;
    // const needsConfiguration: boolean = !this.props.listId;
    // const { error, items, loading} = this.state;

    return (
      // <div className={ styles.spFxProvisioning }>
      //   <WebPartTitle displayMode={this.props.displayMode}
      //     title={this.props.title}
      //     updateProperty={this.props.updateProperty} />
      //     {needsConfiguration &&
      //     <Placeholder
      //       iconName='Edit'
      //       iconText='Configure your web part'
      //       description='Please configure the web part.'
      //       buttonLabel='Configure'
      //       onConfigure={onConfigure} />
      //   }
      //   {!needsConfiguration &&
      //     loading &&
      //     <div style={{ textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading items..." /></div>}
      //   {!needsConfiguration &&
      //     !loading &&
      //     error &&
      //     <div style={{ textAlign: 'center' }}>The following error has occurred while loading items: <span>{error}</span></div>}
      //   {!needsConfiguration &&
      //     !loading &&
      //     items.length === 0 &&
      //     <div style={{ textAlign: 'center' }}>No items found in the selected list</div>}
      //   {!needsConfiguration &&
      //     !loading &&
      //     items.length > 0 &&
      //     <ListView
      //       items={items}
      //       viewFields={[{
      //         displayName: 'Name',
      //         name: 'FileLeafRef',
      //         linkPropertyName: 'FileRef'
      //       }]}
      //       iconFieldName="FileRef"
      //       compact={false}
      //       selectionMode={SelectionMode.none} />
      //   }
      // </div>
      <div></div>
    );
  
  }
}
