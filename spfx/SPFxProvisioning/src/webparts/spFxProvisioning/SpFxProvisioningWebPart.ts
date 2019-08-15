import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxProvisioningWebPartStrings';
import SpFxProvisioning from './components/SpFxProvisioning';
import { ISpFxProvisioningProps } from './components/ISpFxProvisioningProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { sp } from '@pnp/sp';

export interface ISpFxProvisioningWebPartProps {
  listId?: string;
  siteUrl?: string;
  title: string;
}

export default class SpFxProvisioningWebPart extends BaseClientSideWebPart<ISpFxProvisioningWebPartProps> {

  public onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ISpFxProvisioningProps> = React.createElement(
      SpFxProvisioning,
      {
        displayMode: this.displayMode,
        listId: this.properties.listId,
        onConfigure: this._onConfigure,
        siteUrl: this.properties.siteUrl,
        title: this.properties.title,
        updateProperty: value => this.properties.title = value,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldTextWithCallout('siteUrl', {
                  calloutTrigger: CalloutTriggers.Click,
                  key: 'siteUrlFieldId',
                  label: 'Site URL',
                  calloutContent: React.createElement('span', {}, 'URL of the site where the document library to show documents from is located. Leave empty to connect to a document library from the current site'),
                  calloutWidth: 250,
                  value: this.properties.siteUrl
                }),
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.siteUrl,
                  baseTemplate: 101
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private _onConfigure = () => this.context.propertyPane.open();
}
