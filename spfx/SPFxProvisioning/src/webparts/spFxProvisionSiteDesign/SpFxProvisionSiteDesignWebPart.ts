import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxProvisionSiteDesignWebPartStrings';
import SpFxProvisionSiteDesign from './components/SpFxProvisionSiteDesign';
import { ISpFxProvisionSiteDesignProps } from './components/ISpFxProvisionSiteDesignProps';


export interface ISpFxProvisionSiteDesignWebPartProps {
  description: string;
}

export default class SpFxProvisionSiteDesignWebPart extends BaseClientSideWebPart<ISpFxProvisionSiteDesignWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpFxProvisionSiteDesignProps > = React.createElement(
      SpFxProvisionSiteDesign,
      {
        client: this.context.spHttpClient,
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
