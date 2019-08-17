import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import PanelMessageBar from "./components/PanelMessageBar";
import { IPanelMessageBarProps } from './components/PanelMessageBar';

import * as strings from 'SpFxPanelMessagesApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpFxPanelMessagesApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxPanelMessagesApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxPanelMessagesApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFxPanelMessagesApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;


  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve();
  }

  private async _renderPlaceHolders(): Promise<void> {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topPlaceholder.domElement) {

        const element: React.ReactElement<IPanelMessageBarProps> = React.createElement(PanelMessageBar, {});
        ReactDOM.render(element, this._topPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
