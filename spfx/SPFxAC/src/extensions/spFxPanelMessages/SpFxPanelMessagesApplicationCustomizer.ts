import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'SpFxPanelMessagesApplicationCustomizerStrings';

import * as io from "socket.io-client";

import PanelMessageBar from "./components/PanelMessageBar";
import { IPanelMessageBarProps } from './components/PanelMessageBar';

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
  private _socket: SocketIOClient.Socket;
  private _siteUrl: string;

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
        this._siteUrl = this.context.pageContext.site.absoluteUrl;
        this._socket = io("https://spsexpress.azurewebsites.net/");

        this._socket.on('connect', () => {
          this._socket.emit('room', this._siteUrl);
        });

        const element: React.ReactElement<IPanelMessageBarProps> = React.createElement(PanelMessageBar,
          {
            siteUrl: this.context.pageContext.site.absoluteUrl,
            socket: this._socket,

          });
        ReactDOM.render(element, this._topPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}

