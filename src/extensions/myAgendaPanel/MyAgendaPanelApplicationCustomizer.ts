import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "MyAgendaPanelApplicationCustomizerStrings";
import * as React from "react";
import * as ReactDom from "react-dom";
import { IGraphToolkitProps } from "./components/IGraphToolkitProps";
import GraphToolkit from "./components/GraphToolkit";
import { Providers, SharePointProvider } from '@microsoft/mgt/dist/commonjs';
const LOG_SOURCE: string = "MyAgendaPanelApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMyAgendaPanelApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class MyAgendaPanelApplicationCustomizer extends BaseApplicationCustomizer<IMyAgendaPanelApplicationCustomizerProperties> {
  private static _topPlaceholder?: PlaceholderContent;
  private _handleDispose(): void {}
  @override
  protected async  onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context);
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // add handler to he page navigated event which occurs both when
    // the user opens and leaves the page
    this.context.application.navigatedEvent.add(this, this._render);
    let message: string = this.properties.testMessage;
    if (!message) {
      message = "(No properties were provided.)";
    }
  }
  
  private _render(): void {
    // check if the application customizer has already been rendered
    if (!MyAgendaPanelApplicationCustomizer._topPlaceholder) {
      // create a DOM element in the top placeholder for the application customizer
      // to render
      MyAgendaPanelApplicationCustomizer._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: this._handleDispose
        }
      );
    }

    // if the top placeholder is not available, there is no place in the UI
    // for the app customizer to render, so quit.
    if (!MyAgendaPanelApplicationCustomizer._topPlaceholder) {
      return;
    }

    const element: React.ReactElement<IGraphToolkitProps> = React.createElement(
      GraphToolkit,
      {
        description: "Hello world"
      }
    );

    // render the UI using a React component
    ReactDom.render(
      element,
      MyAgendaPanelApplicationCustomizer._topPlaceholder.domElement
    );
  }
}
