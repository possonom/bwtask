import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { Logger, ConsoleListener, LogLevel } from "@pnp/logging";

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { initializeProjectService } from '../../services/projectService';

const packageSolution: any = require("../../../config/package-solution.json");

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public async onInit(): Promise<void> {
    await super.onInit();

    // Initialize PnP logging for SharePoint 2019
    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;
    
    try {
      // Initialize project service for SharePoint 2019
      initializeProjectService(this.context);
      Logger.write('HelloWorld WebPart initialized successfully for SharePoint 2019', LogLevel.Info);
    } catch (error) {
      Logger.write('Error initializing HelloWorld WebPart: ' + (error instanceof Error ? error.message : String(error)), LogLevel.Error);
    }
  }

  public render(): void {
    Logger.write('### INIT HelloWorldWebpart, Version ' + packageSolution.solution.version, LogLevel.Info);

    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        context: this.context,
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
