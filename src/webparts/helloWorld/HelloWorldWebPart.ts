import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import {  Logger, ConsoleListener, LogLevel } from "@pnp/logging";

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

import { PackageSolution } from '../../interfaces/PackageSolution';
import { ITheme } from '@fluentui/react/lib/Styling';
import { loadCustomTheme } from '../../theme/customTheme';

const packageSolution: PackageSolution = require("../../../config/package-solution.json");

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private theme: ITheme;

  // we need to make sure to init pnp and its logger here
  // for PROD we should set the LogLevel to Error
  public async onInit(): Promise<void> {
    this.theme = loadCustomTheme();
    await super.onInit();

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Info;   
  }

  public render(): void {
    
    Logger.write(`### INIT HelloWorldWebpart, Version ${packageSolution.solution.version}`, LogLevel.Info); // to help testers identify the version

    const element: React.ReactElement<IHelloWorldProps > = React.createElement(
      HelloWorld,
      {
        theme: this.theme,
        description: this.properties.description, 
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
