import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsChatExportWpWebPartStrings';
import TeamsChatExportWp from './components/TeamsChatExportWp';
import { ITeamsChatExportWpProps } from './components/ITeamsChatExportWpProps';

export interface ITeamsChatExportWpWebPartProps {
  description: string;
}

export default class TeamsChatExportWpWebPart extends BaseClientSideWebPart<ITeamsChatExportWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITeamsChatExportWpProps> = React.createElement(
      TeamsChatExportWp,
      {
        description: this.properties.description,
        context: this.context
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
