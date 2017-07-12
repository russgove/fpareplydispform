/// see https://github.com/SharePoint/sp-dev-docs/blob/master/docs/spfx/web-parts/guidance/call-microsoft-graph-from-your-web-part.md

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as AuthenticationContext from 'adal-angular';
import * as strings from 'fpaReplyFormStrings';
import FpaReplyForm from './components/FpaReplyForm';
import { IFpaReplyFormProps } from './components/IFpaReplyFormProps';
import { IFpaReplyFormWebPartProps } from './IFpaReplyFormWebPartProps';

export default class FpaReplyFormWebPart extends BaseClientSideWebPart<IFpaReplyFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFpaReplyFormProps > = React.createElement(
      FpaReplyForm,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
