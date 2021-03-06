import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'SocialButtonsWebPartStrings';
import SocialButtons from './components/SocialButtons';
import { ISocialButtonsProps } from './components/ISocialButtonsProps';

export interface ISocialButtonsWebPartProps {
  description: string;
}

export default class SocialButtonsWebPart extends BaseClientSideWebPart<ISocialButtonsWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ISocialButtonsProps > = React.createElement(
      SocialButtons,
      {
        context: this.context,
        pageId : this.context.pageContext.listItem.id
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    let gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://s0.assets-yammer.com/assets/platform_social_buttons.min.js`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);
    return Promise.resolve();
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
