import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PolicyAgreementWebPartStrings';
import PolicyAgreement from './components/PolicyAgreement';
import { IPolicyAgreementProps } from './components/IPolicyAgreementProps';

export interface IPolicyAgreementWebPartProps {
  description: string;
}

export default class PolicyAgreementWebPart extends BaseClientSideWebPart<IPolicyAgreementWebPartProps> {


  public render(): void {
    const element: React.ReactElement<IPolicyAgreementProps> = React.createElement(
      PolicyAgreement,
      {
        description: this.properties.description,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
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
