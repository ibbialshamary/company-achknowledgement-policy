import { IEmployeeListItem } from "../../models/IEmployeeListItem";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "PolicyAgreementWebPartStrings";
import PolicyAgreement from "./components/PolicyAgreement";
import { IPolicyAgreementProps } from "./components/IPolicyAgreementProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IEmployeeList } from "../../models/IEmployeeList";

export interface IPolicyAgreementWebPartProps {
  description: string;
}

export default class PolicyAgreementWebPart extends BaseClientSideWebPart<IPolicyAgreementWebPartProps> {
  private _employees: IEmployeeListItem[] = [];

  public render(): void {
    const element: React.ReactElement<IPolicyAgreementProps> =
      React.createElement(PolicyAgreement, {
        spListItems: this._employees,
        onGetListItems: this._onGetListItems,
        onAddEmployeeToList: this._onAddListItem,
        description: this.properties.description,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
      });

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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private _getCurrentEmployeeByEmail(): Promise<IEmployeeListItem[]> {
    const email: string = this.context.pageContext.user.email;
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Employees')/items?$filter=Email eq '${email}'`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<IEmployeeListItem[]>;
  }

  private _onGetListItems = (): void => {
    this._getCurrentEmployeeByEmail()
      .then((response) => {
        this._employees = response;
        this.render();
      })
      .catch((err) => console.error(err));
  };

  // code for adding employee data to list
  private _getItemEntityType(): Promise<string> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Employees')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      )
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.ListItemEntityTypeFullName;
      }) as Promise<string>;
  }

  private _addListItem(
    email: string,
    firstName: string,
    lastName: string,
    agreed: boolean
  ): Promise<SPHttpClientResponse> {
    return this._getItemEntityType().then((spEntityType) => {
      const request: any = {};
      request.body = JSON.stringify({
        Email: email,
        FirstName: firstName,
        LastName: lastName,
        AcknowledgedCompanyPolicy: agreed,
        "@odata.type": spEntityType,
      });

      return this.context.spHttpClient.post(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Employees')/items`,
        SPHttpClient.configurations.v1,
        request
      );
    });
  }

  private _onAddListItem = (): void => {
    const displayName: string = this.context.pageContext.user.displayName;
    const email: string = this.context.pageContext.user.email;
    const firstName: string = displayName.split(" ")[0];
    const lastName: string = displayName.split(" ")[1];

    this._addListItem(email, firstName, lastName, true)
      .then(() => {
        this._getCurrentEmployeeByEmail()
          .then((response) => {
            this._employees = response;
            this.render();
          })
          .catch((err) => console.error(err));
      })
      .catch((err) => console.error(err));
  };
}
