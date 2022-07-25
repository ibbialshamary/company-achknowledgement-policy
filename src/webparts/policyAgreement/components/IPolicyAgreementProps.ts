import { IEmployeeListItem } from "../../../models/IEmployeeListItem";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ButtonClickedCallback } from "../../../models/ButtonClickedCallback";

export interface IPolicyAgreementProps {
  spListItems: IEmployeeListItem[];
  onGetListItems: ButtonClickedCallback;
  onAddEmployeeToList: ButtonClickedCallback;
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
