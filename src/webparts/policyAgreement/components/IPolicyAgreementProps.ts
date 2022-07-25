import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPolicyAgreementProps {
  description: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext
}
