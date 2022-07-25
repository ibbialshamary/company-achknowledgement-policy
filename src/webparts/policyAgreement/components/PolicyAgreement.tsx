import * as React from "react";
import styles from "./PolicyAgreement.module.scss";
import { IPolicyAgreementProps } from "./IPolicyAgreementProps";
import PolicyDocument from "./PolicyDocument/PolicyDocument";
export default class PolicyAgreement extends React.Component<
  IPolicyAgreementProps,
  {}
> {
  public render(): React.ReactElement<IPolicyAgreementProps> {
    const { description, hasTeamsContext, userDisplayName, context } =
      this.props;

    return (
      <div>
        <h1>Hi, {userDisplayName}! Before you continue</h1>
        <p>Please read the mandatory policy below</p>
        <PolicyDocument context={this.props.context} />
      </div>
    );
  }
}
