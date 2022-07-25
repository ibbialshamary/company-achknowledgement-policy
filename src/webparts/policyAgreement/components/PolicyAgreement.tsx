import * as React from "react";
import styles from "./PolicyAgreement.module.scss";
import { IPolicyAgreementProps } from "./IPolicyAgreementProps";
import PolicyDocument from "./PolicyDocument/PolicyDocument";
import { DefaultButton } from "office-ui-fabric-react";
import { IPolicyAgreementState } from "./IPolicyAgreementState";
import { IEmployeeListItem } from "../../../models/IEmployeeListItem";
export default class PolicyAgreement extends React.Component<
  IPolicyAgreementProps,
  IPolicyAgreementState
> {
  public constructor(props: IPolicyAgreementProps) {
    super(props);

    this.state = {
      userHasAgreed: false,
    };

    this.updateUserAgreementStatus = this.updateUserAgreementStatus.bind(this);
    this.readPolicyAgain = this.readPolicyAgain.bind(this);
  }

  public componentDidMount(): void {
    this.props.onGetListItems();
  }

  public componentDidUpdate(
    prevProps: Readonly<IPolicyAgreementProps>,
    prevState: Readonly<IPolicyAgreementState>
  ): void {
    const employeeAgreed: boolean =
      this.props?.spListItems[0]?.AcknowledgedCompanyPolicy;
    if (!employeeAgreed) return;

    if (prevState === this.state) {
      this.setState({
        userHasAgreed: this.props.spListItems[0].AcknowledgedCompanyPolicy,
      });
    }
  }

  public updateUserAgreementStatus(): void {
    this.setState({
      userHasAgreed: true,
    });

    // get the current user's list item and if agreed to policy, return, else set the state and add list item
    const employees: IEmployeeListItem[] = this.props?.spListItems;
    if (
      employees &&
      employees.length > 0 &&
      employees[0].AcknowledgedCompanyPolicy
    ) {
      console.log("Returning as employee record already exists on list");
      return;
    }

    this.props.onAddEmployeeToList();

    console.log("Added employee details and agreement to list");
  }

  public readPolicyAgain(): void {
    this.setState({
      userHasAgreed: false,
    });
  }

  public render(): React.ReactElement<IPolicyAgreementProps> {
    const { userDisplayName } = this.props;

    return (
      <div className={styles.policyAgreement}>
        {this.state.userHasAgreed ? (
          <>
            <h1>Welcome back</h1>
            <p>
              You have already agreed to the policy agreement and do not have to
              again.
            </p>
            <DefaultButton onClick={this.readPolicyAgain}>
              Read policy again
            </DefaultButton>
          </>
        ) : (
          <>
            <h1>Hi, {userDisplayName}! Before you continue</h1>
            <p>Please read the mandatory policy below</p>
            <PolicyDocument context={this.props.context} />
            <div className={styles.agreeButton}>
              <DefaultButton
                text="Agree and Continue"
                onClick={this.updateUserAgreementStatus}
              />
            </div>
          </>
        )}
      </div>
    );
  }
}
