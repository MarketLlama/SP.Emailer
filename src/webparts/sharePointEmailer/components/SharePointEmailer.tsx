import * as React from 'react';
import styles from './SharePointEmailer.module.scss';
import { ISharePointEmailerProps } from './ISharePointEmailerProps';
import { ISharePointEmailerState } from './ISharePointEmailerState';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp, EmailProperties } from "@pnp/sp";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SecurityTrimmedControl, PermissionLevel } from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import '../emailTemplates/standardEmailTemplate.html';

export default class SharePointEmailer extends React.Component<ISharePointEmailerProps, ISharePointEmailerState> {
  private _emailTemplate = require("../emailTemplates/standardEmailTemplate.html");
  private _users = [];
  private _toUsers = [];
  private _defaultUsers = [];
  private _currentPage = {
    Title: '',
    FileRef: ''
  };

  constructor(props) {
    super(props);
    this.state = {
      showModal: false,
      emailText: ''
    };
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public render(): React.ReactElement<ISharePointEmailerProps> {
    return (
      <div className={styles.sharePointEmailer}>
        <div className={styles.container}>
          <SecurityTrimmedControl context={this.props.context}
                        level={PermissionLevel.currentWeb}
                        permissions={[SPPermission.manageWeb]}>
            <ActionButton iconProps={{ iconName: 'Mail' }} onClick={this._showModal} text="Send Email to Subscribers" />
          </SecurityTrimmedControl>
          <Modal
            isOpen={this.state.showModal}
            onDismiss={this._closeModal}
            isBlocking={false}
            className={styles.modalContainer}
          >
            <div className={styles.modalHeader}>
              <span style={{ padding: "20px" }} >Send Email</span>
              <ActionButton className={styles.closeButton} iconProps={{ iconName: 'Cancel' }} onClick={this._closeModal} />
            </div>
            <div id="subtitleId" className={styles.modalBody}>
              <PeoplePicker
                context={this.props.context}
                titleText="Subcribers"
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                isRequired={true}
                defaultSelectedUsers={this._defaultUsers}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principleTypes={[PrincipalType.User]} />
              <p>
                Enter content of email.{' '}
              </p>
              <TextField multiline rows={11} onChanged={this._setMailText} value={this.state.emailText} />
              <br />
              <PrimaryButton
                iconProps={{ iconName: 'Mail' }}
                text="Send Mail"
                onClick={this._sendEmail}
                style={{ float: "right" }}
              />
            </div>
          </Modal>
        </div>
      </div>
    );
  }

  private _showModal = (): void => {
    this.setState({ showModal: true });
  }

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  }

  private _setMailText = (newText: string): void => {
    this.setState({
      emailText: newText,
    });
  }

  private _getPeoplePickerItems = (items: any[]) =>{
    items.forEach(item => {
      this._toUsers.push(item.secondaryText);
    });
  }

  private _getSubscriptions() {
    return new Promise<any[]>((resolve, reject) => {
      try {
        const PageID = this.props.context.pageContext.listItem.id;
        sp.web.lists.getByTitle("Subscriptions").items.filter("SubscriptionPageID eq '" + PageID + "'").get().then((items: any[]) => {
          resolve(items);
        });

      } catch (e) {
        console.log(e);
        reject();
      }

    });
  }

  private _getPageDetails = async () => {
    var id = this.props.context.pageContext.listItem.id;
    const pages = sp.web.lists.getByTitle('Site Pages').items;
    let page = await pages.getById(id).select("Title", "FileRef").get();
    return page;
  }

  private _getEmailContent = (user: any): string => {
    let emailTemplate = this._emailTemplate.toString();
    emailTemplate = emailTemplate.replace(/{{emailContent}}/gi, this.state.emailText)
      .replace(/{{pageURL}}/gi, window.location.href)
      .replace(/{{userName}}/gi, user.UserName)
      .replace(/{{pageTitle}}/gi, this._currentPage.Title);

    return emailTemplate;
  }

  private _sendEmail = (): void => {

    this._users.forEach(async user => {
      let emailContent = this._getEmailContent(user);

      const emailProps: EmailProperties = {
        To: [user.SubscriptionEmail],
        Subject: "Syngenta Positions : " + this._currentPage.Title,
        Body: emailContent
      };

      await sp.utility.sendEmail(emailProps);
    });

  }

  public componentDidMount() {
    this._main();
  }

  private _main = async() => {
    try {
      //Get details of current page.
      this._currentPage = await this._getPageDetails();
      //Gets the subsribers of the page and sets them as default email contacts for the emailer.
      //Uses the defaultSelectedUsers
      let subscriptions = await this._getSubscriptions();
      this._users = subscriptions;
      subscriptions.forEach(item => {
        this._defaultUsers.push(item.SubscriptionEmail);
        this._toUsers.push(item.SubscriptionEmail);
      });
    } catch (error) {
      console.log(error);
    }
  }
}
