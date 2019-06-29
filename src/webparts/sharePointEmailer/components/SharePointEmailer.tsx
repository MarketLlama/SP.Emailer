import * as React from 'react';
import styles from './SharePointEmailer.module.scss';
import { ISharePointEmailerProps } from './ISharePointEmailerProps';
import { ISharePointEmailerState } from './ISharePointEmailerState';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { sp, EmailProperties } from "@pnp/sp";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SecurityTrimmedControl, PermissionLevel } from "@pnp/spfx-controls-react/lib/SecurityTrimmedControl";
import { SPPermission } from '@microsoft/sp-page-context';
import swal from 'sweetalert';
import '../emailTemplates/standardEmailTemplate.html';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

export interface SubscribedUser {
  Id?: number;
  SubscriptionUserID: string;
  SubscriptionEmail: string;
  SubscriptionPageID?: string;
  Author: Author;
}

export interface Author {
  Title: string;
  FirstName?: string;
}

export default class SharePointEmailer extends React.Component<ISharePointEmailerProps, ISharePointEmailerState> {
  private _emailTemplate = require("../emailTemplates/standardEmailTemplate.html");
  private _users: SubscribedUser[] = [];
  private _defaultUsers = [];
  private _currentPage = {
    Title: '',
    FileRef: ''
  };

  constructor(props) {
    super(props);
    console.log(window.innerWidth);
    this.state = {
      showModal: false,
      emailText: '',
      isLoading: false,
      isMobileViewPort: (window.innerWidth <= 760)
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
              {this.state.isLoading ? <ProgressIndicator /> : null}
              <PeoplePicker
                context={this.props.context}
                titleText="Subscribers"
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                personSelectionLimit={100}
                isRequired={true}
                defaultSelectedUsers={this._defaultUsers}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
              />
              <p>Enter content of email.</p>
              <div>
                {this.state.isMobileViewPort ?
                  <TextField multiline rows={11}
                    onChanged={this._setMailText}
                    value={this.state.emailText} /> :
                  <RichText value={this.state.emailText}
                    onChange={(text) => this._setMailText(text)}
                    className={styles.textBox}
                  />}
              </div>
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
    this._setLoading(false);
    this.setState({ showModal: false, emailText : "" });
  }

  private _setLoading = (isLoading: boolean) => {
    this.setState({ isLoading });
  }

  private _setMailText = (newText: string): string => {
    this.setState({
      emailText: newText,
    });
    return newText;
  }

  private _getPeoplePickerItems = (items: any[]) => {
    this._users = [];
    items.forEach(item => {
      this._users.push({
        SubscriptionEmail: item.secondaryText,
        SubscriptionUserID: item.id,
        Author: {
          Title: item.text
        }
      });
    });
    console.log(this._users);
  }

  private _getSubscriptions() {
    return new Promise<SubscribedUser[]>((resolve, reject) => {
      try {
        const PageID = this.props.context.pageContext.listItem.id;
        sp.web.lists.getByTitle("Subscriptions").items
          .expand('Author').select('Id', 'SubscriptionUserID', 'SubscriptionEmail', 'SubscriptionPageID', 'Author/Title', 'Author/FirstName')
          .filter(`SubscriptionPageID eq '${PageID}'`).get().then((items: SubscribedUser[]) => {
            resolve(items);
          });

      } catch (e) {
        console.log(e);
        reject();
      }
    });
  }

  private _getPageDetails = async () => {
    const id = this.props.context.pageContext.listItem.id;
    const pages = sp.web.lists.getByTitle('Site Pages').items;
    let page = await pages.getById(id).select("Title", "FileRef").get();
    console.log(page);
    return page;
  }

  private _getEmailContent = (user: SubscribedUser): string => {
    let emailTemplate = this._emailTemplate.toString();
    emailTemplate = emailTemplate.replace(/{{emailContent}}/gi, this.state.emailText)
      .replace(/{{pageURL}}/gi, window.location.href)
      .replace(/{{userName}}/gi, user.Author.Title)
      .replace(/{{pageTitle}}/gi, this._currentPage.Title);

    return emailTemplate;
  }

  private _sendEmail = (): void => {
    let emailPromises = [];
    this._setLoading(true);
    this._users.forEach(async user => {
      let emailContent = this._getEmailContent(user);

      const emailProps: EmailProperties = {
        To: [user.SubscriptionEmail],
        Subject: `${this.props.context.pageContext.web.title} : ${this._currentPage.Title}`,
        Body: emailContent
      };

      emailPromises.push(sp.utility.sendEmail(emailProps));
    });
    Promise.all(emailPromises).then(response => {
      this._closeModal();
      swal({
        title: "Email(s) have been sent!",
        text: `Emails have been successfully delivered to subscribers.`,
        icon: "success",
        buttons: {
          confirm: {
            text: "OK",
            value: true,
            visible: true,
            className: "",
            closeModal: true
          }
        }
      });
    }, error => {
      this._closeModal();
      swal("Failure to send email", "Email has not sent please contact SharePoint support.", "error");
    }).catch(error => {
      this._closeModal();
      swal("Failure to send email", "Email has not sent please contact SharePoint support.", "error");
    });
  }

  public componentDidMount() {
    this._main();
  }

  private _main = async () => {
    try {
      //Get details of current page.
      this._currentPage = await this._getPageDetails();
      //Gets the subscribers of the page and sets them as default email contacts for the emailer.
      //Uses the defaultSelectedUsers
      let subscriptions: SubscribedUser[] = await this._getSubscriptions();
      this._users = subscriptions;
      subscriptions.forEach(item => {
        this._defaultUsers.push(item.SubscriptionEmail);
      });
    } catch (error) {
      console.log(error);
    }
  }
}
