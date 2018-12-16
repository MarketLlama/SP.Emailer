import * as React from 'react';
import styles from './SharePointEmailer.module.scss';
import { ISharePointEmailerProps } from './ISharePointEmailerProps';
import { ISharePointEmailerState } from './ISharePointEmailerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton ,ActionButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp, EmailProperties, Web } from "@pnp/sp";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import '../emailTemplates/standardEmailTemplate.html';

export default class SharePointEmailer extends React.Component<ISharePointEmailerProps, ISharePointEmailerState> {
  private _emailTemplate = require("../emailTemplates/standardEmailTemplate.html");
  private _toUsers = [];
  private _defaultUsers = [];

  constructor(props) {
    super(props);
    this.state = {
      showModal: false,
      emailText : ''
    };

    this._main = this._main.bind(this);
    this._setMailText = this._setMailText.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
  }

  public render(): React.ReactElement<ISharePointEmailerProps> {
    return (
      <div className={ styles.sharePointEmailer }>
        <div className={ styles.container }>
          <ActionButton iconProps={{ iconName: 'Mail' }} secondaryText="Opens the Sample Modal" onClick={this._showModal} text="Send Email" />
          <Modal
            titleAriaId="titleId"
            subtitleAriaId="subtitleId"
            isOpen={this.state.showModal}
            onDismiss={this._closeModal}
            isBlocking={false}
            className={ styles.modalContainer }
          >
            <div className={styles.modalHeader}>
              <span style={{padding : "20px"}} id="titleId">Send Email</span>
              <ActionButton className={styles.closeButton} iconProps={{ iconName: 'Cancel' }} onClick={this._closeModal}/>
            </div>
            <div id="subtitleId" className={styles.modalBody}>
              <PeoplePicker
                context={this.props.context}
                titleText="Subcribers"
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                isRequired={true}
                disabled={false}
                defaultSelectedUsers={this._defaultUsers}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principleTypes={[PrincipalType.User]} />
              <p>
                Enter content of email.{' '}
              </p>
              <TextField multiline rows={11} onChanged={this._setMailText} value={this.state.emailText}/>
              <br/>
              <PrimaryButton
                  iconProps={{ iconName: 'Mail' }}
                  text="Send Mail"
                  onClick={this._sendEmail}
                  style={{float:"right"}}
                />
            </div>
          </Modal>
        </div>
      </div>
    );
  }

  private _showModal = (): void => {
    this.setState({ showModal: true });
    console.log('Open Modal');
  }

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  }

  private _setMailText = (newText: string): void => {
    this.setState({
      emailText : newText,
    });
  }

  private _getPeoplePickerItems(items: any[]) {
    items.forEach(item =>{
      this._toUsers.push(item.secondaryText);
    });
  }

  private _getSubscriptions() {
    return new Promise((resolve, reject) => {
      const web = new Web(this.props.context.pageContext.site.absoluteUrl);

      try {
        const PageID = 24;//this.props.context.pageContext.listItem.id;
        web.lists.getByTitle("Subscriptions").items.filter("PageID eq '" + PageID + "'").get().then((items: any[]) => {
          resolve(items);
        });
      } catch (e) {
        console.log(e);
        reject();
      }

    });
  }

  private _getEmailContent = () : string =>{
    let emailTemplate = this._emailTemplate.toString();
    emailTemplate.replace('{{pageContent}}', this.state.emailText);
    emailTemplate.replace('{{pageURL}}',this.props.context.pageContext.web.absoluteUrl);
    emailTemplate.replace('{{userName}}','Some Name');
    emailTemplate.replace('{{pageTitle}}', 'Some Title');
    console.log(emailTemplate);
    //this._emailTemplate = emailTemplate;

    return emailTemplate;
  }

  private _sendEmail = () : void =>{

      let emailContent = this._getEmailContent();

      const emailProps: EmailProperties = {
          To: this._toUsers,
          Subject: "This email is about...",
          Body: emailContent,
      };

      sp.utility.sendEmail(emailProps).then(_ => {
          console.log("Email Sent!");
      });
  }

  public componentDidMount (){
    this._main();
  }

  private async _main(){
    //Gets the subsribers of the page and sets them as default email contacts for the emailer.
    //Uses the defaultSelectedUsers 
    let subscriptions = await this._getSubscriptions();
    console.log(subscriptions);
    let string = JSON.stringify(subscriptions);
    let arr : any[] = JSON.parse(string);
    arr.forEach(item=>{
      this._defaultUsers.push(item.Title);
    });
  }
}
