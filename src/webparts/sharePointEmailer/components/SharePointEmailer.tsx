import * as React from 'react';
import styles from './SharePointEmailer.module.scss';
import { ISharePointEmailerProps } from './ISharePointEmailerProps';
import { ISharePointEmailerState } from './ISharePointEmailerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton ,ActionButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp, EmailProperties } from "@pnp/sp";
import '../emailTemplates/standardEmailTemplate.html';

export default class SharePointEmailer extends React.Component<ISharePointEmailerProps, ISharePointEmailerState> {
  private _emailTemplate = require("../emailTemplates/standardEmailTemplate.html");


  constructor(props) {
    super(props);
    this.state = {
      showModal: false,
      emailText : ''
    };

    this._setMailText = this._setMailText.bind(this);
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
              <span id="titleId">Send Email</span>
              <ActionButton className={styles.closeButton} iconProps={{ iconName: 'Cancel' }} onClick={this._closeModal}/>
            </div>
            <div id="subtitleId" className={styles.modalBody}>
              <p>
                Enter content of email.{' '}
              </p>
              <TextField multiline rows={5} onChanged={this._setMailText} value={this.state.emailText}/>
              <br/>
              <PrimaryButton
                  iconProps={{ iconName: 'Mail' }}
                  text="Send Mail"
                  onClick={this._sendEmail}
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

  private _sendEmail = () : void =>{

      const emailProps: EmailProperties = {
          To: [""],
          Subject: "This email is about...",
          Body: `${this._emailTemplate}`,
      };

      sp.utility.sendEmail(emailProps).then(_ => {
          console.log("Email Sent!");
      });
  }
}
