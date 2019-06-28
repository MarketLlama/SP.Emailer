import * as React from 'react';
import styles from './SocialButtons.module.scss';
import { ISocialButtonsProps } from './ISocialButtonsProps';
import { Icon } from 'office-ui-fabric-react';
import { sp, Items, ItemAddResult } from "@pnp/sp";
import ReactTooltip from 'react-tooltip';
import swal from 'sweetalert';

declare var yam;

export interface ISocialButtonsState {
  isSubscribed : boolean;
  showModal : boolean;
}

export default class SocialButtons extends React.Component<ISocialButtonsProps, ISocialButtonsState> {
  private _subScriptionListItems : Items = {} as Items;
  private _mailToText : string;
  constructor(props : ISocialButtonsProps){
    super(props);
    this.state ={
      isSubscribed : false,
      showModal : false
    };
    sp.setup({
      spfxContext : this.props.context
    });
    this._subScriptionListItems = sp.web.lists.getByTitle('Subscriptions').items;
    this._isSubscribedToPage();
    this._loadYammerIntergration();
    this._mailToText=
      `mailto:ourpointofview@syngenta.com?subject=Syngenta Positions Site&body=Site Page: ${window.location.href}`;
  }
  public render(): React.ReactElement<ISocialButtonsProps> {
    return (
      <div className={ styles.socialButtons }>
        <div className={ styles.container }>
          <ul className={styles.socialButtonList} >
            <li>
            {this.state.isSubscribed?
              <div>
                <a data-tip data-for='unsubscribe' onClick={this._unsubscribeToPage} className={styles.subscribeButton}>
                  <Icon iconName="HeartBroken" />
                </a>
                <ReactTooltip id='unsubscribe' type='dark' effect='solid'>
                  <span>Unsubscribe to article</span>
                </ReactTooltip>
              </div> :
              <div>
                <a data-tip data-for='subscribe' onClick={this._subscribeToPage} className={styles.subscribeButton}>
                  <Icon iconName="Heart" />
                </a>
                <ReactTooltip id='subscribe' type='dark' effect='solid'>
                  <span>Subscribe to article</span>
                </ReactTooltip>
              </div>
            }
            </li>
            <li>
              <a data-tip data-for='yammer' href="#" className={`${styles.yammerButton} yammer-button`}>
                <Icon iconName="YammerLogo" />
              </a>
              <ReactTooltip id='yammer' type='dark' effect='solid'>
                <span>Share on Yammer</span>
              </ReactTooltip>
            </li>
            <li>
              <a data-tip data-for='mail' href={this._mailToText} className={styles.basicButton} target="_top">
                <Icon iconName="Mail" />
              </a>
              <ReactTooltip id='mail' type='dark' effect='solid'>
                <span>Mail feedback to site owner</span>
              </ReactTooltip>
            </li>
          </ul>
        </div>
      </div>
    );
  }

  private _subscribeToPage = () =>{
    this._subScriptionListItems.add({
      SubscriptionPageID : this.props.pageId,
      SubscriptionUserID : this.props.context.pageContext.user.loginName,
      SubscriptionEmail : this.props.context.pageContext.user.email
    }).then((value : ItemAddResult) =>{
      this.setState({
        isSubscribed : true
      });
      swal({
        title: "Subscribed!",
        text: `You are subscribed to this page to recieve email updates.`,
        icon: "success",
        buttons : {
            confirm: {
              text: "OK",
              value: true,
              visible: true,
              className: "",
              closeModal: true
            }
        }
      });
    }, err =>{
      console.log(err);
    });
  }

  private _unsubscribeToPage =() =>{
    const PageID = this.props.pageId;
    const UserID = this.props.context.pageContext.user.loginName;
    this._subScriptionListItems.filter(`SubscriptionPageID eq '${PageID}' and SubscriptionUserID eq '${UserID}'`)
      .get().then((value : any[]) =>{
        if(value.length > 0 ){
          this._subScriptionListItems.getById(value[0].Id).delete().then(v =>{
            this.setState({
              isSubscribed : false
            });
            swal({
              title: "Unsubscribed!",
              text: `You have unsubscribed from this article.`,
              icon: "info",
              buttons : {
                  confirm: {
                    text: "OK",
                    value: true,
                    visible: true,
                    className: "",
                    closeModal: true
                  }
              }
            });
          }, err => console.log(err));
        }
      }, err =>{
        console.log(err);
      });

  }

  private _isSubscribedToPage = ()=>{
    const PageID = this.props.pageId;
    const UserID = this.props.context.pageContext.user.loginName;
    this._subScriptionListItems.filter(`SubscriptionPageID eq '${PageID}' and SubscriptionUserID eq '${UserID}'`)
      .get().then((value : any[]) =>{
        if(value.length > 0){
          this.setState({
            isSubscribed: true
          });
        }
      });
  }

  private _loadYammerIntergration = () =>{
    let options = {
      customButton : true, //false by default. Pass true if you are providing your own button to trigger the share popup
      classSelector: 'yammer-button',//if customButton is true, you must pass the css class name of your button (so we can bind the click event for you)
      defaultMessage: 'Check this out.', //optionally pass a message to prepopulate your post
    };
    //Have to wait for the external yammer file to load.
    setTimeout(() => {
      yam.platform.yammerShare(options);
    }, 3000);
  }

}
