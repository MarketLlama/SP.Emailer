import * as React from 'react';
import styles from './SocialButtons.module.scss';
import { ISocialButtonsProps } from './ISocialButtonsProps';
import { Button } from 'office-ui-fabric-react';
import { sp, Items, ItemAddResult } from "@pnp/sp";

declare var yam;

export interface ISocialButtonsState {
  isSubscribed : boolean;
  showModal : boolean;
}

export default class SocialButtons extends React.Component<ISocialButtonsProps, ISocialButtonsState> {
  private _subScriptionListItems : Items = {} as Items;
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
  }
  public render(): React.ReactElement<ISocialButtonsProps> {
    return (
      <div className={ styles.socialButtons }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              {!this.state.isSubscribed?
              <Button className={styles.button}
                iconProps={{iconName : "Heart"}}
                text="Subscribe"
                onClick={this._subscribeToPage}
              /> :
              <Button className={styles.button}
                iconProps={{iconName : "HeartBroken"}}
                text="Unsubscribe"
                onClick={this._unsubscribeToPage}
              />}
            </div>
            <div className={ styles.column }>
              <Button className={`${styles.yammerButton} yammer-button`}
                iconProps={{iconName : "YammerLogo"}}
                text="Share"
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _subscribeToPage = () =>{
    this._subScriptionListItems.add({
      PageID : this.props.pageId,
      UserID : this.props.context.pageContext.user.loginName,
      SubscriptionEmail : this.props.context.pageContext.user.email
    }).then((value : ItemAddResult) =>{
      this.setState({
        isSubscribed : true
      });
    }, err =>{
      console.log(err);
    });
  }

  private _unsubscribeToPage =() =>{
    const PageID = this.props.pageId;
    const UserID = this.props.context.pageContext.user.loginName;
    this._subScriptionListItems.filter(`PageID eq '${PageID}' and UserID eq ${UserID}`)
      .get().then((value : any[]) =>{
        if(value){
          this._subScriptionListItems.getById(value[0].Id).delete().then(v =>{
            this.setState({
              isSubscribed : false
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
    this._subScriptionListItems.filter(`PageID eq '${PageID}' and UserID eq ${UserID}`)
      .get().then((value : any[]) =>{
        if(value){
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
    }, 5000);
  }

}
