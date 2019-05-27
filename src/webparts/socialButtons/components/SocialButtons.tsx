import * as React from 'react';
import styles from './SocialButtons.module.scss';
import { ISocialButtonsProps } from './ISocialButtonsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Button } from 'office-ui-fabric-react';

export interface ISocialButtonsState {
  isSubscribed : boolean;
}


export default class SocialButtons extends React.Component<ISocialButtonsProps, ISocialButtonsState> {
  constructor(props : ISocialButtonsProps){
    super(props);
    this.state ={
      isSubscribed : true //To do with prop
    };
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
              <Button className={styles.yammerButton}
                iconProps={{iconName : "YammerLogo"}}
                text="Share"
                onClick={this._sharePageOnYammer}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _subscribeToPage = () =>{
    console.log('Subscribed');
    this.setState({
      isSubscribed : true
    });
  }

  private _unsubscribeToPage =() =>{
    this.setState({
      isSubscribed : false
    });
  }

  private _sharePageOnYammer = ()=>{
    console.log('Shared');
  }

  private _isSubscribedToPage = ()=>{

  }

}
