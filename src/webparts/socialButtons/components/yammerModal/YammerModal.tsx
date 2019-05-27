import * as React from 'react';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import Box from 'yamui/dist/components/Box';
import ProgressIndicator from 'yamui/dist/components/ProgressIndicator';
import TextField from 'yamui/dist/components/TextField';
import PreviewCard from 'yamui/dist/components/PreviewCard';
import Block, { GutterSize, TextSize }  from 'yamui/dist/components/Block';
import UserPicker from 'yamui/dist/components/UserPicker';
import {FixedGridRow, FixedGridColumn} from 'yamui/dist/components/FixedGrid';
import Button, { ButtonColor } from 'yamui/dist/components/Button';

export interface YammerModalProps {
  showModal : boolean;
}

export interface YammerModalState {
  showModal : boolean;
}

class YammerModal extends React.Component<YammerModalProps, YammerModalState> {
  constructor(props: YammerModalProps) {
    super(props);
    this.state = {
     showModal : props.showModal
     };
  }
  render() {
    const people = [
      {
        key: 1,
        imageUrl: '',
        imageInitials: 'PV',
        primaryText: 'Annie Lindqvist',
        secondaryText: 'Designer',
      },
      {
        key: 2,
        imageUrl: '',
        imageInitials: 'AR',
        primaryText: 'Aaron Reid',
        secondaryText: 'Designer',
      },
      {
        key: 3,
        imageUrl: '',
        imageInitials: 'AL',
        primaryText: 'Alex Lundberg',
        secondaryText: 'Software Developer',
      },
      {
        key: 4,
        imageUrl: '',
        imageInitials: 'RK',
        primaryText: 'Roko Kolar',
        secondaryText: 'Financial Analyst',
      },
      {
        key: 5,
        imageUrl: '',
        imageInitials: 'CB',
        primaryText: 'Christian Bergqvist',
        secondaryText: 'Sr. Designer',
      },
    ];
    return ( <div>
      <Modal
        isOpen={this.state.showModal}
        onDismiss={this._closeModal}
        isBlocking={false}
      >
        <div>
          <div style={{ backgroundColor: '#f98985', transition: 'background-color 1s' }}>
            <Block padding={GutterSize.XXLARGE} textSize={TextSize.LARGE}>Share on Yammer</Block>
          </div>
          <ProgressIndicator percentComplete={0.3} ariaValueText="Thirty percent" />
          <Box>
            <TextField
              placeHolder="Textfield Placeholder..."
              errorMessage="Error hint goes here"
              description="This should not be shown"
            />
            <PreviewCard
              name="Filename.gif"
              description="this is the file description"
              imageUrl="user.png"
            />
          </Box>
          <UserPicker onResolveSuggestions={() => people} />
          <Button text="Full width" fullWidth={true} color={ButtonColor.PRIMARY} />
        </div>
      </Modal>
    </div> );
  }

  private _closeModal = () =>{
    this.setState({
      showModal : false
    });
  }

  private _getGroups = () =>{

  }
  public componentDidUpdate(prevProps: YammerModalProps, prevState: YammerModalState) {
    if(prevProps.showModal !== prevState.showModal){
      this.setState({
        showModal : !this.state.showModal
      });
    }
  }
}

export default YammerModal;
