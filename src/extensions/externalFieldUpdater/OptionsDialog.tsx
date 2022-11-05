import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
PrimaryButton,
DefaultButton,
DialogFooter,
DialogContent,
Dropdown,
DropdownMenuItemType,
IDropdownOption,
IDropdownStyles
} from 'office-ui-fabric-react';



interface OptionDialogContentProps{
  close: ()=> void;
  submit: (value: IDropdownOption) => void;
}

class OptionDialogContent extends React.Component<OptionDialogContentProps, {}> {
  private _selectedKey: IDropdownOption;
  private dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
  private dropdownOptions = [
    { key: 'Options', text: 'Value', itemType: DropdownMenuItemType.Header },
    { key: 'Yes', text: 'Yes' },
    { key: 'No', text: 'No' },
  ];
  constructor(props) {
      super(props);
      this._selectedKey = { key: 'No', text: 'No' };
  }

  public render(): JSX.Element {
      return <DialogContent
        title='External Value'
        onDismiss={this.props.close}
        showCloseButton={true}
        >
        <Dropdown
          label="Pick Updated Value:"
          // eslint-disable-next-line react/jsx-no-bind
          onChange={this._onChange}
          placeholder="Select an option"
          options={this.dropdownOptions}
          styles={this.dropdownStyles}
        />
        <DialogFooter>
            <DefaultButton text='Cancel' title='Cancel' onClick={this.props.close} />
            <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._selectedKey); }} />
        </DialogFooter>
      </DialogContent>;
  }

  private _onChange = (ev: React.SyntheticEvent<HTMLElement, Event>, selectedKey: IDropdownOption) => {
      this._selectedKey = selectedKey;
  }
}
export default class OptionDialog extends BaseDialog {
    public message: string;
    public selectedKey: IDropdownOption;

    public render(): void {
        ReactDOM.render(<OptionDialogContent
        close={ this.close }
        submit={ this._submit }
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
        isBlocking: false
        };
    }

    protected onAfterClose(): void {
        super.onAfterClose();


        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }

    private _submit = (selectedKey: IDropdownOption) => {
        this.selectedKey = selectedKey;
        this.close();
    }
  }


