import * as React from 'react';
import styles from './TestDialog.module.scss';

import {
    PrimaryButton,
    DefaultButton,
    DialogFooter,
    DialogContent,
    ComboBox,
    IComboBoxOption,
    IPersonaProps,
    IComboBox
} from '@fluentui/react';

import { ITestDialogContentProps } from './ITestDialogContentProps';
import { ITestDialogContentState } from './ITestDialogContentState';

// Import PnP React controls
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export default class TestDialogContent extends
  React.Component<ITestDialogContentProps, ITestDialogContentState> {

    public constructor(props: ITestDialogContentProps) {
        super(props);

        this.state = {
            recipients: [],
            selectedOption: ""
        };
    }

    public render(): JSX.Element {

        const {
            context,
            cancel,
            submit
        } = this.props;

        const { 
            recipients,
            selectedOption
        } = this.state;

        const options: IComboBoxOption[] = [
            { key: '1', text: 'One' },
            { key: '2', text: 'Two' },
            { key: '3', text: 'Three' },
            { key: '4', text: 'Four' },
            { key: '5', text: 'Five' },
            { key: '6', text: 'Six' },
            { key: '7', text: 'Seven' },
            { key: '8', text: 'Eight' },
            { key: '9', text: 'Nine' },
            { key: '10', text: 'Ten' }
        ];

        return (<div className={styles.testDialogRoot}>
            <DialogContent
                title="Test dialog"
                subText="This is a test dialog"
                onDismiss={cancel}>

                <div className={styles.testDialogContent}>
                    <div>
                        <div>
                            <PeoplePicker
                                context={context}
                                titleText="People Picker"
                                personSelectionLimit={10}
                                showtooltip={true}
                                required={true}
                                disabled={false}
                                searchTextLimit={5}
                                onChange={this._getPeoplePickerItems}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={500} />
                        </div>
                        <div>
                            <ComboBox options={options} 
                                onChange={this._optionSelected}
                                allowFreeform={true}
                                autoComplete={'on'} />
                        </div>
                    </div>
                </div>

                <DialogFooter>
                    <DefaultButton text="Cancel"
                            title="Cancel" onClick={cancel} />
                    <PrimaryButton text="Save"
                        title="Save" onClick={() => submit(recipients, selectedOption)} />
                </DialogFooter>
            </DialogContent>
        </div>);
    }

    private _getPeoplePickerItems = (items: IPersonaProps[]): void => {
        const recipients: string[] = items.map((i) => i.secondaryText?.toLowerCase() || "");
        this.setState({
            recipients: recipients
        });
    }

    private _optionSelected = (event: React.FormEvent<IComboBox>, item: IComboBoxOption): void => {
        const newValue = item.key.toString();
        this.setState({
            selectedOption: newValue
        });
    }
}
