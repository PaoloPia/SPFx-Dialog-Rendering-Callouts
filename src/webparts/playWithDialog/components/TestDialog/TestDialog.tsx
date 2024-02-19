import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
    Dialog
} from '@fluentui/react';

import TestDialogContent from './TestDialogContent';
import { ITestDialogContentProps } from './ITestDialogContentProps';

const TestDialog: React.FC<ITestDialogContentProps> = (props: ITestDialogContentProps) => {
    return (
        <Dialog
            hidden={false}
            onDismiss={props.cancel}
            dialogContentProps={{
                type: 1,
                title: 'Test Dialog'
            }}
            modalProps={{
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
            }}
        >
            <TestDialogContent
                context={props.context}
                cancel={props.cancel}
                submit={props.submit}
            />
        </Dialog>
    );
};  

export default class TestDialogManager {

    public context: WebPartContext;
    public cancel: () => Promise<void>;
    public submit: (recipients: string[], selectedOption: string) => Promise<void>;

    // this is dialog container
    private domElement: HTMLDivElement | null = null;

    public async close(): Promise<void> {
        if (this.domElement) {
            ReactDOM.unmountComponentAtNode(this.domElement);
            this.domElement.remove();
            this.domElement = null;
        }
    }

    public async show(): Promise<void> {
        this.domElement = document.createElement('div');
        document.body.appendChild(this.domElement);

        ReactDOM.render(
            <TestDialog
                context={this.context}
                cancel={this.closeWithCancel}
                submit={this.closeWithSubmit}
            />, this.domElement);
    }

    private closeWithCancel = async (): Promise<void> => {
        await this.cancel(); 
        await this.close();
    }

    private closeWithSubmit =  async (recipients: string[], selectedOption: string): Promise<void> => {
        await this.submit(recipients, selectedOption); 
        await this.close();
    }
}