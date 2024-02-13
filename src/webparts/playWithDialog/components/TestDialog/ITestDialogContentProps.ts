import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITestDialogContentProps {
    context: WebPartContext;
    cancel: () => Promise<void>;
    submit: (recipients: string[], selectedOption: string) => Promise<void>;
}