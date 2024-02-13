import * as React from 'react';
import styles from './PlayWithDialog.module.scss';
import type { IPlayWithDialogProps } from './IPlayWithDialogProps';
import { DefaultButton } from '@fluentui/react';import TestDialog from './TestDialog/TestDialog';

export default class PlayWithDialog extends React.Component<IPlayWithDialogProps, {}> {
  public render(): React.ReactElement<IPlayWithDialogProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.playWithDialog} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <DefaultButton text="Open Dialog" onClick={this.openDialog} />
        </div>
      </section>
    );
  }

  private openDialog = async (): Promise<void> => {
    const testDialog: TestDialog = new TestDialog();
    testDialog.context = this.props.context;
    testDialog.cancel = async (): Promise<void> => { 
      console.log('Cancelled');
    };
    testDialog.submit = async (recipients: string[], selectedOption: string): Promise<void> => { 
      console.log(recipients);
      console.log(selectedOption);
    };
    await testDialog.show();
  }
}