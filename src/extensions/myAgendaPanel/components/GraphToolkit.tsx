import * as React from 'react';
import styles from './GraphToolkit.module.scss';
import { IGraphToolkitProps } from './IGraphToolkitProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-agenda': any;
      'template': any;
      'mgt-person': any;
    }
  }
}
export interface IGraphToolkitState {
  showPanel: boolean;
}
export default class GraphToolkit extends React.Component<IGraphToolkitProps, IGraphToolkitState> {
  public state = {
    showPanel: false
  };

  public render(): React.ReactElement<IGraphToolkitProps> {

    return (
      <div>
        <DefaultButton
          text="My events"
          iconProps={{ iconName: 'Calendar' }}
          onClick={this._showPanel}
        />

        <Panel
          isBlocking={false}
          isOpen={this.state.showPanel}
          onDismiss={this._hidePanel}
          type={PanelType.medium}
          headerText="Non-Modal Panel"
          closeButtonAriaLabel="Close"
        >
          <div className={styles.container}>
            <div className={styles.row}>

              <mgt-agenda
                event-query="me/calendar/events"
              >
              </mgt-agenda>
            </div>
          </div>
        </Panel>
      </div>


    );
  }
  private _showPanel = (): void => {
    this.setState({ showPanel: true });
  }

  private _hidePanel = (): void => {
    this.setState({ showPanel: false });
  }
}
