/// see https://github.com/SharePoint/sp-dev-docs/blob/master/docs/spfx/web-parts/guidance/call-microsoft-graph-from-your-web-part.md
import * as React from 'react';
import styles from './FpaReplyForm.module.scss';
import { IFpaReplyFormProps } from './IFpaReplyFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as AuthenticationContext from 'adal-angular';
import adalConfig from '../AdalConfig';
import { IAdalConfig } from '../../IAdalConfig';
import * as junk from '../WebPartAuthenticationContext';
export default class FpaReplyForm extends React.Component<IFpaReplyFormProps, void> {
  
  public render(): React.ReactElement<IFpaReplyFormProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p className="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p className="ms-font-l ms-fontColor-white">{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
