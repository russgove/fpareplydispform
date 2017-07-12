/// see https://github.com/SharePoint/sp-dev-docs/blob/master/docs/spfx/web-parts/guidance/call-microsoft-graph-from-your-web-part.md
import * as React from 'react';
import styles from './FpaReplyForm.module.scss';
import { IFpaReplyFormProps } from './IFpaReplyFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as AuthenticationContext from 'adal-angular';
import adalConfig from '../AdalConfig';
import { IAdalConfig } from '../../IAdalConfig';
import * as junk from '../WebPartAuthenticationContext';
import { IMeeting } from './IMeeting';
import { ICalendarMeeting } from './ICalendarMeeting';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ListItem } from './ListItem';

export interface IFpaReplyFormState {
  loading: boolean,
  error: string,
  upcomingMeetings: Array<IMeeting>
  signedIn: boolean,
}
export default class FpaReplyForm extends React.Component<IFpaReplyFormProps, IFpaReplyFormState> {

  private authCtx: adal.AuthenticationContext;
  constructor(props: IFpaReplyFormProps, context?: any) {
    super(props);
    debugger;
    this.state = {
      loading: false,
      error: null,
      upcomingMeetings: [],
      signedIn: false
    };

    const config: IAdalConfig = adalConfig;
    config.popUp = true;
    config.webPartId = this.props.webPartId;
    config.callback = (error: any, token: string): void => {
      this.setState((previousState: IFpaReplyFormState, currentProps: IFpaReplyFormProps): IFpaReplyFormState => {
        previousState.error = error;
        previousState.signedIn = !(!this.authCtx.getCachedUser());
        return previousState;
      });
    };

    this.authCtx = new AuthenticationContext(config);
    window["Logging"] = {
      level: 4,
      log: function (message) {
        console.log("ADAL MESSAGE ::" + message);
      }
    };

    AuthenticationContext.prototype._singletonInstance = undefined;
  }
  public componentDidMount(): void {
    debugger;
    this.authCtx.handleWindowCallback();

    if (window !== window.top) {
      return;
    }

    this.setState((previousState: IFpaReplyFormState, props: IFpaReplyFormProps): IFpaReplyFormState => {
      previousState.error = this.authCtx.getLoginError();
      previousState.signedIn = !(!this.authCtx.getCachedUser());
      return previousState;
    });
  }
  public signIn(): void {
    debugger;
    this.authCtx.login();
  }
  public componentDidUpdate(prevProps: IFpaReplyFormProps, prevState: IFpaReplyFormState, prevContext: any): void {
    debugger;
    if (prevState.signedIn !== this.state.signedIn) {
      this.loadUpcomingMeetings();
    }
  }
  private static getDateTime(date: Date): string {
    return `${date.getHours()}:${FpaReplyForm.getPaddedNumber(date.getMinutes())}`;
  }

  public render(): React.ReactElement<IFpaReplyFormProps> {
    debugger;
    const login: JSX.Element = this.state.signedIn ? <div /> : <button className={`${styles.button} ${styles.buttonCompound}`} onClick={() => { this.signIn(); }}><span className={styles.buttonLabel}>Sign in</span><span className={styles.buttonDescription}>Sign in to see your upcoming meetings</span></button>;
    const loading: JSX.Element = this.state.loading ? <div style={{ margin: '0 auto', width: '7em' }}><div className={styles.spinner}><div className={`${styles.spinnerCircle} ${styles.spinnerNormal}`}></div><div className={styles.spinnerLabel}>Loading...</div></div></div> : <div />;
    const error: JSX.Element = this.state.error ? <div><strong>Error: </strong> {this.state.error}</div> : <div />;
    const meetingItems: JSX.Element[] = this.state.upcomingMeetings.map((item: IMeeting, index: number, meetings: IMeeting[]): JSX.Element => {
      return <ListItem key={index} item={
        {
          primaryText: item.subject,
          secondaryText: item.location,
          tertiaryText: item.organizer,
          metaText: FpaReplyForm.getDateTime(item.start),
          isUnread: item.status === 'busy'
        }
      }
        actions={[
          {
            icon: 'View',
            item: item,
            action: (): void => {
              window.open(item.webLink, '_blank');
            }
          }
        ]} />;
    });
    let meetings: JSX.Element = <div>{meetingItems}</div>;

    if (this.state.upcomingMeetings.length === 0 &&
      this.state.signedIn &&
      !this.state.loading &&
      !this.state.error) {
      meetings = <div style={{ textAlign: 'center' }}>No upcoming meetings: ) </div>;
    }

    return (
      <div className={styles.helloWorld}>
        <div className={'ms-font-xl ' + styles.webPartTitle}>{escape(this.props.title)}</div>
        {login}
        {loading}
        {error}
        {meetings}
      </div>
    );
  }
  private getGraphAccessToken(): Promise<string> {
    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {
      const graphResource: string = 'f8f8d2ad-7c9d-4aac-80eb-3f00a263c879';
      const accessToken: string = this.authCtx.getCachedToken(graphResource);
      if (accessToken) {
        resolve(accessToken);
        return;
      }

      if (this.authCtx.loginInProgress()) {
        reject('Login already in progress');
        return;
      }

      this.authCtx.acquireToken(graphResource, (error: string, token: string) => {
        if (error) {
          reject(error);
          return;
        }

        if (token) {
          resolve(token);
        }
        else {
          reject('Couldn\'t retrieve access token');
        }
      });
    });
  }
  private loadUpcomingMeetings(): void {
    this.setState((previousState: IFpaReplyFormState, props: IFpaReplyFormProps): IFpaReplyFormState => {
      previousState.loading = true;
      return previousState;
    });

    this.getGraphAccessToken()
      .then((accessToken: string): Promise<IMeeting[]> => {
        return FpaReplyForm.getUpcomingMeetings(accessToken, this.props.httpClient);
      })
      .then((upcomingMeetings: IMeeting[]): void => {
        this.setState((prevState: IFpaReplyFormState, props: IFpaReplyFormProps): IFpaReplyFormState => {
          prevState.loading = false;
          prevState.upcomingMeetings = upcomingMeetings;
          return prevState;
        });
      }, (error: any): void => {
        this.setState((prevState: IFpaReplyFormState, props: IFpaReplyFormProps): IFpaReplyFormState => {
          prevState.loading = false;
          prevState.error = error;
          return prevState;
        });
      });
  }
  private static getUpcomingMeetings(accessToken: string, httpClient: HttpClient): Promise<IMeeting[]> {
    return new Promise<IMeeting[]>((resolve: (upcomingMeetings: IMeeting[]) => void, reject: (error: any) => void): void => {
      const now: Date = new Date();
      const dateString: string = now.getUTCFullYear() + '-' + FpaReplyForm.getPaddedNumber(now.getUTCMonth() + 1) + '-' + FpaReplyForm.getPaddedNumber(now.getUTCDate());
      const startDate: string = dateString + 'T' + FpaReplyForm.getPaddedNumber(now.getUTCHours()) + ':' + FpaReplyForm.getPaddedNumber(now.getUTCMinutes()) + ':' + FpaReplyForm.getPaddedNumber(now.getUTCSeconds()) + 'Z';
      const endDate: string = dateString + 'T23:59:59Z';

      httpClient.get(`https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${startDate}&endDateTime=${endDate}&$orderby1=Start&$select=id,subject,start,end,webLink,isAllDay,location,organizer,showAs`, HttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata.metadata=none',
          'Authorization': 'Bearer ' + accessToken
        }
      })
        .then((response: HttpClientResponse): Promise<{ value: ICalendarMeeting[] }> => {
          return response.json();
        })
        .then((todayMeetings: { value: ICalendarMeeting[] }): void => {
          const upcomingMeetings: IMeeting[] = [];

          for (let i: number = 0; i < todayMeetings.value.length; i++) {
            const meeting: ICalendarMeeting = todayMeetings.value[i];
            const meetingStartDate: Date = new Date(meeting.start.dateTime + 'Z');
            if (meetingStartDate.getDate() === now.getDate()) {
              upcomingMeetings.push(FpaReplyForm.getMeeting(meeting));
            }
          }
          resolve(upcomingMeetings);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private static getMeeting(calendarMeeting: ICalendarMeeting): IMeeting {
    return {
      id: calendarMeeting.id,
      subject: calendarMeeting.subject,
      start: new Date(calendarMeeting.start.dateTime + 'Z'),
      end: new Date(calendarMeeting.end.dateTime + 'Z'),
      webLink: calendarMeeting.webLink,
      isAllDay: calendarMeeting.isAllDay,
      location: calendarMeeting.location.displayName,
      organizer: `${calendarMeeting.organizer.emailAddress.name} <${calendarMeeting.organizer.emailAddress.address}>`,
      status: calendarMeeting.showAs
    };
  }

  private static getPaddedNumber(n: number): string {
    if (n < 10) {
      return '0' + n;
    }
    else {
      return n.toString();
    }
  }

}
