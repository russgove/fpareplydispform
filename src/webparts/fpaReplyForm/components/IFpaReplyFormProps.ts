import { HttpClient } from '@microsoft/sp-http';
export interface IFpaReplyFormProps {
  description: string;
  title: string;
  httpClient: HttpClient;
  webPartId: string;
}
