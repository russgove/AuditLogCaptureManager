import { AadHttpClient } from '@microsoft/sp-http';

export interface IAuditLogCaptureManagerProps {
  aadHttpClient: AadHttpClient;
  managementApiUrl: string;
}
