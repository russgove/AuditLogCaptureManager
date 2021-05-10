import { AadHttpClient } from '@microsoft/sp-http';
import { QueryClient } from 'react-query';
import { BaseComponentContext } from '@microsoft/sp-component-base';
export interface IAuditLogCaptureManagerState {
  currentAction: string;
  selectedDateFormat: string;
  aadHttpClient: AadHttpClient;
  managementApiUrl: string;
  queryClient: QueryClient;
  auditItemContentTypeId: string;
  webPartContext: BaseComponentContext;
}
