import { AadHttpClient } from '@microsoft/sp-http';
import { QueryClient, QueryClientProvider, useQuery } from 'react-query';

export interface IAuditLogCaptureManagerState {
  currentAction: string;
  selectedDateFormat: string;
  aadHttpClient: AadHttpClient;
  managementApiUrl: string;
  queryClient: QueryClient;
  auditItemContentTypeId: string;
}
