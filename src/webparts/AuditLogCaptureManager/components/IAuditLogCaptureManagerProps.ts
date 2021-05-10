import { AadHttpClient } from '@microsoft/sp-http';
import { QueryClient, QueryClientProvider, useQuery } from 'react-query';
import { BaseComponentContext } from '@microsoft/sp-component-base';
export interface IAuditLogCaptureManagerProps {
  aadHttpClient: AadHttpClient;
  managementApiUrl: string;
  queryClient: QueryClient;
  auditItemContentTypeId: string;
  webPartContext: BaseComponentContext;
}
