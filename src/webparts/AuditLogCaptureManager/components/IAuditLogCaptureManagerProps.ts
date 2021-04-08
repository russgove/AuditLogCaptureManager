import { AadHttpClient } from '@microsoft/sp-http';
import { QueryClient, QueryClientProvider, useQuery } from 'react-query';
export interface IAuditLogCaptureManagerProps {
  aadHttpClient: AadHttpClient;
  managementApiUrl: string;
  queryClient: QueryClient;
}
