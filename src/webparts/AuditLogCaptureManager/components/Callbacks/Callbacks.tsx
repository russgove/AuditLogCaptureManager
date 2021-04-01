import { CutomPropertyContext } from '../AuditLogCaptureManager'
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import { useState } from 'react';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICallbacksProps {
  description: string;
}
export const Callbacks: React.FunctionComponent<ICallbacksProps> = (props) => {
  debugger;
  const parentcontext: any = React.useContext<any>(CutomPropertyContext)
  const [lists, setLists] = useState<string | string[]>('');

  return (
    <div>
      Callbacks Go Here!
      {parentcontext.managementApiUrl}
    </div>

  );
};
