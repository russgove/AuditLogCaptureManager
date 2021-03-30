import * as React from 'react';
import { useState } from 'react';

import { WebPartContext } from '@microsoft/sp-webpart-base';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICallbacksProps {
  description: string;
}
export const Callbacks: React.FunctionComponent<ICallbacksProps> = (props) => {
  const [lists, setLists] = useState<string | string[]>('');

  return (
    <div>
      Callbacks Go Here!
    </div>

  );
};
