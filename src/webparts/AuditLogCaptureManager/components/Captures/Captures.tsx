import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IViewField, ListView } from '@pnp/spfx-controls-react/lib/controls/listView'
import * as React from 'react';
import { useState } from 'react';

export const ListItemsWebPartContext = React.createContext<WebPartContext>(null);
export interface ICapturesProps {
    description: string;
}
export const Captures: React.FunctionComponent<ICapturesProps> = (props) => {
    const viewFields: IViewField[] = [
        {
            name: 'Name', linkPropertyName: 'Name', sorting: true, isResizable: true
        }
    ];

    const [captures, setCaptures] = useState<Array<any>>();
    React.useEffect(() => {                           // side effect hook
        // call API with props.greeting parameter
        setCaptures([{ Name: "S" }, { Name: "SDWS" }, { Name: "SSSs" }])
    }, []);// empty array only runs when component is mounted
    return (
        <div>
            Captures
            <ListView items={captures} viewFields={viewFields}></ListView>
        </div>

    );
};
