import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';

export function renderDate(dateFormat: string): (item?: any, index?: number, column?: IColumn) => JSX.Element {

    return (item?: any, index?: number, column?: IColumn) => {


        console.log(`in col render selected date format ois ${dateFormat} local selected date format is  `);
        var date: Date = new Date(item[column.fieldName]);
        var displayDate = (dateFormat === "UTC") ? date.toUTCString() : date.toLocaleString();

        return (
            <div>
                { displayDate}
            </div>
        );
    };
}