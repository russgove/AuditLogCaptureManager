import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
export interface IDateFormatPickerProps {

    onFormatChange: (e: string) => void;
    selectedDateFormat: string;
}
export const DateFormatPicker: React.FunctionComponent<IDateFormatPickerProps> = (props): JSX.Element => {
 
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;

    const dateFormatChoices: Array<IChoiceGroupOption> = [
        { key: "UTC", text: "UTC" },
        { key: "Local", text: timezone }
    ];
    return (
        <ChoiceGroup label="Date Format" defaultSelectedKey={props.selectedDateFormat} options={dateFormatChoices}
            onChange={
                (e, option) => {
               
                    props.onFormatChange(option.key);
                    console.log(`date format is now ${option.key}`);
                }}></ChoiceGroup>
    );
};