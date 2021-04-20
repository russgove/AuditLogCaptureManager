import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
export interface IDateFormatPickerProps {
    selectedDateFormat: React.MutableRefObject<string>;
}
export const DateFormatPicker: React.FunctionComponent<IDateFormatPickerProps> = (props): JSX.Element => {
    debugger;
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;

    const dateFormatChoices: Array<IChoiceGroupOption> = [
        { key: "UTC", text: "UTC" },
        { key: "Local", text: timezone }
    ];
    return (
        <ChoiceGroup label="Date Format" defaultSelectedKey={props.selectedDateFormat.current} options={dateFormatChoices}
            onChange={
                (e, option) => {
                    debugger;
                    props.selectedDateFormat.current = option.key;
                    console.log(`date format is now ${option.key}`);
                }}></ChoiceGroup>
    );
};