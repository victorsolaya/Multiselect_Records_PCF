import { IInputs } from "./generated/ManifestTypes";

export interface IColumnObject {
    key: string,
    name: string,
    fieldName: string,
    minWidth: number,
    maxWidth?: number,
    isResizable: boolean
}

export interface IMultiselectProps {
    openWindow: any;
    width: string | number | undefined;
    eventOnChangeValue(textFieldValue: string): void;
    triggerFilter(searchValue: string) : void;
    attributeid: any;
    data: any;
    requestUrl: string,
    authorizationToken: string,
    userInput?: string;
    userInputChanged: (newValue: string) => void,
    labid: string,
    widthColumn: number,
    columns: [IColumnObject],
    inputValue?: string,
    dataToBeSet: string,
    headerVisible: boolean,
    isControlVisible: boolean,
    isControlDisabled: boolean,
    populatedFieldVisible: boolean,
    records: any,
    logicalName: string,
    widthProp: number,
    heightProp: number,
    filterTags: boolean,
    numberIfRecordsToBeShown: number,
    context: ComponentFramework.Context<IInputs>
}