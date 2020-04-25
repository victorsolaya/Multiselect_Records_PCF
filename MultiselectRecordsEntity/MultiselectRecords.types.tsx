export interface IColumnObject {
    key: string,
    name: string,
    fieldName: string,
    minWidth: number,
    maxWidth?: number,
    isResizable: boolean
}

export interface IMultiselectProps {
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
    isControlDisabled: boolean
}