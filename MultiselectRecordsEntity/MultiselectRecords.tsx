import * as React from "react";

import { Stack, IDetailsRowProps, IRenderFunction, CommandBarButton, PrimaryButton, IIconProps, TextField, DetailsList, DetailsListLayoutMode, Selection, SelectionMode, initializeIcons, Spinner, SpinnerSize } from 'office-ui-fabric-react/lib';
import { textFieldStyles } from './MultiselectRecords.styles'
const clearIcon: IIconProps = { iconName: 'Clear' };

export class MultiselectRecords extends React.Component<any> {
    private _selection: Selection;
    private _allItems: [];
    private _columns: any;
    private _textFieldValue: string;
    private _loading: boolean;
    private _selectedItems: any;
    private _selectedRecordsItems: []
    private _searchValue: string;
    constructor(props: any) {
        super(props);
        initializeIcons();
        this._allItems = this.props.records;
        this._columns = this.props.columns;
        this._textFieldValue = this.props.inputValue || "";
        this._searchValue = "";
        this._selectedItems = [];
        this._selectedRecordsItems = [];
        this._selection = new Selection({
            onSelectionChanged: () => { this._getSelectionDetails() }
        });
        this.state = {
            records: this.props.records
        }
    }

    public componentDidUpdate(prevProps: any): void {

        if (this.props != prevProps) {
            this.setState((prevState: any): any => {
                this._allItems = []
                this._selection.setAllSelected(false);
                prevState.userInput = this.props.userInput;
                this._selectedRecordsItems.forEach((item, index) => {
                    this._allItems.push(item)
                })
                const propsRecords: [] = this.props.records;
                propsRecords.forEach((item) => {
                    const index = this._allItems.findIndex(x => x[this.props.data] === item[this.props.data]);
                    if (index == -1) {
                        this._allItems.push(item)
                    }
                })
                setTimeout(() => {
                    this.selectIndexFromNames();
                }, 0)
                this._loading = false;
                return prevState;
            });
        }
    }

    public componentDidMount(): void {
        if (this._textFieldValue != "") {
            setTimeout(() => {
                this.selectIndexFromNames();
            }, 0)
        }
    }

    private onRenderRow(props: IDetailsRowProps, defaultRender?: IRenderFunction<IDetailsRowProps>): JSX.Element {
        return (
            <div data-selection-toggle="true">
                {defaultRender && defaultRender(props)}
            </div>
        );
    };

    private _showMainTextField(): JSX.Element {
        if (this.props.isControlVisible) {
            return (
                <TextField className={"text"}
                    onChange={this.userInputOnChange}
                    autoComplete="off"
                    value={this._textFieldValue}
                    styles={textFieldStyles}
                    disabled={this.props.isControlDisabled}
                    placeholder="---"
                />
            );
        } else {
            return (
                <></>
            );
        }
    }

    private _showSearchTextField(): JSX.Element {
        if (this.props.isControlVisible) {
            return (
                <TextField className={"text"}
                    onChange={this.filterRecords}
                    autoComplete="off"
                    value={this._searchValue}
                    styles={{ root: { flex: 1, position: 'relative', marginTop: 10 } }}
                    disabled={this.props.isControlDisabled}
                />
            );
        } else {
            return (
                <></>
            );
        }
    }

    private _showButtonSelectElements(): JSX.Element {
        if (this._allItems.length > 0) {
            return (
                <PrimaryButton text="Select Elements" allowDisabledFocus onClick={this.setFieldValue} />
            );
        } else {
            return (
                <></>
            );
        }
    }

    private _showDetailsList(): JSX.Element {

        if (this._allItems.length > 0) {

            return (
                <Stack>
                    <CommandBarButton iconProps={clearIcon} text="Close" onClick={this.clearItems} styles={{
                        root: {
                            flex: 1,
                            width: 100,
                            padding: 10,
                            zIndex: 1995,
                            backgroundColor: 'lightgrey'
                        }
                    }} />

                    <DetailsList
                        isHeaderVisible={this.props.headerVisible}
                        items={this._allItems}
                        columns={this._columns}
                        setKey="set"
                        selection={this._selection}
                        selectionMode={SelectionMode.multiple}
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionPreservedOnEmptyClick={true}
                        ariaLabelForSelectionColumn="Toggle selection"
                        checkButtonAriaLabel="Checkbox"
                        onRenderRow={this.onRenderRow}

                        styles={{
                            root: {
                                flex: 1,
                                position: 'absolute',
                                zIndex: 1995,
                                boxShadow: "rgba(0, 0, 0, 0.133) 0px 0px 7.2px 0px, rgba(0, 0, 0, 0.11) 0px 0px 1.8px 0px"
                            }
                        }}
                    /></Stack>
            )
        } else {
            return (
                <></>
            )
        }
    }

    public render(): JSX.Element {

        let spinner;

        /**
         * If _allItems is more than 0 then we will create the list.
         * _allItems will populate once it request from fetch
         */
        if (this._loading) {
            spinner = <Spinner size={SpinnerSize.medium} styles={{
                root: { margin: 10 }
            }} />
        }

        return (
            <div className={"divContainer"}>
                <div className={"control"} >
                    <Stack style={{ flexDirection: 'row' }}>
                        {this._showMainTextField()}
                        {this._showButtonSelectElements()}

                    </Stack>
                    <Stack style={{ flexDirection: 'row' }}>
                        {this._showSearchTextField()}
                    </Stack>
                    {spinner}
                    {this._showDetailsList()}

                </div>
            </div>
        );
    }

    private userInputOnChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
        // Get the target
        const target = event.target as HTMLTextAreaElement;
        //Set the value of our textfield to the input
        this._textFieldValue = target.value;
        //This is needed for loading the textFieldValue
        this.setState((prevState: any): any => prevState);
    }

    /**
     * Everytime is triggered the onKeyUp it will trigger this functionality
     */
    private filterRecords = async (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): Promise<any> => {
        // Get the target
        const target = event.target as HTMLTextAreaElement;
        //Set the value of our textfield to the input
        this._searchValue = target.value;
        // Await for the fetchRequest to response
        this._loading = false;
        //This is needed for loading the textFieldValue
        this.setState((prevState: any): any => prevState);
        // Send filter outside
        this.props.triggerFilter(this._searchValue)
    };

    private _getSelectionDetails = (): void => {
        const listSelection: any = this._selection.getSelection();
        const listArray: any = Array.isArray(listSelection) ? listSelection : [listSelection];
        const arrayItems: [] = listArray;
        this._selectedItems = [];
        this._selectedRecordsItems = []
        for (let item of arrayItems) {
            this._selectedRecordsItems.push(item);
            this._selectedItems.push(item[this.props.data]);
        }
    }

    private setFieldValue = (): void => {
        const valueToBeAssigned: string = this._selectedItems.join(this.props.delimiter);
        this._textFieldValue = valueToBeAssigned;
        this.props.eventOnChangeValue(valueToBeAssigned);
    }

    private selectIndexFromNames = (): void => {
        var values = this._textFieldValue.split(this.props.delimiter);
        for (var item of values) {
            var index = this._allItems.findIndex(x => x[this.props.data] == item);
            this._selection.setIndexSelected(index, true, true);
        }
    }

    private clearItems = (): void => {
        this._allItems = [];
        this.setState((prevState: any): any => prevState);


    }

}
