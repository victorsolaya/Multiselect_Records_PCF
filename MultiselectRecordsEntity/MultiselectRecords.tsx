import * as React from "react";

import { Stack, IDetailsRowProps, IRenderFunction, CommandBarButton, PrimaryButton, IIconProps, TextField, DetailsList, DetailsListLayoutMode, Selection, SelectionMode, initializeIcons, Spinner, SpinnerSize } from 'office-ui-fabric-react/lib';
import { textFieldStyles } from './MultiselectRecords.styles'
const clearIcon: IIconProps = { iconName: 'Clear' };
const acceptIcon: IIconProps = { iconName: 'Accept', styles: { root: { color: 'white' } } };


export class MultiselectRecords extends React.Component<any> {
    private _selection: Selection;
    private _allItems: [];
    private _columns: any;
    private _textFieldValue: string;
    private _selectedItems: any;
    private _selectedRecordsItems: []
    private _searchValue: string;
    private _isError: boolean;
    constructor(props: any) {
        super(props);
        initializeIcons();
        this._allItems = this.props.records != -1 ? this.props.records : [];
        this._columns = this.props.columns;
        this._textFieldValue = this.props.inputValue || "";
        this._searchValue = "";
        this._selectedItems = [];
        this._selectedRecordsItems = [];
        this._selection = new Selection();
        this.state = {
            records: this.props.records != -1 ? this.props.records : []
        }
        this._isError = this.props.records != -1 ? false : true;
    }

    public componentDidUpdate(prevProps: any): void {

        if (this.props != prevProps) {
            this.setState((prevState: any): any => {
                this._allItems = []
                if (this.props.records !== -1) {
                    this._selection.setAllSelected(false);
                    this._selectedRecordsItems.forEach((item) => {
                        this._allItems.push(item)
                    })
                    const propsRecords: [] = this.props.records;
                    propsRecords.forEach((item) => {
                        const index = this._allItems.findIndex(x => x[this.props.data] === item[this.props.data]);
                        if (index == -1) {
                            this._allItems.push(item)
                        }
                    })
                    this._isError = false;
                    setTimeout(() => {
                        this.selectIndexFromNames();
                    }, 0)
                } else {
                    this._isError = true;
                }
                return prevState;
            });
        }
    }

    /**
     * When component mounts to select the indexes.
     */
    public componentDidMount(): void {
        if (this._textFieldValue != "") {
            setTimeout(() => {
                this.selectIndexFromNames();
            }, 0)
        }
    }

    /**
     * Method to select item when you click on the row
     */
    private onRenderRow(props: IDetailsRowProps, defaultRender?: IRenderFunction<IDetailsRowProps>): JSX.Element {
        return (
            <div data-selection-toggle="true">
                {defaultRender && defaultRender(props)}
            </div>
        );
    };

    /**
     * Renders the main text
     */
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

    /**
     * Renders the search box
     */
    private _showSearchTextField(): JSX.Element {
        if (this.props.isControlVisible) {
            return (
                <TextField className={"text"}
                    onChange={this.filterRecords}
                    autoComplete="off"
                    value={this._searchValue}
                    styles={{ root: { flex: 1, position: 'relative', marginTop: 10 } }}
                    disabled={this.props.isControlDisabled}
                    placeholder="Search..."
                    errorMessage={this._isError ? "More than 50 records have been retrieved. Please, make it less or equals 50." : ""}
                />
            );
        } else {
            return (
                <></>
            );
        }
    }

    /**
     * Renders the list and the buttons
     */
    private _showDetailsList(): JSX.Element {

        if (this._allItems.length > 0) {

            return (
                <Stack>
                    <div style={{ display: 'flex', justifyContent: 'space-between', flex: 1, alignContent: 'space-between' }}>
                        <CommandBarButton iconProps={acceptIcon} text="Select Elements" onClick={this.setFieldValue} styles={{
                            root: {
                                flex: 1,
                                width: 200,
                                padding: 10,
                                zIndex: 1995,
                                backgroundColor: '#0078D4',
                                color: 'white',
                                textAlign: 'left'
                            }
                        }} />

                        <CommandBarButton iconProps={clearIcon} text="Close" onClick={this.clearItems} styles={{
                            root: {
                                flex: 1,
                                width: 200,
                                padding: 10,
                                zIndex: 1995,
                                backgroundColor: 'lightgrey',
                                textAlign: 'left'
                            }
                        }} />

                    </div>
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

    /**
     * Main method to render
     */
    public render(): JSX.Element {
        /**
         * If _allItems is more than 0 then we will create the list.
         * _allItems will populate once it request from fetch
         */

        return (
            <div className={"divContainer"}>
                <div className={"control"} >
                    <Stack style={{ flexDirection: 'row' }}>
                        {this._showMainTextField()}

                    </Stack>
                    <Stack style={{ flexDirection: 'row' }}>
                        {this._showSearchTextField()}

                    </Stack>
                    {this._showDetailsList()}

                </div>
            </div>
        );
    }

    /**
     * When the main field is changed
     */
    private userInputOnChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
        // Get the target
        const target = event.target as HTMLTextAreaElement;
        //Set the value of our textfield to the input
        this._textFieldValue = target.value;
        //This is needed for loading the textFieldValue
        this.setState((prevState: any): any => prevState);
        this.props.eventOnChangeValue(this._textFieldValue);
    }

    /**
     * Main trigger when the searchbox is changed
     */
    private filterRecords = async (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): Promise<any> => {
        // Get the target
        const target = event.target as HTMLTextAreaElement;
        //Set the value of our textfield to the input
        this._searchValue = target.value;
        //This is needed for loading the textFieldValue
        this.setState((prevState: any): any => prevState);
        // Send filter outside
        this.props.triggerFilter(this._searchValue)
    };

    /**
     * Event when the select elements is clicked
     */
    private setFieldValue = (): void => {
        this.fillSelectedItems();
        const valueToBeAssigned: string = this._selectedItems.join(this.props.delimiter);

        this._textFieldValue = valueToBeAssigned;
        this.props.eventOnChangeValue(valueToBeAssigned);
    }

    /**
     * Method to fill the selected items from the box
     */
    private fillSelectedItems = (): void => {
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

    /**
     * Selects the rows
     */
    private selectIndexFromNames = (): void => {
        var values = this._textFieldValue.split(this.props.delimiter);
        for (var item of values) {
            var index = this._allItems.findIndex(x => x[this.props.data] == item);
            if (index !== -1) {
                this._selection.setIndexSelected(index, true, true);
            }
        }
        this.fillSelectedItems();
    }

    /**
     * When close button is triggered
     */
    private clearItems = (): void => {
        this._allItems = [];
        this._searchValue = "";
        this.setState((prevState: any): any => prevState);
    }

}
