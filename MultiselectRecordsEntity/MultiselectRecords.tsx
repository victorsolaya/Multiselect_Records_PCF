import * as React from "react";

import { Stack, IDetailsRowProps, IRenderFunction, CommandBarButton, PrimaryButton, IIconProps, TextField, DetailsList, DetailsListLayoutMode, Selection, SelectionMode, initializeIcons, Spinner, SpinnerSize, IconButton } from 'office-ui-fabric-react/lib';
import { textFieldStyles } from './MultiselectRecords.styles'
import { IMultiselectProps } from './MultiselectRecords.types'
import { useEffect, useState, useRef } from "react";
import { Utilities } from './Utilities/Utilities';

const MultiselectRecords = (props: IMultiselectProps) => {
    const refSearchInput = useRef(null);
    const listRef = useRef(null);
    const clearIcon: IIconProps = { iconName: 'Clear' };
    const acceptIcon: IIconProps = { iconName: 'Accept', styles: { root: { color: 'white' } } };
    const searchIcon: IIconProps = { iconName: 'Search' };

    let timeout: any = 0;
    initializeIcons();
    // STATE
    const [showList, setShowList] = useState(false);
    const [listItems, setListItems] = useState(props.records);
    const [textFieldValue, setTextFieldValue] = useState(props.inputValue || "");
    const [searchValue, setSearchValue] = useState("");
    const [selectedRecordItems, setSelectedRecordItems] = useState([])
    const [selectedItems, setSelectedItems] = useState([])

    const [selection, setSelection] = useState(new Selection());
    const [myItems, setMyItems] = useState([]);
    const [records, setRecords] = useState(props.records !== -1 ? props.records : []);
    const [isError, setIsError] = useState(0)
    const [errorMessage, setErrorMessage] = useState("")
    
    // EFFECT
    useEffect(() => {
        setItemsFromTextFieldValue();
        selectRecordItemsAndPushThem();
    }, [textFieldValue]);
    useEffect(() => {
        setItemsFromRecordsIfNull();
    }, [showList, myItems]);
    useEffect(() => {
        async function DoOnLoad() {
            const recordsRetrieved: any = await getRecordsFromTextField();
            if(recordsRetrieved.entities.length > 0) {
                setSelectedItems(recordsRetrieved.entities);
            }
        }
        DoOnLoad();
    },[])

    useEffect(() => {
        const errorMessageText = isError === 0 ? "" : isError === -1 ? `More than ${props.numberIfRecordsToBeShown} records have been retrieved. Please, make it less or equals ${props.numberIfRecordsToBeShown}.` : "There are no records to be shown";
        setErrorMessage(errorMessageText);
    }, [isError])

    // FUNCTIONS
const getRecordsFromTextField = async () => {
    let filter = `$filter=`;
        const fieldValueArray = getTextFieldJSON();
        if(fieldValueArray.length == 0) {
            return null;
        }
        for(let fieldValue of fieldValueArray) {
            if(fieldValue != null) {
                filter += `${props.attributeid} eq ${JSON.parse(fieldValue)["id"]} or `
            }
        }
        filter = filter.substring(0, filter.length - 4);
        const addFilter = filter == "$filter=" ? "" : `&${filter}`;
        const recordsRetrieved: any = await Xrm.WebApi.retrieveMultipleRecords(props.logicalName, `?$select=${props.attributeid},${props.data}${addFilter}`)
        return recordsRetrieved;
}

    const setSelectedItemsWhenOpened = async(recordsPassedParam: any = null) => {
        const recordsRetrieved: any = await getRecordsFromTextField();
        const recordsPassedParamCopy = recordsPassedParam.slice();
        
        if(recordsPassedParamCopy != null || recordsRetrieved != null && recordsRetrieved.entities.length != 0 ) {
            const selectedItemsWhenOpened: any = recordsRetrieved != null ? recordsRetrieved.entities : [];
            for (var item of selectedItemsWhenOpened) {
                var itemFiltered = recordsPassedParamCopy.filter((x: any) => x[props.attributeid] == item[props.attributeid]);
                var index = recordsPassedParamCopy.findIndex((x: any) => x[props.attributeid] == item[props.attributeid]);
                if(index != -1) {
                    recordsPassedParamCopy.splice(index, 1);
                    recordsPassedParamCopy.unshift(itemFiltered[0]);
                }
            }
            //setSelectedItems(selectedItemsWhenOpened);
            setListItems(recordsPassedParamCopy);
            selection.setItems(recordsPassedParamCopy);
            let selectedItemsConcat:any = selectedItemsWhenOpened.concat(selectedItems);
            selectedItemsConcat = selectedItemsConcat.filter((a: any, b: any) => selectedItemsConcat.indexOf(a) === b)

            for(let item of selectedItemsConcat) {
                const indexItem = recordsPassedParamCopy.findIndex((x:any) => item[props.attributeid] == x[props.attributeid])
                if(indexItem != -1) {
                    selection.setIndexSelected(parseInt(indexItem), true, true);
                }
            }
        }
    }

    const setItemsFromRecordsIfNull = () => {
        if(listItems.length == 0) {
            setListItems(records);
        }
        
    }
    const setItemsFromTextFieldValue = () => {
        let valuesInTheTextField: any = []
        if(textFieldValue != null && textFieldValue != "" && textFieldValue != "[]") {
            let textFieldTemp = JSON.parse(textFieldValue);
            for(var item of textFieldTemp) {
                valuesInTheTextField.push(JSON.stringify(item));
            }
        }
        setMyItems(valuesInTheTextField);
    }

    const selectRecordItemsAndPushThem = () => {
        selection.setAllSelected(false);
        let itemsSelected: any = []
        if(selectedRecordItems != null && selectedRecordItems.length > 0) {
            for(var selectedRecordItem of selectedRecordItems) {
                itemsSelected.push(selectedRecordItem);
            }
        } else {
            if(textFieldValue != "" && textFieldValue != "[]") {
                let textFieldTemp = JSON.parse(textFieldValue);
                for(var item of textFieldTemp) {
                    itemsSelected.push(JSON.stringify(item));
                }
                setMyItems(itemsSelected)
            }
        }
        selectIndexFromNames()
    }


    /**
     * Method to select item when you click on the row
     */
    const onRenderRow = (props: IDetailsRowProps, defaultRender?: IRenderFunction<IDetailsRowProps>): JSX.Element => {
        return (
            <div data-selection-toggle="true" onClick={() => onClickRow}>
                {defaultRender && defaultRender(props)}
            </div>
        );
    };

    const onClickRow = async (item: any, index: number, event: React.FocusEvent<HTMLElement>) => {
        const row = event.target;
        let selectedItemsCopy: any = selectedItems;
        if(row.classList.contains("is-selected")) {
            row.classList.remove("is-selected");
            selectedItemsCopy = selectedItems.filter( (x:any) => x[props.attributeid] != item[props.attributeid])
        } else {
            row.classList.add("is-selected");
            selectedItemsCopy = selectedItems;
            selectedItemsCopy.push(item)
        }
        //Remove duplicates
        if(selectedItemsCopy.length > 0) {
            selectedItemsCopy = selectedItemsCopy.filter((a: any, b: any) => selectedItemsCopy.indexOf(a) === b)
        }
        setSelectedItems(selectedItemsCopy);
    }

    /**
     * Renders the main text
     */
    const _showMainTextField = (): JSX.Element => {
        if (props.isControlVisible) {
            return (
                <TextField className={"text"}
                    onChange={userInputOnChange}
                    autoComplete="off"
                    value={textFieldValue}
                    styles={textFieldStyles}
                    disabled={props.isControlDisabled}
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
     * Renders the main text
     */
    const _showSecondaryTextField =(): JSX.Element => {
        if (props.isControlVisible) {
            if(textFieldValue != ""  && textFieldValue != "[]") {
                return (
                    <Stack gap="5" horizontal wrap maxWidth={props.widthProp}>
                        {myItems != null && myItems.length > 0 && myItems.map( (item: any ) => {
                            const theItem = JSON.parse(item);
                            return (
                                <Stack horizontal style={{border:"1px solid #106EBE"}}>
                                    <PrimaryButton className="buttonContainer" style={{borderRadius: 0}} key={theItem.id} data-id={theItem.id} text={theItem.name} onClick={triggerItemClick} />
                                    <IconButton primary iconProps={clearIcon} title="Clear" ariaLabel="Clear" onClick={removeFieldValue}/>
                                </Stack>
                            )
                        })}
                    </Stack>
                );
        }
    }
        return (
            <></>
        );
        
    }

    /**
     * Renders the search box
     */
    const _showSearchTextField = (): JSX.Element => {
        if (props.isControlVisible) {
            return (
                <>
                    <TextField className={"text"}
                        componentRef={refSearchInput}
                        onChange={filterRecords}
                        width={props.widthProp}
                        autoComplete="off"
                        styles={{ root: { flex: 1, position: 'relative', marginTop: 10 } }}
                        disabled={props.isControlDisabled}
                        placeholder="Search..."
                        errorMessage={errorMessage}
                        onKeyUp={enterFilterRecords}
                    />
                    <PrimaryButton
                     iconProps={searchIcon}
                      title="Search"
                       ariaLabel="Search" 
                       onClick={filterRecordsClick} 
                       cellPadding={0}
                       width={40}
                       
                       styles={{ root: { position: 'relative', marginTop: 10, padding:0, minWidth:40 } }}
/>
                </>
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
    const _showDetailsList = (): JSX.Element => {

        if (listItems.length > 0 && showList == true) {

            return (
                <Stack>
                    <div style={{display: 'flex', justifyContent: 'space-between', flex: 1, alignContent: 'space-between' }}>
                        <CommandBarButton iconProps={acceptIcon} text="Select Elements" onClick={setFieldValue} styles={{
                            root: {
                                flex: 1,
                                padding: 10,
                                zIndex: 1995,
                                backgroundColor: '#0078D4',
                                color: 'white',
                                textAlign: 'left'
                            }
                        }} />

                        <CommandBarButton iconProps={clearIcon} text="Close" onClick={clearItems} styles={{
                            root: {
                                flex: 1,
                                padding: 10,
                                zIndex: 1995,
                                backgroundColor: 'lightgrey',
                                textAlign: 'left'
                            }
                        }} />

                    </div>
                    <DetailsList
                        isHeaderVisible={props.headerVisible}
                        items={listItems}
                        columns={props.columns}
                        setKey="set"
                        selection={selection}
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionPreservedOnEmptyClick={true}
                        ariaLabelForSelectionColumn="Toggle selection"
                        checkButtonAriaLabel="Checkbox"
                        onRenderRow={onRenderRow}
                       
                        componentRef={listRef}
                        styles={{
                            root: {
                                width: props.width,
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

    const triggerItemClick = (a:any):void => {
        const dataid = a.currentTarget.getAttribute("data-id")
        openRecord(props.logicalName, dataid); 
    }

    const openRecord =(logicalName: string, id: string): void => {
        const version = Xrm.Utility.getGlobalContext()
          .getVersion()
          .split(".");
        const mobile = Xrm.Utility.getGlobalContext().client.getClient() == "Mobile";
        // MFD (main form dialog) is available past ["9", "1", "0000", "15631"]
        // But doesn't work on mobile client
        if (
          !mobile &&
          version.length == 4 &&
          Number.parseFloat(version[0] + "." + version[1]) >= 9.1 &&
          Number.parseFloat(version[2] + "." + version[3]) >= 0.15631
        ) {
            switch (props.openWindow.toLowerCase())
            {
                case "no action":
                    break;
                case "in a new window":
                    (Xrm.Navigation as any).openForm({
                        entityName: logicalName,
                        entityId: id,
                        openInNewWindow: true
                        });
                break;
                case "in the same window":
                    (Xrm.Navigation as any).openForm({
                        entityName: logicalName,
                        entityId: id,
                        openInNewWindow: false
                        });
                    break;
                default:
                case "in a pop up":
                    (Xrm.Navigation as any).navigateTo(
                        {
                        entityName: logicalName,
                        pageType: "entityrecord",
                        formType: 2,
                        entityId: id,
                        },
                        { target: 2, position: 1, width: { value: 80, unit: "%" } },
                    );
                    break;
                }
        } else {
            if(props.openWindow.toLowerCase() != "no action") {
                Xrm.Navigation.openForm({
                    entityName: logicalName,
                    entityId: id,
                    });
            }
        }
    } 

    /**
     * When the main field is changed
     */
    const userInputOnChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
        // Get the target
        const target = event.target as HTMLTextAreaElement;
        //Set the value of our textfield to the input
        setTextFieldValue(target.value)
        //This is needed for loading the textFieldValue
        props.eventOnChangeValue(textFieldValue);
    }

    /**
     * Main trigger when the searchbox is changed
     */
    const filterRecords = (event: React.FormEvent | null): void => {
        //Set the value of our textfield to the input
        if(event != null) {
            if((event.target as any).value == "" || (event.target as any).value == null) {
                setIsError(0);
                setShowList(false)
                return;
            }
        }
        clearTimeout(timeout);
        timeout = setTimeout(async () => {
            await filterRecordsClick();
        }, 500);
    };

    const filterRecordsClick = async () => {
        const searchInputRef: any = refSearchInput.current;
            let searchInputValue = searchInputRef.value
            setSearchValue(searchInputValue);

            // Send filter outside
            const recordsRetrieved: any = await props.triggerFilter(searchInputValue);
            const myItemsFromTextField = getTextFieldJSON()
            if(props.filterTags == true) {
                if(myItemsFromTextField.length > 0 && searchInputValue != "" && searchInputValue != null && recordsRetrieved != -1 && recordsRetrieved != -2) {
                    const intersection = myItemsFromTextField.filter((x:any) => recordsRetrieved.some((y:any) => y[props.attributeid] === JSON.parse(x).id))
                    setMyItems(intersection);
                } else if(searchInputValue == "" || searchInputValue == null) {
                    setMyItems(myItemsFromTextField);
                }
            }
            
            if(recordsRetrieved != -1 && recordsRetrieved != -2) {
                setIsError(0)
                // setListItems(recordsRetrieved);
                // selection.setItems(recordsRetrieved)
                
                setRecords(recordsRetrieved);
                setSelectedItemsWhenOpened(recordsRetrieved);
                if(showList == false) {
                    setShowList(true)
                }
                
            } else {
                setListItems([]);
                selection.setItems([])
                const numberOfError = recordsRetrieved
                setIsError(numberOfError);
                setShowList(false)
            }
    }

    const enterFilterRecords = (event: any): void => {
        if(event.key === 'Enter'){
            filterRecords(null);
        }
    }
        
    /**
     * Event when the select elements is clicked
     */
    const setFieldValue = (): void => {
        const valueToBeAssigned: any = fillSelectedItems()[0];
        setMyItems(valueToBeAssigned);
        setTextFieldValue("[" + valueToBeAssigned.toString() + "]")
        props.eventOnChangeValue("[" + valueToBeAssigned.toString() + "]") ;
        setIsError(0);
        setShowList(false);
    }

    const getTextFieldJSON = () => {
        let myItemsCopy: any = [];
        if(textFieldValue != "" && textFieldValue != "[]") {
            let textFieldTemp = JSON.parse(textFieldValue);
            for(var item of textFieldTemp) {
                myItemsCopy.push(JSON.stringify(item));
            }
        }
        return myItemsCopy
    }

    /**
     * Click on the remove button next to the tag
     */
    const removeFieldValue = (event: any): void => {
        const selectedRecordItemsCopy = fillSelectedItems()[1];
        const id = event.currentTarget.parentElement.getElementsByClassName('buttonContainer')[0].getAttribute("data-id");
        const filterItemsWithoutTheRemovedOne: any = selectedRecordItemsCopy.filter( (myItem: any) => myItem[props.attributeid] != id).map((x: any) => x);
        const filterSelectedItemsWithoutTheRemovedOne: any = selectedItems.filter( (myItem: any) => myItem[props.attributeid] != id).map((x: any) => x);

        setSelectedRecordItems(filterItemsWithoutTheRemovedOne)
        setSelectedItems(filterSelectedItemsWithoutTheRemovedOne)
        setSearchValue("")
        const filteredText = JSON.parse(textFieldValue).filter( (myItem: any) => myItem["id"] != id).map((x: any) => x)
        const filteredTextString = filteredText.map((x:any) => JSON.stringify(x));
        const text: any = JSON.stringify(filteredText);
        setTextFieldValue(text)
        setMyItems(filteredTextString);
        setIsError(0);
        setShowList(false)
        props.triggerFilter("")
        props.eventOnChangeValue(text);
    }

    
    const removeDuplicatesFromArray = (arr: []) => [...new Set(
        arr.map(el => JSON.stringify(el))
      )].map(e => JSON.parse(e));

    /**
     * Method to fill the selected items from the box
     */
    const fillSelectedItems = (): any => {
        
        let listSelection: any = selectedItems ;
        listSelection = listSelection.concat(selection.getSelection());
        const listArray: any = Array.isArray(listSelection) ? listSelection : [listSelection];
        const arrayItems:any[] = removeDuplicatesFromArray(listArray);
          
        setSelectedRecordItems([])
        let selectedRecordItemsCopy: any = [];
        let selectedItemsCopy: any = [];

        for (let item of arrayItems) {
            
            selectedRecordItemsCopy.push(item);
            let json = { "id": item[props.attributeid], "name": item[props.data]}
            
            selectedItemsCopy.push(JSON.stringify(json));
        }
        setSelectedRecordItems(selectedRecordItemsCopy);
        return [selectedItemsCopy, selectedRecordItemsCopy];
    }

    /**
     * Selects the rows
     */
    const selectIndexFromNames = (recordsProp: any = null): void => {
        if(textFieldValue != "") {
            if(!Utilities.isJson(textFieldValue)) {
                return;
            }
            var values = JSON.parse(textFieldValue)
            const arrayAllItems = recordsProp != null ? recordsProp : listItems != null && listItems.length > 0 ? listItems : recordsProp;
            if(arrayAllItems != null && arrayAllItems.length > 0) {

                for (var item of values.reverse()) {
                    var itemFiltered = arrayAllItems.filter((x: any) => x[props.attributeid] == item["id"]);
                    var index = arrayAllItems.findIndex((x: any) => x[props.attributeid] == item["id"]);
                    if(index != -1) {
                        arrayAllItems.splice(index, 1);
                        arrayAllItems.unshift(itemFiltered[0]);
                    }
                }
                
                const copy = arrayAllItems.slice();
                selection.setItems([],true);
                setListItems(copy.slice());
                selection.setItems(copy.slice(),true);

                for(let item of selectedItems) {
                    const indexItem = copy.findIndex((x:any) => item[props.attributeid] == x[props.attributeid])
                    if(indexItem != -1) {
                        selection.setIndexSelected(parseInt(indexItem), true, true);
                    }
                }
            
                
            }
        }
    }

    /**
     * When close button is triggered
     */
    const clearItems = (): void => {
        setSearchValue("")
        setListItems([])
        const copyItems = getTextFieldJSON();
        setMyItems(copyItems);
        setIsError(0);
        setShowList(false)
        props.triggerFilter("")
    }

     /**
         * If _allItems is more than 0 then we will create the list.
         * _allItems will populate once it request from fetch
         */
        return (
            <div className={"divContainer"}>
                <div className={"control"} >
                    {props.populatedFieldVisible == true ? <Stack horizontal>
                        {_showMainTextField()}
                    </Stack> : <></>}
                    <Stack horizontal>
                        {_showSecondaryTextField()}
                    </Stack>
                    <Stack horizontal>
                        {_showSearchTextField()}
                    </Stack>
                    {_showDetailsList()}

                </div>
            </div>
        );
}

export default MultiselectRecords;
