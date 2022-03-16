import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import MultiselectRecords  from './MultiselectRecords'
import { MultiselectModel } from './Model/MultiselectRetrieveData';
import { Utilities } from './Utilities/Utilities';
export class MultiselectRecordsEntity implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _filterTags: string;
	private _numberOfRecordsToBeShown: number;
	private _isMultiple: any;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}
	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// reference to the notifyOutputChanged method
	private _notifyOutputChanged: () => void;
	// reference to the container div
	private _container: HTMLDivElement;
	private _records: any;
	private _entityName: string;
	private _isFetchXml: boolean;
	private _headerVisible: "True" | "False";

	private _filter: string;
	private _value: string;
	private _originalFilter: string;
	private _isFake: boolean;
	private _entityRecordId: string;
	private _entityRecordName: string;
	private _filterDynamicValues: string;
	private _populatedFieldVisible: "True" | "False";
	private props: any = {
		records: [],
		eventOnChangeValue: this.eventOnChangeValue.bind(this),
		triggerFilter: this.triggerFilter.bind(this),
		inputValue: "",
		columns: "",
		headerVisible: false,
		data: "",
		attributeid: "",
		isControlDisabled: false,
		isControlVisible: true,
		selectedRecords: [],
		populatedFieldVisible: true,
		logicalName: "",
		openWindow: "",
		widthProp: 200,
		heightProp: 500,
		filterTags: true,
		context: null,
		isMultiple: true
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code

		this._context = context;
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;
		this.props.inputValue = this._context.parameters.field.raw || "";
		this._entityName = this._context.parameters.entityName.raw || "";
		this._filter = this._context.parameters.filter.raw || "";
		this._originalFilter = this._context.parameters.filter.raw || "";

		this._isFetchXml = this._originalFilter.includes("fetch") ? true : false;
		this._headerVisible = this._context.parameters.headerVisible.raw || "False";

		this.props.headerVisible = this._headerVisible == "True" ? true : false;
		const columns: string = this._context.parameters.columns.raw?.trim() || "Name,name";
		this.props.columns = Utilities.parseColumns(columns);
		this.props.data = this._context.parameters.data.raw || "name";
		this.props.attributeid = this._context.parameters.attributeid.raw || "accountid";
		this._filterDynamicValues = this._context.parameters.filterDynamicValues.raw || "";
		const contextPage = (context as any).page;
		this._entityRecordId = contextPage.entityId;
		this._entityRecordName = contextPage.entityTypeName;

		this._isFake = false;
	
		this._populatedFieldVisible = this._context.parameters.populatedFieldVisible.raw || "False";
		this.props.populatedFieldVisible = this._populatedFieldVisible == "True" ? true : false;

		this.props.logicalName = this._entityName;
		this.props.openWindow = this._context.parameters.openFormOptions.raw || "Pop up";
		this._filterTags = this._context.parameters.filterTags.raw || "True";
		this.props.filterTags = this._filterTags == "True" ? true : false;
		this._numberOfRecordsToBeShown = this._context.parameters.recordsToBeReturned.raw || 50;
		this.props.numberIfRecordsToBeShown = this._numberOfRecordsToBeShown;
		const _isMultiple = this._context.parameters.isMultiple.raw || "True";
		this.props.isMultiple = _isMultiple == "True" ? true : false;
		context.mode.trackContainerResize(true);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {
		// Add code to update control view
		this.props.inputValue = this._context.parameters.field.raw || null;
		this.props.isControlDisabled = context.mode.isControlDisabled;
		this.props.isControlVisible = context.mode.isVisible;
		this.props.widthProp = context.mode.allocatedWidth - 20;
		this.props.heightProp = context.mode.allocatedHeight;
		this.props.context = context;
		this.renderElement()
	}

	/**
	 * Retrieves the records when the filter has triggered
	 */
	private async retrieveRecordsFromFetch(input: string): Promise<any> {
		if (this._isFake == false) {
			if (this._isFetchXml == true) {
				this._records = await MultiselectModel.GetDataFromApiWithFetchXml(this._context, this._entityName, this._filter, this._numberOfRecordsToBeShown);
			} else {
				this._records = await MultiselectModel.GetDataFromApi(this._context, this._entityName, this._filter, this._numberOfRecordsToBeShown);
			}
		} else {
			this._records = await MultiselectModel.GetDataFromMock(input);
		}
		this.props.records = this._records;
		if (this.props.inputValue === null) {
			this.props.inputValue = "";
		}
		return this.props.records
	}

	/**
	 * Event when the main field is changed
	 */
	private async eventOnChangeValue(newValue: string) {
		if (this.props.inputValue !== newValue) {
			this._value = newValue;
			this._notifyOutputChanged();
		}
	}

	/**
	 * Event when the search box has changed
	 */
	private async triggerFilter(newInput: string) {
		this._filter = this._originalFilter.replace(/\{0\}/g, newInput);
		if (this._filterDynamicValues != "") {
			if (this._isFake == false) {
				this._filter = await this.filteredUrlFromDynamicValues(this._filter);
			}
		}
		var records = await this.retrieveRecordsFromFetch(newInput);
		return records;
	}

	/**
	 * Parsed and fill the new fulter with the dynamic values from the entity record
	 */
	private async filteredUrlFromDynamicValues(_filter: string): Promise<string> {
		const arrayDynamicValues: string[] = this._filterDynamicValues.split(",");
		const result = await MultiselectModel.GetDataFromEntity(this._context, this._entityRecordName, this._entityRecordId, this._filterDynamicValues);
		arrayDynamicValues.forEach((value: string, index: number) => {
			index++;
			var apiValue = result[value];
			var replaceindex = `${index}`;
			var regex = new RegExp("\\{" + replaceindex + "\\}", "g")
			_filter = _filter.replace(regex, apiValue);
		});
		return _filter;
	}

	/**
	 * Method to render the component
	 */
	private renderElement(): void {
		ReactDOM.render(
			React.createElement(
				MultiselectRecords,
				this.props
			),
			this._container
		);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			field: this._value
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		ReactDOM.unmountComponentAtNode(this._container);
	}
}
