import { IInputs } from "../generated/ManifestTypes";
import * as mockData from '../mockRecords.json'
export class MultiselectModel {
    /**
     * Retrieves data if the filter is of type: ?$filter
     */
    public static GetDataFromApi(context: ComponentFramework.Context<IInputs>, entityName: string, filter: string, showRecordsToBeShown: number) {
        return context.webAPI.retrieveMultipleRecords(entityName, filter)
            .then(function (results) {
                // If the results are == or less than 50 GO FOR IT
                if (results != null && results.entities != null && results.entities.length > 0 && results.entities.length <= showRecordsToBeShown) {
                    return results?.entities;
                }
                 // If results are more than 50 break
                 if (results.entities.length > showRecordsToBeShown) {
                    return -1;
                }
                // If no results
                return -2;
            })
    }

    /**
     * Retrieves data if the filter is of type <fetchxml...>
     */
    public static GetDataFromApiWithFetchXml(context: ComponentFramework.Context<IInputs>, entityName: string, fetchxml: string, showRecordsToBeShown: number) {
        return context.webAPI.retrieveMultipleRecords(entityName, "?fetchXml=" + fetchxml)
            .then(function (results) {
                // If the results are == or less than 50 GO FOR IT
                if (results != null && results.entities != null && results.entities.length > 0 && results.entities.length <= showRecordsToBeShown) {
                    return results?.entities;
                }
                // If results are more than 50 break
                if (results.entities.length > showRecordsToBeShown) {
                    return -1;
                }
                // If no results
                return -2;
            })
    }

    /**
     * Gets the data from the entity record where the control is set
     */
    public static GetDataFromEntity(context: ComponentFramework.Context<IInputs>, entityName: string, entityId: string, selects: string) {
        return context.webAPI.retrieveRecord(entityName, entityId, "?$select=" + selects)
            .then(function (result) {
                return result;
            })
    }

    /**
     * Mock data when this._isFake = true in index.ts
     */
    public static GetDataFromMock(filter: any) {
        var mockDataFiltered = mockData.filter(x => x.name.toLowerCase().includes(filter.toLowerCase()))
        return Promise.resolve(mockDataFiltered);

    }
}