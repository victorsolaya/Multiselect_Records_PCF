import { IInputs } from "../generated/ManifestTypes";
import * as mockData from '../mockRecords.json'
export class MultiselectModel {
    public static GetDataFromApi(context: ComponentFramework.Context<IInputs>, entityName: string, filter: string) {
        return context.webAPI.retrieveMultipleRecords(entityName, filter)
            .then(function (results) {
                if (results != null && results.entities != null && results.entities.length > 0 && results.entities.length <= 50) {
                    return results?.entities;
                }
                return [];
            })
    }

    public static GetDataFromApiWithFetchXml(context: ComponentFramework.Context<IInputs>, entityName: string, fetchxml: string) {
        return context.webAPI.retrieveMultipleRecords(entityName, "?fetchXml=" + fetchxml)
            .then(function (results) {
                if (results != null && results.entities != null && results.entities.length > 0 && results.entities.length <= 50) {

                    return results?.entities;
                }
                return [];
            })
    }

    public static GetDataFromEntity(context: ComponentFramework.Context<IInputs>, entityName: string, entityId: string, selects: string) {
        return context.webAPI.retrieveRecord(entityName, entityId, "?$select=" + selects)
            .then(function (result) {
                return result;
            })
    }

    public static GetDataFromMock() {
        return Promise.resolve(mockData);
    }
}