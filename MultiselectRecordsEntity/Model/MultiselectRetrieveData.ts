import { IInputs } from "../generated/ManifestTypes";
import * as mockData from '../mockRecords.json'
export class MultiselectModel {
    public static GetDataFromApi(context: ComponentFramework.Context<IInputs>, entityName: string, filter: string) {
        return context.webAPI.retrieveMultipleRecords(entityName, filter)
            .then(function (results) {
                return results?.entities;
            })
    }

    public static GetDataFromApiWithFetchXml(context: ComponentFramework.Context<IInputs>, entityName: string, fetchxml: string) {
        return context.webAPI.retrieveMultipleRecords(entityName, "?fetchXml=" + fetchxml)
            .then(function (results) {
                return results?.entities;
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