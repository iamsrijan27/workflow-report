export default interface IService {
    getFieldDropdownValues(listName: string, fieldName: string): Promise<any>;
    startWorkflowRequest(context: any, postURL: string, itemId: string, selectedColumn: string, adminComments: string): Promise<any>;
    getAllItems(listName: string): Promise<any>;
    getAllItemsBounce(listName: string): Promise<any>;
    deleteItemByID(listName: string, id: any): Promise<any>;
    getSPGroupUsers(groupName: string): Promise<string[]>;
}