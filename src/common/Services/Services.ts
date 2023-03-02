import { WebPartContext } from '@microsoft/sp-webpart-base';
import IService from './IServices';
import { sp } from '@pnp/sp';
import "@pnp/sp/views"
import "@pnp/sp/webs"
import "@pnp/sp/lists/web"
import "@pnp/sp/items"
import "@pnp/sp/fields"
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups"
import { IListItem } from '../../webparts/workFlowReport/Models/IListItem'
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';

export default class Service implements IService {
    public static setup(context: WebPartContext): void {
        sp.setup({
            spfxContext: {pageContext: context.pageContext}
        });
    }

    //Get choices for dropdown field
    public getFieldDropdownValues(listName: string, fieldName: string): Promise<any> {
        var ddValues: any[] = [];
        return new Promise<any>((resolve: (ddValues: any) => void, reject: (error: any) => void) => {
            //sp.web.lists.getByTitle('Category Contacts').views.getByTitle('All Items').fields().then((fields) => {
            sp.web.lists.getByTitle(listName).views.getByTitle('All Items').fields().then((fields) => {
                let internames: string[] = (fields as any).Items;
                let filterstring: string = internames.map(x => `(InternalName eq '${x}')`).join(` or `);
               // sp.web.lists.getByTitle('Category Contacts').fields.filter(filterstring).select('InternalName', 'Title').get().then((fieldsNames) => {
                sp.web.lists.getByTitle(listName).fields.filter(filterstring).select('InternalName', 'Title').get().then((fieldsNames) => {
                    fieldsNames.map(item => {
                        ddValues.push({
                            key: item.InternalName,
                            text: item.InternalName,
                            value: item.InternalName,
                            label: item.Title
                        });
                    });
                    if (ddValues.length > 0) {
                        ddValues = ddValues.sort((a, b) => (a.label > b.label) ? 1 : -1);
                    }
                    resolve(ddValues);
                }).catch(error => {
                    reject(error);
                });
            }).catch(error => {
                reject(error);
            });
        });
    }

    public startWorkflowRequest(context: any, postURL: string, itemId: string, selectedColumn: string, adminComments: string): Promise<any> {
        return new Promise<any>((resolve: (response: any) => void, reject: (error: any) => void) => {
            const body: string = JSON.stringify({ 'itemId': itemId.toString(), 'columnNames': selectedColumn, 'adminComments': adminComments, 'flowTypes': 2 });
            const requestHeaders: Headers = new Headers();
            requestHeaders.append('Content-type', 'application/json');
            const httpClientOptions: IHttpClientOptions = {
                body: body,
                headers: requestHeaders
            };
            context.httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions).then((response: HttpClientResponse) => {
                resolve(response.json());
            }).catch(error => {
                reject(error);
            });
        });
    }

    //Get All SharePoint list items
    // public getAllItems(listName: string, selectfields: string, expandFields: string): Promise<any> {
    public getAllItems(listName: string): Promise<any> {
        return new Promise<any>((resolve: (listItems: any) => void, reject: (error: any) => void) => {
            //sp.web.lists.getByTitle(listName).items.orderBy("WorkflowName",true).getAll().then((listItems: any) => {
            sp.web.lists.getByTitle(listName).items.top(4999).orderBy("WorkflowName",true).get().then((listItems: any) => {
            // sp.web.lists.getByTitle("Category Contacts").items.select(selectfields).expand(expandFields).getAll().then((listItems: any) => {
                resolve(listItems);
            }).catch(error => {
                reject(error);
            });
        });
    }
    //Get All SharePoint list items
    // public getAllItems(listName: string, selectfields: string, expandFields: string): Promise<any> {
        public getAllItemsBounce(listName: string): Promise<any> {
            return new Promise<any>((resolve: (listItems: any) => void, reject: (error: any) => void) => {
                sp.web.lists.getByTitle(listName).items.getAll().then((listItems: any) => {
                    resolve(listItems);
                }).catch(error => {
                    reject(error);
                });
            });
        }

    //Delete SharePoint list items by ID
    public deleteItemByID(listName: string, itemId: any): Promise<any> {
        return new Promise<any>((resolve: (items: any) => void, reject: (error: any) => void) => {
            sp.web.lists.getByTitle(listName).items.getById(itemId).delete().then((item: any) => {
                resolve(item);
            }).catch(error => {
                reject(error);
            });
        });
    }

    //Get SharePoint list items by ID
    public getItemByID(listName: string, itemId: any, selectfields: string, expandFields: string): Promise<any> {
        return new Promise<any>((resolve: (items: any) => void, reject: (error: any) => void) => {
            sp.web.lists.getByTitle(listName).items.getById(itemId).select(selectfields).expand(expandFields).get().then((item: any) => {
                resolve(item);
            }).catch(error => {
                reject(error);
            });
        });
    }

    //Get SharePoint Group Users Email
    public getSPGroupUsers(groupName: string): Promise<string[]> {
        let userEmails: string[] = [];
        return new Promise<any>((resolve: (currentUserGroups: any) => void, reject: (error: any) => void) => {
            sp.web.siteGroups.getByName(groupName).users.get().then(users => {
                users.forEach(user => {
                    userEmails.push(user.Email.toLowerCase());
                });
                resolve(userEmails);
            }).catch(error => {
                reject(error);
            });
        });
    }
}