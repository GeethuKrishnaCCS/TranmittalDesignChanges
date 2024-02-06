import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as Constant from "../shared/Constants";
import { SPFI, SPFx } from "@pnp/sp";
import { PagedItemCollection } from '@pnp/sp/items';
export class BaseService {
    private _transmittalSP: SPFI;

    constructor(context: WebPartContext,hubUrl:string) {
        this._transmittalSP = new SPFI(hubUrl).using(SPFx(context));
    }
    public async getHubListItems(listname: string): Promise<any[]> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._transmittalSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname)
                    .items
                    .top(250)
                    .getPaged();
            }
            else {
                items = await items.getNext();
            }
            if (items.results.length > 0) {
                finalItems = finalItems.concat(items.results);
            }
        } while (items.hasNext);

        return finalItems;
    }
    public async getUserMessages(listname: string, pageName: string): Promise<any[]> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._transmittalSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname)
                    .items
                    .top(250)
                    .getPaged();
            }
            else {
                items = await items.getNext();
            }
            if (items.results.length > 0) {
                finalItems = finalItems.concat(items.results);
            }
        } while (items.hasNext);

        return finalItems;
    }
    public async getHubUserType(listname: string, UserEmail: string,hubUrl): Promise<any[]> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._transmittalSP.web.getList(hubUrl + "/Lists/" + listname)
                    .items.filter("Title eq '" + UserEmail + "'")
                    .top(250)
                    .getPaged();
            }
            else {
                items = await items.getNext();
            }
            if (items.results.length > 0) {
                finalItems = finalItems.concat(items.results);
            }
        } while (items.hasNext);

        return finalItems;
    }
    public async getHubItemsWithFilter(listname: string,filter:string,hubUrl:string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._transmittalSP.web.getList(hubUrl + "/Lists/" + listname)
                    .items
                    .filter(filter)
                    .top(250)
                    .orderBy("Title", true)
                    .getPaged();
            }
            else {
                items = await items.getNext();
            }
            if (items.results.length > 0) {
                finalItems = finalItems.concat(items.results);
            }
        } while (items.hasNext);
        return finalItems;
    }
    public createNewProcess(data: any, listname: string): Promise<any> {
        return this._transmittalSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname)
            .items.add(data);
    }
    public updateItem(data: any, id: number, listname: string): Promise<any> {
        return this._transmittalSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.getById(id).update(data);
    }
    public getTriggerUrl(listname: string, Title: string): Promise<any> {
        return this._transmittalSP.web.getList(Constant.hubsiterelurl + "/Lists/" + listname).items.filter("Title eq '" + Title + "'")()
    }
}