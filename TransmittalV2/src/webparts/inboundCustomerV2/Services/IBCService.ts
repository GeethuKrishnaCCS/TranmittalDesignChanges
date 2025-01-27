
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, SPFx  } from "@pnp/sp";
import { PagedItemCollection } from '@pnp/sp/items';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { BaseService } from './BaseService';

export class IBCService extends BaseService {
    private _spfi: SPFI;    
    private ctx: WebPartContext;
    private _hublSP: SPFI;

    constructor(context: WebPartContext,hubUrl:string) {
        super(context,hubUrl);
        this.ctx = context;
        this._spfi =  getSP(this.ctx);
        this._hublSP = new SPFI(hubUrl).using(SPFx(context));
    }
    public async getHubItemsWithFilter(listname: string,filter:string,hubUrl:string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._hublSP.web.getList(hubUrl + "/Lists/" + listname)
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
    public async getListItems(siteUrl: string, listName: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
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
    public getItembyID(url: string, listname: string, id: number): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public getCurrentUserId(): Promise<any> {
        return this._spfi.web.currentUser();
    }
    public async getTransmitFor(siteUrl: string, listName: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .filter("AcceptanceCode eq false")
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
    public async getLibraryItems(url: string,): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(url)
                    .items
                    .select("FileLeafRef,ID,FileSizeDisplay,TransmittalDocument,TransmittalStatus,DocumentName,DocumentIndexId,WorkflowStatus,DocumentStatus,Category")
                    .filter("TransmittalStatus ne 'Ongoing' and (TransmittalDocument ne '" + false + "') and (DocumentStatus eq 'Active') and (WorkflowStatus eq 'Published')")
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
    public async getSourceLibraryItems(url: string,): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(url)
                    .items
                    .select("FileLeafRef,ID,FileSizeDisplay,TransmittalDocument,TransmittalStatus,DocumentName,AcceptanceCode/ID,AcceptanceCode/Title,CustomerDocumentNo,SubcontractorDocumentNo,DocumentStatus")
                    .expand("AcceptanceCode")
                    .filter("TransmittalStatus ne 'Ongoing' and (TransmittalDocument ne '" + false + "') and (WorkflowStatus ne 'Draft') and (WorkflowStatus ne 'Under Approval') and (WorkflowStatus ne 'Under Review') and (DocumentStatus eq 'Active')")
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
    public async getDIItems(siteUrl: string, listName: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .filter("TransmittalStatus ne 'Ongoing' and (TransmittalDocument ne '" + false + "') and (DocumentStatus eq 'Active') and (WorkflowStatus eq 'Published')")
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
    public async getDLItemById(DlUrl: string, id: number): Promise<any> {
        const SourcedocumentItem = await this._spfi.web.getList(DlUrl).items.getById(id)();
        return SourcedocumentItem;
    }
    public async getItemForSelectInDL(DlUrl: string, select: string, filter: string, expand: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(DlUrl)
                    .items
                    .select(select)
                    .expand(expand)
                    .filter(filter)
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
    public async getItemForSelectInLists(siteUrl: string, listName: string, select: string, filter: string,): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .select(select)
                    .filter(filter)
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
}


