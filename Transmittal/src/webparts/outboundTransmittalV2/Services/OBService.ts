import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";
import { PagedItemCollection } from '@pnp/sp/items';

export class OBService extends BaseService {
    private _spfi: any;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = SPFI;
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
        return this._spfi.web.currentUser.get();
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
    public async addToList(siteUrl: string, listName: string, items: any): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listName).items.add(items);
    }
    public async updateList(siteUrl: string, listName: string, items: any, itemId: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listName).items.getById(itemId).update(items);
    }
    public async updateLibrary(libraryUrl: string, items: any, itemId: number): Promise<any> {
        return this._spfi.web.getList(libraryUrl).items.getById(itemId).update(items);
    }
    public async getItemForSelectInListsWithFilter(siteUrl: string, listName: string, filter: string,): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
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
    public async getItemForSelectExpandInListsWithFilter(siteUrl: string, listName: string, select: string, filter: string, expand: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .select(select)
                    .filter(filter)
                    .expand(expand)
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
    public async getItemWithSelectAndExpand(siteUrl: string, listName: string, select: string, expand: string): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .select(select)
                    .expand(expand)
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
    public async getItemWithSelectAndExpandWithId(siteUrl: string, listName: string, select: string, expand: string, Id: any): Promise<any> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._spfi.web.getList(siteUrl + "/Lists/" + listName)
                    .items
                    .select(select)
                    .expand(expand)
                    .getById(Id)
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

    public async getlistItemById(siteUrl: string, listName: string, id: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listName).items.getById(id)();
    }

    public async uploadDocument(filename: string, filedata: any, libraryname: string, metadata: any): Promise<any> {
        const file = await this._spfi.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });
        const item = await file.file.getItem();
        item.update({
            Title: metadata.Title,
            TransmittalIDId: metadata.TransmittalIDId,
            Size: metadata.Size,
            Comments: metadata.Comments,
            TransmittalStatus: metadata.TransmittalStatus,
            Slno: metadata.Slno,
            SentDate: metadata.SentDate,
        });
        return file;
    }
}


