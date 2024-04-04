import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import { PagedItemCollection } from '@pnp/sp/items';

export class BaseService {
    private _sp: SPFI;
    private sphub: SPFI;

    constructor(context: WebPartContext, huburl: string) {
        this._sp = getSP(context);
        this.sphub = new SPFI(huburl).using(SPFx(context));
    }

    public getListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items();
    }
    public gethubUserMessageListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("Title,Message").filter("PageName eq 'InboundSib-Contractor'")()
    }
    public getLibraryItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items();
    }
    public getCurrentUser(): Promise<any> {
        return this._sp.web.currentUser();
    }
    
    public createhubNewItem(url: string, listname: string, data: any): Promise<any> {
        console.log(data);
        return this.sphub.web.getList(url + "/Lists/" + listname).items.add(data);
    }
   
    public updatehubItem(url: string, listname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this.sphub.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public updateLibraryItem(url: string, libraryname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/" + libraryname).items.getById(id).update(data);
    }
    public uploadDocument(libraryName: string, Filename: any, filedata: any): Promise<any> {
        return this._sp.web.getFolderByServerRelativePath(libraryName).files.addUsingPath(Filename, filedata, { Overwrite: true });
    }
    public getDocument(Url: string, publisheddocumentLibrary: string, publishName: string): Promise<any> {
        return this._sp.web.getFileByServerRelativePath(Url + "/" + publisheddocumentLibrary + "/" + publishName).getBuffer()
    }
    public getDrpdwnListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("Title,ID")()
    }
    public getRevisionListItems(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).select("ID,StartPrefix,Pattern,StartWith,EndWith,MinN,MaxN,AutoIncrement")()
    }
    public getByEmail(email: string): Promise<any> {
        return this._sp.web.siteUsers.getByEmail(email)()
    }
    public getByhubEmail(email: string): Promise<any> {
        return this.sphub.web.siteUsers.getByEmail(email)()
    }
    public getByUserId(id: any): Promise<any> {
        return this.sphub.web.siteUsers.getById(id)()
    }
    public getHubsiteData(): Promise<any> {
        return this._sp.web.hubSiteData()
    }
    public getListItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public getLibraryItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.getById(id)();
    }
    public gethubItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public getApproverData(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("ID,Title,Approver/Title,Approver/ID,Approver/EMail").expand("Approver")()
    }
    public getIndexData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID).select("WorkflowStatus,SourceDocument,DocumentStatus")();
    }
    public getIndexDataId(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID)
            .select("DocumentID,DocumentName,DepartmentID,BusinessUnitID,Owner/ID,Owner/Title,Owner/EMail,Approver/ID,Approver/Title,Approver/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID,Reviewers/ID,Reviewers/Title,Reviewers/EMail").expand("Owner,Approver,Reviewers")();
    }
    public getIndexProjectData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID)
            .select("RevisionCodingId,RevisionLevelId,TransmittalRevision,AcceptanceCodeId,DocumentController/ID,DocumentController/Title,DocumentController/EMail").expand("DocumentController")();
    }
    public getRevisionLevelData(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("ID,Title")()
    }
    public getSourceLibraryItems(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.filter('DocumentIndexId eq ' + ID)()
    }
    public getpreviousheader(url: string, listname: string, IndexID: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("ID").filter("DocumentIndex eq '" + IndexID + "' and(WorkflowStatus eq 'Returned with comments')")();
    }
    public gettriggerUnderReviewPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPermission_UnderReview'")()
    }
    public gettriggerUnderApprovalPermission(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPermission_UnderApproval'")()
    }
    public getdirectpublish(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_PermissionWebpart'")()
    }
    public getnotification(url: string, listname: string, emailuser: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference")()
    }
    public getemail(url: string, listname: string, type: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq '" + type + "'")();
    }
    public gettaskdelegation(url: string, listname: string, Id: number): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + Id + "' and(Status eq 'Active')")();
    }
    public gettransmittaloutlooklibraryitem(url: string, libraryname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + libraryname).items
        .select("ID,BaseName,SubContractor").filter("From eq 'Sub-Contractor'")();
    }
    public getIndexItems(url: string, listname: string, id: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.filter("ID eq '" + id + "'")();
    }
    public getdocumentIndexItem(url: string, listname: string, Id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items
        .select("Owner/Title,Owner/ID,RevisionCoding/Title,RevisionCoding/ID,DocumentID,Title,SubcontractorDocumentNo,DocumentName")
        .expand("Owner,RevisionCoding")
        .filter("ID eq '" + Id + "'")();
    }
    public gettransmittaloutlooklibraryitemName(url: string, libraryname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + libraryname).items
        .select("LinkFilename,ID")();
    }
    public gettransmittaloutlooklibraryitemBuffer(fileurl: string): Promise<any> {
        return this._sp.web.getFileByServerRelativePath(fileurl).getBuffer();
    }
    public deleteListItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).delete();
    }
    public updateListItem(url: string, listname: string, data: any, id: number): Promise<any> {
        console.log(data);
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public getTransmittalDetailList(url: string, listname: string, Id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items
        .select("TransmittalHeaderId,DocumentIndex/ID,DocumentIndex/Title,DocumentIndex/DocumentName,Owner/ID,Owner/Title,SubContractorDocumentNumber,ReceivedDate,Comments,ID")
      .expand("DocumentIndex,Owner")
      .filter("TransmittalHeader/ID eq '" + Id + "' ")();
    }
    public deleteLibraryItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.getById(id).delete();
    }
    public getadditionaldocumentItems(url: string, listname: string, id: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.filter("TransmittalIDId eq '" + id + "'")();
    }
    public getIdSettingsItem(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.filter("TransmittalCategory eq 'Inbound Sub-contractor'")();
    }
    public createNewItem(url: string, listname: string, data: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.add(data);
    }
    public gethubpermissionListItems(url: string, listname: string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items.filter("Title eq 'EMEC_DocumentPermission-Create Document'")();
    }
    public getinboundHeader(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(ID).select("TransmittalStatus")();
    }
    public getinboundHeaderData(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.select("Id,Title,TransmittalDate,DocumentController/ID,DocumentController/Title,SubContractorID,SubContractor")
        .expand("DocumentController")();
    }
    public gettransmittalOutlookLibraryData(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items
        .filter("From eq 'Sub-Contractor'").select("ID,BaseName,SubContractor")();
    }
    public getinboundAdditionalDocumentsData(url: string, listname: string, ID: any): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items
        .filter("TransmittalIDId eq '" + ID + "' ")();
    }
    public gethubSubcontractorListItems(url: string, listname: string,projectNumber:string): Promise<any> {
        return this.sphub.web.getList(url + "/Lists/" + listname).items
        .filter("ProjectId eq '" + projectNumber + "'")();
    }
    public async getListItemsPaged(url: string, listname: string): Promise<any[]> {
        let finalItems: any[] = [];
        let items: PagedItemCollection<any[]> = undefined;
        do {
            if (!items) {
                items = await this._sp.web.getList(url + "/Lists/" + listname)
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
    
} 