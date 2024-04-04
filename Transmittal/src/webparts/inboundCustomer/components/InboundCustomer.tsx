
import * as React from 'react';
import styles from './InboundCustomer.module.scss';
import { DatePicker, Dropdown, IconButton, IDropdownOption, DialogType, IIconProps, Label, IButton, TextField, DefaultButton, MessageBar, Dialog, DialogFooter, PrimaryButton, ProgressIndicator } from 'office-ui-fabric-react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib';
//import { ReactTable } from "react-table";
import * as moment from 'moment';
import { IFile } from '@pnp/spfx-controls-react/lib/services/FileBrowserService.types';
import SimpleReactValidator from 'simple-react-validator';
import { MSGraphClientV3, HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IEmecInboundCustomerWpState, IInboundCustomerProps, ITableData2, projectData } from './IInboundCustomerProps';
import { Web } from '@pnp/sp/webs';
import { BaseService } from '../services/BaseService';

export default class EmecInboundCustomerWp extends React.Component<IInboundCustomerProps, IEmecInboundCustomerWpState, {}> {
  private validator: SimpleReactValidator;
  private _Service: BaseService;
  private invalidUser;
  private valid;
  private today;
  private transmittalID;
  private currentEmail;
  private currentId;
  private currentUser;
  private addindoc = [];
  private fileInput;
  private docidfilter = [];
  private postUrl;
  private SourceDocumentID;
  private typeForDelete;
  private postUrlForPermission;
  private file2;
  private keyfordelete;
  private flowUrlForDLUpdate;
  public reqWeb;
  private additionalfile;
  constructor(props: IInboundCustomerProps) {
    super(props);
    this.fileInput = React.createRef();
    this.state = {
      recallConfirmMsgDiv: "none",
      recallConfirmMsg: true,
      confirmDeleteDialog: true,
      deleteConfirmation: "none",
      confirmCancelDialog: true,
      outlookContractNumber: null,
      docaddselected: true,
      access: "none",
      accessDeniedMsgBar: "none",
      docselected: true,
      TransmittalHeaderId: null,
      statusKey: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      deleteConfirmMsg: "none",
      tempDocIndexIDForDelete: null,
      SourceDocumentID: null,
      DocumentID: null,
      ownerEmail: null,
      queryParamNo: false,
      queryParamYes: false,
      OwnerId: "",
      OwnerTitle: "",
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      currentinBoundDetailItem: [],
      currentInboundAdditionalItem: [],
      AddIndex2: true,
      projectdivVisible: false,
      transIdvisible: true,
      inboundTransmittalHeaderId: "",
      Attachments: null,
      Attachments2: null,
      incrementSequenceNumber: "",
      outlookCustomerID: "",
      outlookCustomer: "",
      outlookCustomerDocNo: "",
      outlookPONumber: "",
      todayDate: null,
      ReactTableResult: [],
      ReactTableResult2: [],
      AddIndex: true,
      documentIndexOption: [],
      TransmittalCodeSettings: [],
      receivedDate: null,
      receivedDate2: null,
      transmittalCode: "",
      comments: "",
      comments2: "",
      projectName: "",
      projectNumber: "",
      docId: "",
      docKey: "",
      transCodeKey: "",
      transcodeText: "",
      documentIndex: "",
      poNumber: "",
      transmittalID: "",
      transmittalStatus: "",
      btnsvisible: false,
      addDocsVisible: false,
      loaderDisplay: "none",
      webpartView: "",
      submitDisable: false
    };
    this._Service = new BaseService(this.props.context, window.location.protocol + "//" + window.location.hostname + this.props.hubSiteUrl);
    this.Addindex = this.Addindex.bind(this);
    this.Addindex2 = this.Addindex2.bind(this);
    this.AddDoc = this.AddDoc.bind(this);
    this.AddDoc2 = this.AddDoc2.bind(this);
    this._hideGrid = this._hideGrid.bind(this);
    this._bindData = this._bindData.bind(this);
    this._handleTableData = this._handleTableData.bind(this);
    this.transcodechange = this.transcodechange.bind(this);
    this._onreceivedDateChange = this._onreceivedDateChange.bind(this);
    this.DocIndex = this.DocIndex.bind(this);
    this.handleDeleteRow = this.handleDeleteRow.bind(this);
    this.handleDeleteRow2 = this.handleDeleteRow2.bind(this);
    this._trannsmittalIDGeneration = this._trannsmittalIDGeneration.bind(this);
    this._transmittalSequenceNumber = this._transmittalSequenceNumber.bind(this);
    this.submit = this.submit.bind(this);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this.bindInboundTransmittalSavedData = this.bindInboundTransmittalSavedData.bind(this);
    this._projectInformation = this._projectInformation.bind(this);
    this.saveAsDraft = this.saveAsDraft.bind(this);
    this.itemDeleteFromGrid = this.itemDeleteFromGrid.bind(this);
    this.itemDeleteFromExternalGrid = this.itemDeleteFromExternalGrid.bind(this);
    this._confirmDeleteItem = this._confirmDeleteItem.bind(this);
    this._LAUrlGettingForPermission = this._LAUrlGettingForPermission.bind(this);
    this.triggerProjectPermissionFlow = this.triggerProjectPermissionFlow.bind(this);
    this._LAUrlGettingForDocumentLibraryUpdate = this._LAUrlGettingForDocumentLibraryUpdate.bind(this);
    this.triggerProjectDLUpdate = this.triggerProjectDLUpdate.bind(this);
  }
  //Format Date
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  };
  //_LAUrlGettingForDocumentLibraryUpdate
  private _LAUrlGettingForDocumentLibraryUpdate = async () => {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.requestList).
      items.filter("Title eq 'EMEC_InboundTransmittal'").get();
    console.log("PosturlForPermission", laUrl[0].PostUrl);
    this.flowUrlForDLUpdate = laUrl[0].PostUrl;
    this.triggerProjectDLUpdate();
  }
  //Cancel
  private dialogCancelContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to Cancel?',
    //subText: '<b>Do you want to cancel? </b> ',
  };
  protected async triggerProjectDLUpdate() {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.flowUrlForDLUpdate;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteUrl': siteUrl,
      'TransmittalNo': this.transmittalID
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
  }
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields",
      }
    });
    this._projectInformation();
  }
  public async componentDidMount() {
    this._projectInformation();
    this.today = new Date();
    this.setState({ todayDate: moment(this.today).format('DD-MM-YYYY') });
    const user = await this._Service.getCurrentUser();
    this.currentEmail = this.props.context.pageContext.user.email;
    this.currentId = user.Id;
    this.currentUser = this.props.context.pageContext.user.displayName;
    this.reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubsite);
    this._bindData();
    this._queryParamGetting();
    this.setState({ access: "none", accessDeniedMsgBar: "none" });
  }
  private _LAUrlGettingForPermission = async () => {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.requestList).items.filter("Title eq 'EMEC_PermissionWebpart'").get();
    console.log("PosturlForPermission", laUrl[0].PostUrl);
    this.postUrlForPermission = laUrl[0].PostUrl;
    this.triggerProjectPermissionFlow(laUrl[0].PostUrl);
  }
  protected async triggerProjectPermissionFlow(PostUrl) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = PostUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': 'Project_SendTransmittal',
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.currentEmail
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] == "Valid") {
        this.setState({
          loaderDisplay: "none",
          webpartView: "",
        });
      }
      else {
        this.setState({
          webpartView: "none",
          loaderDisplay: "none",
          statusMessage: { isShowMessage: true, message: "You are not permitted to perform this operations", messageType: 4 },
        });
        setTimeout(() => {
          this.setState({ statusMessage: { isShowMessage: true, message: "You are not permitted to perform this operations", messageType: 1 } });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + "/" + this.props.siteUrl);
        }, 20000);
      }
    }
  }
  private _handleTableData(tableRowColl) {
    this.setState({ ReactTableResult: tableRowColl });
  }
  public _titleChange = (ev: React.FormEvent<HTMLInputElement>, comments?: string) => {
    this.setState({ comments: comments || '' });
  }
  public _titleChange2 = (ev: React.FormEvent<HTMLInputElement>, comments2?: string) => {
    this.setState({ comments2: comments2 || '' });
  }
  //Approval Date Change
  public _onreceivedDateChange = (date?: Date): void => {
    this.setState({ receivedDate: date });
  }
  public _onreceivedDateChange2 = (date?: Date): void => {
    this.setState({ receivedDate2: date });
  }
  public transcodechange(option: { key: any; text: any }) {
    this.setState({ transCodeKey: option.key, transcodeText: option.text });
    console.log(this.state.transCodeKey);
  }
  public async DocIndex(option: { key: any; text: any }) {
    this.setState({ docId: option.key, docKey: option.text });
    let select: "DocumentID,DocumentName,Owner/ID,Owner/Title,Owner/EMail,Revision,SourceDocument,CriticalDocument,CustomerDocumentNo";
    let expand: "Owner";
    const documentIndexItem: any = await this._Service.getItemSelectExpandById(this.props.siteUrl, this.props.documentIndexList, select, expand, option.key);
    console.log(documentIndexItem);
    this.setState({
      OwnerId: documentIndexItem.Owner.ID,
      OwnerTitle: documentIndexItem.Owner.Title,
      ownerEmail: documentIndexItem.Owner.EMail,
      outlookCustomerDocNo: documentIndexItem.CustomerDocumentNo
    });
  }
  public Addindex() {
    if ((this.state.docId == "") || (this.state.receivedDate == null) ||
      // (this.state.outlookCustomerDocNo == null) 
      (this.state.transCodeKey == "")
      || (this.state.outlookPONumber == "") || (this.state.comments == "" && this.state.Attachments == null)) {
      this.validator.showMessages();
      this.forceUpdate();
    }
    else {
      this.validator.hideMessages();
      this.setState({ AddIndex: false });
      this.AddDoc();
    }
  }
  public Addindex2() {
    if ((this.state.receivedDate2 != null)) {
      this.setState({ AddIndex2: false });
      this.AddDoc2();
      this.validator.hideMessages();
    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }
  }
  //Get Access Groups
  private async _accessGroups() {
    let accessGroup = [];
    let ok = "No";
    let filter: "Title eq 'Project_SendTransmittal'";
    let select: "AccessGroups,AccessFields";
    accessGroup = await this._Service.getItemWithSelectFilter(this.props.siteUrl, this.props.PermissionMatrixSettings, select, filter);
    if (accessGroup.length > 0) {
      let accessGroupItems: any[] = accessGroup[0].AccessGroups.split(',');
      this._gettingGroupID(accessGroupItems);

      console.log(accessGroupItems);
    }
  }
  private async _gettingGroupID(AccessGroupItems) {
    let AG;
    for (let a = 0; a < AccessGroupItems.length; a++) {
      AG = AccessGroupItems[a];
      const accessGroupID: any = await this._Service.getItemWithFilter(this.props.siteUrl, this.props.accessGroupDetailsList, "Title eq '" + AG + "'");
      let AccessGroupID;
      if (accessGroupID.length > 0) {
        console.log(accessGroupID);
        AccessGroupID = accessGroupID[0].GroupID;
        console.log("AccessGroupID", AccessGroupID);
        this.GetGroupMembers(this.props.context, AccessGroupID);
      }
    }
  }

  public async GetGroupMembers(context: WebPartContext, groupId: string): Promise<any[]> {
    let users: string[] = [];
    try {
      let client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
      let response = await client
        .api(`/groups/${groupId}/members`)
        .version('v1.0')
        .select(['mail', 'displayName'])
        .get();
      response.value.map((item: any) => {
        users.push(item);
      });
    } catch (error) {
      console.log('MSGraphService.GetGroupMembers Error: ', error);
    }
    console.log('MSGraphService.GetGroupMembers: ', users, "GroupID:", groupId);
    this._checkingCurrent(users);
    return users;
  }
  private async _checkingCurrent(userEmail) {
    for (var k in userEmail) {
      if (this.currentEmail == userEmail[k].mail) {
        this.setState({ access: "none", accessDeniedMsgBar: "none" });
        this.valid = "Yes";

        break;
      }
    }
    if (this.valid != "Yes") {

      this.setState({
        accessDeniedMsgBar: "", access: "none",
        statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/Lists/" + this.props.documentIndexList);
      }, 10000);
    }
  }
  public AddDoc() {
    let tableData = this.state.ReactTableResult;
    let data: projectData = {
      OwnerTitle: this.state.OwnerTitle,
      OwnerId: this.state.OwnerId,
      documentIndex: this.state.docKey,
      companyDocNumber: this.state.outlookCustomerDocNo,
      receivedDate: moment(this.state.receivedDate).format("DD/MM/YYYY"),
      receiveDate: this.state.receivedDate,

      transmittalCode: this.state.transcodeText,
      poNumber: this.state.outlookPONumber,
      comments: this.state.comments,
      Attachments: this.state.Attachments == null ? null : this.state.Attachments,
      transmittalID: this.state.transmittalID,
      docId: this.state.docId,
      transCodeKey: this.state.transCodeKey,
      docKey: this.state.docKey,
      ss: this.state.Attachments == null ? null : this.state.Attachments.name,
      url: "",
      DetailId: null
    };
    let dupIndex = tableData.filter(({ documentIndex }) => documentIndex === this.state.docKey);
    console.log(dupIndex);
    if (dupIndex.length < 1) {
      tableData.push(data);
      this.setState({ docselected: true });
    }
    else {
      this.setState({ docselected: false });
    }
    this.setState({
      ReactTableResult: tableData,
      Attachments: null,
      outlookCustomerDocNo: null,
      receivedDate: null,
      transCodeKey: null,
      //outlookPONumber: null,
      comments: "",
      docId: null
    });
    (document.querySelector("#studyAttachment") as HTMLInputElement).value = null;
  }
  public AddDoc2() {
    this.setState({ AddIndex2: false });
    let tableData2 = this.state.ReactTableResult2;
    let data2: ITableData2 = {
      documentName2: this.state.transmittalID + this.fileInput.current.files[0].name,
      comments2: this.state.comments2,
      additionalId: null,
      adddocurl: this.state.transmittalID + this.fileInput.current.files[0].name,
      receivedDate2: moment(this.state.receivedDate2).format("DD/MM/YYYY"),
      Attachments2: this.state.Attachments2 == null ? null : this.state.Attachments2,
      receiveDate2: this.state.receivedDate2,
    };
    tableData2.push(data2);
    this.setState({ ReactTableResult2: tableData2, });
    this.setState({
      receivedDate2: null,
      comments2: "",
      Attachments2: null,
    });
    (document.querySelector("#additionalfile") as HTMLInputElement).value = null;
  }
  public _CompanyDocChange = (ev: React.FormEvent<HTMLInputElement>, CompanyDocNo?: any) => {
    this.setState({ outlookCustomerDocNo: CompanyDocNo });
  }
  public _PoNumbChange = (ev: React.FormEvent<HTMLInputElement>, ponumber?: any) => {
    this.setState({ outlookPONumber: ponumber });
  }
  public handleDeleteRow(i) {
    let rows = this.state.ReactTableResult;
    rows.splice(i, 1);
    this.setState({
      ReactTableResult: rows
    });
    if (rows.length == 0) {
      this.setState({ AddIndex: true });
    }
    this.forceUpdate();
  }
  public handleDeleteRow2(i) {
    let rows = this.state.ReactTableResult2;
    rows.splice(i, 1);
    this.setState({ ReactTableResult2: rows });
    if (rows.length == 0) {
      this.setState({ AddIndex: true });
    }
  }
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,
      deleteConfirmation: "none",
      confirmDeleteDialog: true,
      recallConfirmMsg: true,
      recallConfirmMsgDiv: "none",
    });
    this.validator.hideMessages();
  }
  //confirm cancel button click
  private _cancelConfirmYes = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,

    });

    window.location.replace(window.location.protocol + "//" + window.location.hostname + "/" + this.props.siteUrl);
    this.validator.hideMessages();
  }
  private _hideGrid() {
    this.setState({
      confirmCancelDialog: false,
      cancelConfirmMsg: "",
    });
  }
  private handleFileUpload = (event) => {
    this.additionalfile = event.target.value;
    console.log(event.target.value);
    this.file2 = (document.getElementById("studyAttachment") as HTMLInputElement).files[0];
    this.setState({
      Attachments: this.file2
    });
  }
  private handleFileUpload2 = (event) => {

    let file3 = (document.getElementById("additionalfile") as HTMLInputElement).files[0];
    this.setState({
      Attachments2: file3
    });
  }
  //send mail
  public _sendmail = async (docid, emailuser, type, name, TransmittalAcceptanceCode) => {
    let formatday = moment(new Date()).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    const notificationPreference: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + "NotificationPreferenceSettings").items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference").get();
    if (notificationPreference.length > 0) {
      if (notificationPreference[0].Preference == "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document") {
        mailSend = "Yes";
      }
      else {
        mailSend = "Yes";
      }
    }
    if (mailSend == "Yes") {
      const emailNotification: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + "EmailNotificationSettings").items.get();
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title == type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }
      }
      // Subject = Subject.replace('[DocumentName]', docid);
      // Body = Body.replace('[DocumentName]', docid);
      let replacedSubject1 = replaceString(Subject, '[DocumentName]', docid);
      let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
      let replaceBody = replaceString(replaceRequester, '[DocumentName]', docid);
      let finalBody = replaceString(replaceBody, '[TransmittalAcceptanceCode]', TransmittalAcceptanceCode);
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject1,
          "body": {
            "contentType": "HTML",
            "content": finalBody
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": emailuser
              }
            }],
        }
      };
      // Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient("3")
        .then((client: MSGraphClientV3): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody);
        });

    }
  }
  //trannsmittalIDGeneration
  public async _trannsmittalIDGeneration() {
    let prefix;
    let separator; let sequenceNumber;
    let title;
    let counter;
    let transmittalID;
    let transmitTo;
    let select: "TransmittalCategory eq 'Inbound Customer'  and (TransmittalType eq 'Document')";
    await this._Service.getItemWithSelect(this.props.siteUrl, this.props.TransmittalIDSettings, select)
      .then(transmittalIdSettingsItems => {
        console.log("transmittalIdSettingsItems", transmittalIdSettingsItems);
        prefix = transmittalIdSettingsItems[0].Prefix;
        separator = transmittalIdSettingsItems[0].Separator;
        sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
        title = transmittalIdSettingsItems[0].Title;
        counter = transmittalIdSettingsItems[0].Counter;
        let increment = counter + 1;
        var incrementValue = increment.toString();
        this._transmittalSequenceNumber(incrementValue, sequenceNumber);
        transmittalID = prefix + separator + title + separator + this.state.projectNumber + separator + this.state.incrementSequenceNumber;
        console.log("transmittalID", transmittalID);
        this.setState({
          transmittalID: transmittalID,
        });
        //counter updation 
        let updateItem = {
          Counter: increment,
        }
        this._Service.UpdateItemById(this.props.siteUrl, this.props.TransmittalIDSettings, transmittalIdSettingsItems[0].ID, updateItem);
      });
  }
  //transmittalSequenceNumber
  public _transmittalSequenceNumber(incrementvalue, sequenceNumber) {
    var incrementSequenceNumber = incrementvalue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }
  protected async _triggerPermission(SourceDocutID) {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.requestList).items.filter("Title eq 'EMEC_DocumentPermission-Create Document'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': SourceDocutID
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  }
  //submit
  public async submit() {
    if (this.state.ReactTableResult.length == 0) {
      this.setState({ statusMessage: { isShowMessage: true, message: "Please add atleast one project document ", messageType: 1 }, });
      setTimeout(() => {
        this.setState({ statusMessage: { isShowMessage: false, message: "Please add atleast one project document ", messageType: 1 }, });
      }, 2000);
      this.setState({ submitDisable: false });
    }
    else {
      this.setState({ submitDisable: true });
      if (this.transmittalID == null || this.transmittalID == "") {
        let inboundHeader = {
          Title: this.state.transmittalID,
          TransmittalStatus: "Completed",
          Customer: this.state.outlookCustomer,
          CustomerID: parseInt(this.state.outlookCustomerID),
          TransmittalCategory: "Customer",
          TransmittalDate: this.today,
          TransmittedById: this.currentId,
        }
        this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundHeader)
          .then(async inboundTransmittalHeader => {
            this.setState({ inboundTransmittalHeaderId: inboundTransmittalHeader.data.ID });
            this.transmittalID = inboundTransmittalHeader.data.ID;
            let inboundLinks = {
              TransmittalDetails: {
                Description: "Transmittal Details",
                Url: this.props.siteUrl + "/Lists/" + "InboundTransmittalDetails" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
              },
              TransmittalLink: {
                Description: "Project workspace",
                Url: this.props.siteUrl + "/SitePages/" + "InboundTransmittal" + ".aspx?trid=" + inboundTransmittalHeader.data.ID + ""
              },
              InboundAdditionalDetails: {
                Description: "Inbound Additional Details",
                //Url: this.props.siteUrl + "/Lists/" + "InboundAdditionalDocuments" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" +this.transmittalID+ "FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
                Url: this.props.siteUrl + "/" + "InboundAdditionalDocuments" + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
              },
            }
            this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundLinks, inboundTransmittalHeader.data.ID)
            if (this.state.ReactTableResult.length > 0) {
              for (let i = 0; i < this.state.ReactTableResult.length; i++) {
                let inboundDetails = {
                  Title: this.state.ReactTableResult[i].transmittalID,
                  DocumentIndexId: this.state.ReactTableResult[i].docId,
                  Comments: this.state.ReactTableResult[i].comments,
                  OwnerId: this.state.OwnerId,
                  ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                  CustomerDocumentNumber: this.state.ReactTableResult[i].companyDocNumber,
                  PONumber: this.state.ReactTableResult[i].poNumber,
                  TransmittalCodeId: this.state.ReactTableResult[i].transCodeKey,
                  TransmittalHeaderId: this.state.inboundTransmittalHeaderId,
                }
                await this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalDetails, inboundDetails)
                  .then(async iar => {
                    await iar.item.attachmentFiles.add(this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments.name, this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments);
                    let updateIndex = {
                      TransmittalLocation: "IN from Customer",
                      TransmittalStatus: "Received",
                      CurrentStatusCode: this.state.ReactTableResult[i].transmittalCode,
                      CurrentReplyDate: this.state.ReactTableResult[i].receiveDate,
                      CustomerDocumentNo: this.state.ReactTableResult[i].companyDocNumber
                    }
                    await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, updateIndex, parseInt(this.state.ReactTableResult[i].docId))
                      .then(iar => {
                        this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, this.state.ReactTableResult[i].docId)
                          .then(async iar => {
                            console.log(iar);
                            this.setState({ SourceDocumentID: iar.SourceDocumentID });
                            this.setState({ DocumentID: iar.DocumentID });
                            let select = "ID,TransmittalHeader/ID,DocumentIndex/ID,DocumentIndex/Title";
                            let filter = "DocumentIndex/ID eq '" + Number(this.state.ReactTableResult[i].docId) + "'";
                            let expand = "DocumentIndex,TransmittalHeader";
                            await this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.OutboundTransmittalDetails, select, filter, expand)
                              .then(async outboundTransmittalDetailsListNameh => {
                                if (outboundTransmittalDetailsListNameh.length > 0) {
                                  for (var k = 0; k < outboundTransmittalDetailsListNameh.length; k++) {
                                    this.setState({ TransmittalHeaderId: outboundTransmittalDetailsListNameh[k].TransmittalHeader.ID });
                                    let ouboundUpdate = {
                                      TransmittalStatus: "Completed",
                                      ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                                      ReceivedComments: this.state.ReactTableResult[i].comments,
                                    }
                                    this._Service.updateItem(this.props.siteUrl, this.props.OutboundTransmittalDetails, ouboundUpdate, parseInt(outboundTransmittalDetailsListNameh[k].ID));
                                  }
                                }
                                let transmittalHistory = {
                                  Title: this.state.DocumentID,
                                  Status: "IN from Customer",
                                  DocumentIndexId: this.state.ReactTableResult[i].docId,
                                  LogDate: this.state.ReactTableResult[i].receiveDate,
                                }
                                await this._Service.createNewItem(this.props.siteUrl, this.props.TransmittalHistory, transmittalHistory);
                                let select = "SourceDocumentID,Owner/ID,Owner/Title,Owner/EMail,DocumentName,Approver/ID,Approver/Title,Approver/EMail,Reviewers/ID,Reviewers/Title,Reviewers/EMail";
                                let expand = "Owner,Approver,Reviewers";
                                this._Service.getItemSelectExpandById(this.props.siteUrl, "DocumentIndex", select, expand, this.state.ReactTableResult[i].docId)
                                  .then(forGettingOwner => {
                                    this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Owner.EMail, "InboundTransmittalFromCustomer", forGettingOwner.Owner.Title, this.state.ReactTableResult[i].transmittalCode);
                                    this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Approver.EMail, "InboundTransmittalFromCustomer", forGettingOwner.Approver.Title, this.state.ReactTableResult[i].transmittalCode);
                                    if (forGettingOwner.Reviewers) {
                                      for (var k in forGettingOwner.Reviewers) {
                                        if (forGettingOwner.Reviewers[k].EMail != forGettingOwner.Owner.EMail || forGettingOwner.Reviewers[k].EMail != forGettingOwner.Approver.EMail) {
                                          this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Reviewers[k].EMail, "InboundTransmittalFromCustomer", forGettingOwner.Reviewers[k].Title, this.state.ReactTableResult[i].transmittalCode);
                                        }
                                      }
                                    }
                                    this._triggerPermission(forGettingOwner.SourceDocumentID);
                                  });
                                let selectHeaderItems = "TransmittalHeader/ID,TransmittalStatus";
                                let filter = "TransmittalHeader/ID eq '" + Number(this.state.TransmittalHeaderId) + "' ";
                                let expanditems = "TransmittalHeader";
                                this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.OutboundTransmittalDetails, select, filter, expand)
                                  .then(outboundTransmittalDetailsListName => {
                                    let length = outboundTransmittalDetailsListName.length;
                                    let recievedlength = 0;
                                    if (outboundTransmittalDetailsListName.length > 0) {
                                      for (var k = 0; k < outboundTransmittalDetailsListName.length; k++) {
                                        if (outboundTransmittalDetailsListName[k].TransmittalStatus == "Completed") {
                                          recievedlength = recievedlength + 1;
                                        }
                                      }
                                    }
                                    if (length == recievedlength) {
                                      let updateTransmittalStatus = {
                                        TransmittalStatus: "Completed",
                                      }
                                      this._Service.updateItem(this.props.siteUrl, this.props.OutboundTransmittalHeader, updateTransmittalStatus, parseInt(this.state.TransmittalHeaderId));
                                    }
                                  });
                              });
                          });
                      });
                  })
              }
              this._LAUrlGettingForDocumentLibraryUpdate();
            }
            if (this.state.ReactTableResult2.length > 0) {
              for (var i = 0; i < this.state.ReactTableResult2.length; i++) {
                await this._Service.uploadDocument(this.props.siteUrl + "/" + this.props.InboundAdditionalDocuments, this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name, this.state.ReactTableResult2[i].Attachments2)
                  .then(async f => {
                    console.log("File Uploaded");
                    await f.file.getItem().then(async item => {
                      await item.update({
                        Title: this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name,
                        Comments: this.state.ReactTableResult2[i].comments2,
                        ReceivedDate: this.state.ReactTableResult2[i].receiveDate2,
                        TransmittalIDId: this.state.inboundTransmittalHeaderId,
                        Customer: this.state.outlookCustomer,
                        CustomerID: this.state.outlookCustomerID,
                      });
                    });
                  });
              }
            }
          });
      }
      else {
        let inboundHeader = {
          Title: this.state.transmittalID,
          TransmittalStatus: "Completed",
          Customer: this.state.outlookCustomer,
          CustomerID: parseInt(this.state.outlookCustomerID),
          TransmittalCategory: "Customer",
          TransmittalDate: this.today,
          TransmittedById: this.currentId,
        }
        this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundHeader, this.transmittalID)
          .then(async inboundTransmittalHeader => {
            this.setState({ inboundTransmittalHeaderId: this.transmittalID });
            this.transmittalID = inboundTransmittalHeader.data.ID;
            let inboundLinks = {
              TransmittalDetails: {
                Description: "Transmittal Details",
                Url: this.props.siteUrl + "/Lists/" + "InboundTransmittalDetails" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
              },
              TransmittalLink: {
                Description: "Project workspace",
                Url: this.props.siteUrl + "/SitePages/" + "InboundTransmittal" + ".aspx?trid=" + inboundTransmittalHeader.data.ID + ""
              },
              InboundAdditionalDetails: {
                Description: "Inbound Additional Details",
                //Url: this.props.siteUrl + "/Lists/" + "InboundAdditionalDocuments" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" +this.transmittalID+ "FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
                Url: this.props.siteUrl + "/" + "InboundAdditionalDocuments" + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
              },
            }
            this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundLinks, this.transmittalID);
            let selectHeaderItems = "ID,Title,Attachments,AttachmentFiles,DocumentIndex/ID,DocumentIndex/Title,DocumentIndex/DocumentName,Comments,Owner/ID,Owner/Title,TransmittalHeader/ID,ReceivedDate,CustomerDocumentNumber,Comments,TransmittalCode/Title,TransmittalCode/ID,PONumber";
            let expandItems = "AttachmentFiles,DocumentIndex,Owner,TransmittalHeader,TransmittalCode";
            let filter = "TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ";
            this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.InboundTransmittalDetails, selectHeaderItems, filter, expandItems)
              .then(inboundTransmittalDetailsListName => {
                if (inboundTransmittalDetailsListName.length > 0) {
                  this.setState({ currentinBoundDetailItem: inboundTransmittalDetailsListName });
                }
              });
            if (this.state.ReactTableResult.length > this.state.currentinBoundDetailItem.length) {
              for (let i = this.state.currentinBoundDetailItem.length; i < this.state.ReactTableResult.length; i++) {
                let inboundDetails = {
                  Title: this.state.ReactTableResult[i].transmittalID,
                  DocumentIndexId: this.state.ReactTableResult[i].docId,
                  Comments: this.state.ReactTableResult[i].comments,
                  OwnerId: this.state.OwnerId,
                  ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                  CustomerDocumentNumber: this.state.ReactTableResult[i].companyDocNumber,
                  PONumber: this.state.ReactTableResult[i].poNumber,
                  TransmittalCodeId: this.state.ReactTableResult[i].transCodeKey,
                  TransmittalHeaderId: this.state.inboundTransmittalHeaderId,
                }
                await this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalDetails, inboundDetails)
                  .then(iar => {
                    iar.item.attachmentFiles.add(this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments.name, this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments);
                  });
              }
            }
            if (this.state.ReactTableResult.length > 0) {
              for (let i = 0; i < this.state.ReactTableResult.length; i++) {
                let updateIndex = {
                  TransmittalLocation: "IN from Customer",
                  TransmittalStatus: "Received",
                  CurrentStatusCode: this.state.ReactTableResult[i].transmittalCode,
                  CurrentReplyDate: this.state.ReactTableResult[i].receiveDate,
                  CustomerDocumentNo: this.state.ReactTableResult[i].companyDocNumber
                }
                await this._Service.updateItem(this.props.siteUrl, this.props.documentIndexList, updateIndex, parseInt(this.state.ReactTableResult[i].docId))
                  .then(iar => {
                    this._Service.getItemById(this.props.siteUrl, this.props.documentIndexList, this.state.ReactTableResult[i].docId)
                      .then(async di => {
                        this.setState({ SourceDocumentID: di.SourceDocumentID });
                        this.setState({ DocumentID: di.DocumentID });
                        let select = "ID,TransmittalHeader/ID,DocumentIndex/ID,DocumentIndex/Title";
                        let expand = "DocumentIndex,TransmittalHeader";
                        let filter = "DocumentIndex/ID eq '" + Number(this.state.ReactTableResult[i].docId) + "'";
                        await this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.OutboundTransmittalDetails, select, filter, expand)
                          .then(async outboundTransmittalDetailsListNameh => {
                            if (outboundTransmittalDetailsListNameh.length > 0) {
                              for (var k = 0; k < outboundTransmittalDetailsListNameh.length; k++) {
                                this.setState({ TransmittalHeaderId: outboundTransmittalDetailsListNameh[k].TransmittalHeader.ID });
                                let obDetailItems = {
                                  TransmittalStatus: "Completed",
                                  ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                                  ReceivedComments: this.state.ReactTableResult[i].comments,
                                }
                                this._Service.updateItem(this.props.siteUrl, this.props.OutboundTransmittalDetails, obDetailItems, parseInt(outboundTransmittalDetailsListNameh[k].ID))
                              }
                            }
                            let updateIndex = {
                              Title: this.state.DocumentID,
                              Status: "IN from Customer",
                              DocumentIndexId: this.state.ReactTableResult[i].docId,
                              LogDate: this.state.ReactTableResult[i].receiveDate,
                            }
                            await this._Service.createNewItem(this.props.siteUrl, this.props.documentIndexList, updateIndex);
                            let select = "SourceDocumentID,Owner/ID,Owner/Title,Owner/EMail,DocumentName,Approver/ID,Approver/Title,Approver/EMail,Reviewers/ID,Reviewers/Title,Reviewers/EMail";
                            let expands = "Owner,Approver,Reviewers";
                            this._Service.getItemSelectExpandById(this.props.siteUrl, "DocumentIndex", select, expands, this.state.ReactTableResult[i].docId)
                              .then(forGettingOwner => {
                                this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Owner.EMail, "InboundTransmittalFromCustomer", forGettingOwner.Owner.Title, this.state.ReactTableResult[i].transmittalCode);
                                this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Approver.EMail, "InboundTransmittalFromCustomer", forGettingOwner.Approver.Title, this.state.ReactTableResult[i].transmittalCode);
                                if (forGettingOwner.Reviewers) {
                                  for (var k in forGettingOwner.Reviewers) {
                                    if (forGettingOwner.Reviewers[k].EMail != forGettingOwner.Owner.EMail || forGettingOwner.Reviewers[k].EMail != forGettingOwner.Approver.EMail) {
                                      this._sendmail(forGettingOwner.DocumentName, forGettingOwner.Reviewers[k].EMail, "InboundTransmittalFromCustomer", forGettingOwner.Reviewers[k].Title, this.state.ReactTableResult[i].transmittalCode);
                                    }
                                  }
                                }
                                this._triggerPermission(forGettingOwner.SourceDocumentID);
                              });
                            //new code                            
                            let selectHeaderItems = "TransmittalHeader/ID,TransmittalStatus";
                            let filter = "TransmittalHeader/ID eq '" + Number(this.state.TransmittalHeaderId) + "' ";
                            let expand = "TransmittalHeader";
                            this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.OutboundTransmittalDetails, selectHeaderItems, filter, expand)
                              .then(outboundTransmittalDetailsListName => {
                                let length = outboundTransmittalDetailsListName.length;
                                let recievedlength = 0;
                                if (outboundTransmittalDetailsListName.length > 0) {
                                  for (var k = 0; k < outboundTransmittalDetailsListName.length; k++) {
                                    if (outboundTransmittalDetailsListName[k].TransmittalStatus == "Completed") {
                                      recievedlength = recievedlength + 1;
                                    }
                                  }
                                }
                                if (length == recievedlength) {
                                  let updateTS = {
                                    TransmittalStatus: "Completed",
                                  }
                                  this._Service.updateItem(this.props.siteUrl, this.props.OutboundTransmittalHeader, updateTS, parseInt(this.state.TransmittalHeaderId));
                                }
                              });      //last
                          });
                      });
                  });
              }
              await this._LAUrlGettingForDocumentLibraryUpdate()
            }
            await this._Service.getItemWithFilterDL(this.props.siteUrl, this.props.InboundAdditionalDocuments, "TransmittalIDId eq '" + this.transmittalID + "' ")
              .then(inboundAdditionalDocumentsListName => {
                this.setState({ currentInboundAdditionalItem: inboundAdditionalDocumentsListName });
              });
            let i;
            if (this.state.ReactTableResult2.length > this.state.currentInboundAdditionalItem.length) {
              for (i = this.state.currentInboundAdditionalItem.length; i < this.state.ReactTableResult2.length; i++) {
                await this._Service.uploadDocument(this.props.siteUrl + "/" + this.props.InboundAdditionalDocuments, this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name, this.state.ReactTableResult2[i].Attachments2)
                  .then(async f => {
                    console.log("File Uploaded");
                    await f.file.getItem().then(async item => {
                      await item.update({
                        Title: this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name,
                        Comments: this.state.ReactTableResult2[i].comments2,
                        ReceivedDate: this.state.ReactTableResult2[i].receiveDate2,
                        TransmittalIDId: this.state.inboundTransmittalHeaderId,
                        Customer: this.state.outlookCustomer,
                        CustomerID: this.state.outlookCustomerID,
                      });
                    });
                  });
              }
            }
          });
      }
      this.setState({
        btnsvisible: true,
        statusMessage: { isShowMessage: true, message: "Transmittal sent successfully", messageType: 4 },
      });
      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl)
      }, 10000);
    }
  }
  //query param getting
  private _queryParamGetting() {
    let params = new URLSearchParams(window.location.search);
    this.transmittalID = params.get('trid');
    if (this.transmittalID != "" && this.transmittalID != null) {
      this.bindInboundTransmittalSavedData(this.transmittalID);
      this.setState({ queryParamYes: true, queryParamNo: false });
    }
    else {
      this.setState({ queryParamYes: false, queryParamNo: true });
      this._trannsmittalIDGeneration();
      this._projectInformation();
    }
  }
  //binding saved data
  private bindInboundTransmittalSavedData(transmittalID) {
    try {
      this._Service.getItemWithFilter(this.props.siteUrl, this.props.InboundTransmittalHeader, "ID eq '" + this.transmittalID + "' ")
        .then(inboundTransmittalHeader => {
          //  alert(inboundTransmittalHeader[0].Title);
          this.setState({
            outlookCustomer: inboundTransmittalHeader[0].Customer,
            transmittalStatus: inboundTransmittalHeader[0].TransmittalStatus,
            transmittalID: inboundTransmittalHeader[0].Title,
          });
          if (this.state.transmittalStatus == "Completed") {
            this.setState({ projectdivVisible: true, addDocsVisible: true, btnsvisible: true });
          }
          else {
            this.setState({ projectdivVisible: false, btnsvisible: false, addDocsVisible: false });
          }
        });
      let tableData = this.state.ReactTableResult;
      let _docurl = this.props.siteUrl + "/Lists/" + this.props.InboundTransmittalDetails + "/";
      let data: projectData;
      let selectHeaderItems = "ID,Title,Attachments,AttachmentFiles,DocumentIndex/ID,DocumentIndex/Title,DocumentIndex/DocumentName,Comments,Owner/ID,Owner/Title,TransmittalHeader/ID,ReceivedDate,CustomerDocumentNumber,Comments,TransmittalCode/Title,TransmittalCode/ID,PONumber";
      let expandItems = "AttachmentFiles,DocumentIndex,Owner,TransmittalHeader,TransmittalCode";
      let filter = "TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ";
      this._Service.getItemWithFilterExpand(this.props.siteUrl, this.props.InboundTransmittalDetails, selectHeaderItems, expandItems, filter)
        .then(inboundTransmittalDetailsListName => {
          if (inboundTransmittalDetailsListName.length > 0) {
            for (var k = 0; k < inboundTransmittalDetailsListName.length; k++) {
              //this.setState({ transmittalIDedit: inboundTransmittalDetailsListName[k].Title });
              data = {
                OwnerTitle: inboundTransmittalDetailsListName[k].Owner['Title'],
                OwnerId: inboundTransmittalDetailsListName[k].Owner['ID'],
                documentIndex: inboundTransmittalDetailsListName[k].DocumentIndex['DocumentName'],
                companyDocNumber: inboundTransmittalDetailsListName[k].CustomerDocumentNumber,
                receivedDate: inboundTransmittalDetailsListName[k].ReceivedDate == null ? null : moment(inboundTransmittalDetailsListName[k].ReceivedDate).format("DD/MM/YYYY"),
                receiveDate: inboundTransmittalDetailsListName[k].ReceivedDate,
                transmittalCode: inboundTransmittalDetailsListName[k].TransmittalCode['Title'] == null ? null : inboundTransmittalDetailsListName[k].TransmittalCode['Title'],
                poNumber: inboundTransmittalDetailsListName[k].PONumber,
                comments: inboundTransmittalDetailsListName[k].Comments,
                transmittalID: inboundTransmittalDetailsListName[k].Title,
                docId: inboundTransmittalDetailsListName[k].DocumentIndex['ID'],
                transCodeKey: inboundTransmittalDetailsListName[k].TransmittalCode['ID'] == null ? null : inboundTransmittalDetailsListName[k].TransmittalCode['ID'],
                docKey: inboundTransmittalDetailsListName[k].DocumentIndex['Title'],
                Attachments: inboundTransmittalDetailsListName[k].AttachmentFiles[0] == null ? null : inboundTransmittalDetailsListName[k].AttachmentFiles[0],
                ss: inboundTransmittalDetailsListName[k].AttachmentFiles[0] == null ? null : inboundTransmittalDetailsListName[k].AttachmentFiles[0].FileName,
                url: inboundTransmittalDetailsListName[k].AttachmentFiles[0] == null ? null : inboundTransmittalDetailsListName[k].AttachmentFiles[0].ServerRelativeUrl,
                DetailId: inboundTransmittalDetailsListName[k].ID,
              }
              tableData.push(data);
            }
            this.setState({
              ReactTableResult: tableData,
              currentinBoundDetailItem: inboundTransmittalDetailsListName,
              AddIndex: false
            });
          }
        });
      let tableData2 = this.state.ReactTableResult2;
      let data2: ITableData2;
      const h = this._Service.getListItems(this.props.siteUrl, this.props.InboundAdditionalDocuments);
      console.log(h);
      this._Service.getItemWithFilterDL(this.props.siteUrl, this.props.InboundAdditionalDocuments, "TransmittalIDId eq '" + this.transmittalID + "' ")
        .then(inboundAdditionalDocumentsListName => {
          console.log(inboundAdditionalDocumentsListName);
          if (inboundAdditionalDocumentsListName.length > 0) {
            for (var k = 0; k < inboundAdditionalDocumentsListName.length; k++) {
              data2 = {
                Attachments2: null,
                documentName2: inboundAdditionalDocumentsListName[k].Title,
                adddocurl: inboundAdditionalDocumentsListName[k].ServerRedirectedEmbedUrl,
                receivedDate2: inboundAdditionalDocumentsListName[k].ReceivedDate == null ? null : moment(inboundAdditionalDocumentsListName[k].ReceivedDate).format("DD/MM/YYYY"),
                comments2: inboundAdditionalDocumentsListName[k].Comments,
                additionalId: inboundAdditionalDocumentsListName[k].Id,
                receiveDate2: inboundAdditionalDocumentsListName[k].ReceivedDate,
              }
              tableData2.push(data2);
            }
            this.setState({ ReactTableResult2: tableData2, currentInboundAdditionalItem: inboundAdditionalDocumentsListName });
            this.setState({ AddIndex2: false });
          }
        });
      this.setState({ transIdvisible: false });
    }
    catch (ex) {
      console.log("bindInboundTransmittalSavedData Error: " + ex);
    }
  }

  private _confirmYesCancel = () => {
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
  }

  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',

  };
  private dialogDeleteProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to delete?',
  };
  private modalProps = {
    isBlocking: true,
  };
  private _openDeleteConfirmation = (items, key, type) => {
    if (this.transmittalID == "" && this.transmittalID == null) {
      this.setState({
        deleteConfirmMsg: "",
        confirmDialog: false,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      if (type == "ProjectDocuments") {
        this.typeForDelete = "ProjectDocuments";
        this.keyfordelete = key;
      }
      else if (type == "AdditionalDocuments") {
        this.typeForDelete = "AdditionalDocuments";
        this.keyfordelete = key;
      }
    }
    else {
      this.setState({
        deleteConfirmMsg: "",
        confirmDialog: false,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      if (type == "ProjectDocuments") {
        this.typeForDelete = "ProjectDocuments";
        this.keyfordelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.DetailId,
        });
      } else if (type == "AdditionalDocuments") {
        this.typeForDelete = "AdditionalDocuments";
        this.keyfordelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.additionalId,
        });
      }
    }
  }
  public itemDeleteFromExternalGrid(items, key) {
    this.state.ReactTableResult2.splice(key, 1);
    console.log("after removal", this.state.ReactTableResult2);
    this.setState({
      ReactTableResult2: this.state.ReactTableResult2,
    });
  }
  private _confirmDeleteItem = async (docID, items, key) => {
    if (this.transmittalID == "" || this.transmittalID == null) {
      this.setState({
        cancelConfirmMsg: "none",
        confirmDialog: true,
      });
      this.validator.hideMessages();
      if (this.typeForDelete == "ProjectDocuments") {
        this.itemDeleteFromGrid(items, key);
      }
      else if (this.typeForDelete == "AdditionalDocuments") {
        this.itemDeleteFromExternalGrid(items, key);
      }
    }
    else {
      this.setState({
        cancelConfirmMsg: "none",
        confirmDialog: true,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      if (this.typeForDelete == "ProjectDocuments") {
        this.itemDeleteFromGrid(items, key);
        this._Service.deleteItem(this.props.siteUrl, this.props.InboundTransmittalDetails, parseInt(docID));
        this.setState({
          ReactTableResult: this.state.ReactTableResult,
        });
      }
      else if (this.typeForDelete == "AdditionalDocuments") {
        this.itemDeleteFromExternalGrid(items, key);
        this._Service.updateItemDL(this.props.siteUrl, this.props.InboundAdditionalDocuments + "/", parseInt(docID));
        this.setState({
          ReactTableResult2: this.state.ReactTableResult2,
        });
      }
    }
  }
  //Itemdeletefromgrid
  public itemDeleteFromGrid(items, key) {
    console.log(items);
    this.state.ReactTableResult.splice(key, 1);
    console.log("after removal", this.state.ReactTableResult);
    this.setState({ ReactTableResult: this.state.ReactTableResult, confirmDialog: true, });
    if (this.state.ReactTableResult.length == 0) {
      this.setState({ AddIndex: true });
    }
    this.setState({
      ReactTableResult: this.state.ReactTableResult
    });
  }
  //save as draft
  public async saveAsDraft() {

    if (this.state.ReactTableResult.length == 0) {
      this.setState({ submitDisable: false });
      this.setState({ statusMessage: { isShowMessage: true, message: "Please add atleast one project document ", messageType: 1 }, });
      setTimeout(() => {
        this.setState({ statusMessage: { isShowMessage: true, message: "Please add atleast one project document ", messageType: 1 }, });
      }, 2000);
    }
    else {
      this.setState({ submitDisable: true });
      if (this.transmittalID == null || this.transmittalID == "") {
        let inboundHeader = {
          Title: this.state.transmittalID,
          TransmittalStatus: "Draft",
          Customer: this.state.outlookCustomer,
          CustomerID: parseInt(this.state.outlookCustomerID),
          TransmittalCategory: "Customer",
          //TransmittalDate: this.state.todayDate,
          TransmittedById: this.currentId,
        }
        const inboundTransmittalHeader = await this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundHeader);
        if (inboundTransmittalHeader) {
          this.setState({ inboundTransmittalHeaderId: inboundTransmittalHeader.data.ID });
          let inboundLinks = {
            TransmittalDetails: {
              Description: "Transmittal Details",
              Url: this.props.siteUrl + "/Lists/" + "InboundTransmittalDetails" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
            },
            TransmittalLink: {
              Description: "Project workspace",
              Url: this.props.siteUrl + "/SitePages/" + "InboundTransmittal" + ".aspx?trid=" + inboundTransmittalHeader.data.ID + ""
            },
            InboundAdditionalDetails: {
              Description: "Inbound Additional Details",
              //Url: this.props.siteUrl + "/Lists/" + "InboundAdditionalDocuments" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" +this.transmittalID+ "FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
              Url: this.props.siteUrl + "/" + "InboundAdditionalDocuments" + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
            },
          }
          const updateTransmittalHeader = await this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundLinks, inboundTransmittalHeader.data.ID);
          if (updateTransmittalHeader) {
            if (this.state.ReactTableResult.length > 0) {
              for (var i in this.state.ReactTableResult) {
                try {
                  let inboundDetails = {
                    Title: this.state.ReactTableResult[i].transmittalID,
                    DocumentIndexId: this.state.ReactTableResult[i].docId,
                    Comments: this.state.ReactTableResult[i].comments,
                    OwnerId: this.state.ReactTableResult[i].OwnerId,
                    ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                    TransmittalCodeId: this.state.ReactTableResult[i].transCodeKey,
                    CustomerDocumentNumber: this.state.ReactTableResult[i].companyDocNumber,
                    PONumber: this.state.ReactTableResult[i].poNumber,
                    TransmittalHeaderId: this.state.inboundTransmittalHeaderId,
                  }
                  const inboundTransmittalDetails = await this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalDetails, inboundDetails)
                  if (inboundTransmittalDetails) {
                    inboundTransmittalDetails.item.attachmentFiles.add(this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments.name, this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments);
                  }
                }
                catch (ex) {
                  console.log("Approved Error: " + ex);
                }
              }
            }
            if (this.state.ReactTableResult2.length > 0) {
              for (var i in this.state.ReactTableResult2) {
                const inboundadditionaldocuments = await this._Service.uploadDocument(this.props.siteUrl + "/" + this.props.InboundAdditionalDocuments, this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name, this.state.ReactTableResult2[i].Attachments2)
                if (inboundadditionaldocuments) {
                  const item = await inboundadditionaldocuments.file.getItem();
                  if (item) {
                    const update = await item.update({
                      Title: this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name,
                      Comments: this.state.ReactTableResult2[i].comments2,
                      ReceivedDate: this.state.ReactTableResult2[i].receiveDate2,
                      TransmittalIDId: this.state.inboundTransmittalHeaderId,
                      Customer: this.state.outlookCustomer,
                      CustomerID: parseInt(this.state.outlookCustomerID),
                    });
                  }
                }
              }
            }
          }
        }
      }
      else {
        let inboundHeader = {
          Title: this.state.transmittalID,
          TransmittalStatus: "Draft",
          Customer: this.state.outlookCustomer,
          CustomerID: parseInt(this.state.outlookCustomerID),
          TransmittalCategory: "Customer",
          // TransmittalDate: this.state.todayDate,
          TransmittedById: this.currentId,
        }
        const inboundTransmittalHeader = await this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundHeader, this.transmittalID)

        if (inboundTransmittalHeader) {
          this.setState({ inboundTransmittalHeaderId: this.transmittalID });
          let inboundLinks = {
            TransmittalDetails: {
              Description: "Transmittal Details",
              Url: this.props.siteUrl + "/Lists/" + "InboundTransmittalDetails" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
            },
            TransmittalLink: {
              Description: "Project workspace",
              Url: this.props.siteUrl + "/SitePages/" + "InboundTransmittal" + ".aspx?trid=" + inboundTransmittalHeader.data.ID + ""
            },
            InboundAdditionalDetails: {
              Description: "Inbound Additional Details",
              //Url: this.props.siteUrl + "/Lists/" + "InboundAdditionalDocuments" + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" +this.transmittalID+ "FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
              Url: this.props.siteUrl + "/" + "InboundAdditionalDocuments" + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=d22d3ef1%2Dca95%2D4a3c%2Da124%2Dbeb594f07906"
            },
          }
          const updateTransmittalHeader = await this._Service.updateItem(this.props.siteUrl, this.props.InboundTransmittalHeader, inboundLinks, inboundTransmittalHeader.data.ID);

          if (updateTransmittalHeader) {
            if (this.state.ReactTableResult.length > 0) {
              for (var i in this.state.ReactTableResult) {
                try {
                  let inboundDetails = {
                    Title: this.state.ReactTableResult[i].transmittalID,
                    DocumentIndexId: this.state.ReactTableResult[i].docId,
                    Comments: this.state.ReactTableResult[i].comments,
                    OwnerId: this.state.ReactTableResult[i].OwnerId,
                    ReceivedDate: this.state.ReactTableResult[i].receiveDate,
                    TransmittalCodeId: this.state.ReactTableResult[i].transCodeKey,
                    CustomerDocumentNumber: this.state.ReactTableResult[i].companyDocNumber,
                    PONumber: this.state.ReactTableResult[i].poNumber,
                    TransmittalHeaderId: this.state.inboundTransmittalHeaderId,
                  }
                  const inboundTransmittalDetails = await this._Service.createNewItem(this.props.siteUrl, this.props.InboundTransmittalDetails, inboundDetails)
                  if (inboundTransmittalDetails) {
                    inboundTransmittalDetails.item.attachmentFiles.add(this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments.name, this.state.ReactTableResult[i].Attachments == null ? null : this.state.ReactTableResult[i].Attachments);
                  }
                }
                catch (ex) {
                  console.log("Approved Error: " + ex);
                }
              }
            }
            if (this.state.ReactTableResult2.length > 0) {
              for (var i in this.state.ReactTableResult2) {
                const inboundadditionaldocuments = await this._Service.uploadDocument(this.props.siteUrl + "/" + this.props.InboundAdditionalDocuments, this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name, this.state.ReactTableResult2[i].Attachments2)
                if (inboundadditionaldocuments) {
                  const item = await inboundadditionaldocuments.file.getItem();
                  if (item) {
                    const update = await item.update({
                      Title: this.state.transmittalID + this.state.ReactTableResult2[i].Attachments2.name,
                      Comments: this.state.ReactTableResult2[i].comments2,
                      ReceivedDate: this.state.ReactTableResult2[i].receiveDate2,
                      TransmittalIDId: this.state.inboundTransmittalHeaderId,
                      Customer: this.state.outlookCustomer,
                      CustomerID: parseInt(this.state.outlookCustomerID),
                    });
                  }
                }
              }
            }
          }
        }
      }
      this.setState({
        btnsvisible: true,
        statusMessage: { isShowMessage: true, message: "Transmittal Saved as Draft", messageType: 4 },
      });
      setTimeout(() => {
        window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl)

      }, 6000);
    }
  }
  //getting project information
  public _projectInformation = async () => {
    const projectInformation = await this._Service.getListItems(this.props.siteUrl, this.props.projectInformationListName);
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
        if (projectInformation[k].Key == "ProjectName") {
          this.setState({ projectName: projectInformation[k].Title, });
        }
        if (projectInformation[k].Key == "ProjectNumber") {
          this.setState({ projectNumber: projectInformation[k].Title, });
        }
        if (projectInformation[k].Key == "Customer") {
          this.setState({ outlookCustomer: projectInformation[k].Title, });
        }
        if (projectInformation[k].Key == "CustomerID") {
          this.setState({ outlookCustomerID: projectInformation[k].Title, });
        }
        if (projectInformation[k].Key == "ContractNumber") {
          this.setState({ outlookPONumber: projectInformation[k].Title, });
        }
        // ContractNumber
      }
    }
  }
  //Binding data on page load
  public async _bindData() {
    this.setState({ todayDate: moment(this.today).format('DD/MM/YYYY') });
    this._projectInformation();
    let documentIndexArray = [];
    let outlookLibraryArray = [];
    let sorted_documentIndexArray: any[];
    let TransmittalCodeSettingsArray = [];
    let departmentArray = [];
    let sorted_Department: any[];
    let categoryArray = [];
    let sorted_Category: any[];
    //Get Document Index
    let select = "ID,DocumentName,TransmittalStatus";
    let filter = "TransmittalStatus eq 'Ongoing' and (WorkflowStatus eq 'Published')";
    const documentIndex: any[] = await this._Service.getItemWithSelectFilter(this.props.siteUrl, this.props.documentIndexList, select, filter);
    for (let i = 0; i < documentIndex.length; i++) {
      let documentIndexdata = {
        key: documentIndex[i].ID,
        text: documentIndex[i].DocumentName,
      }
      documentIndexArray.push(documentIndexdata);
      this.setState({
        documentIndexOption: documentIndexArray,
      });
    }
    const TransmittalCodeSettings: any[] = await this._Service.getListItems(this.props.siteUrl, this.props.TransmittalCodeSettings);
    for (let i = 0; i < TransmittalCodeSettings.length; i++) {
      if (TransmittalCodeSettings[i].AcceptanceCode == true) {
        let TransmittalCodeSettingsoptions = {
          key: TransmittalCodeSettings[i].ID,
          text: TransmittalCodeSettings[i].Title
        };
        TransmittalCodeSettingsArray.push(TransmittalCodeSettingsoptions);
      }
      this.setState({ TransmittalCodeSettings: TransmittalCodeSettingsArray, });
    }
  }
  //Rendering Controls
  public render(): React.ReactElement<IInboundCustomerProps> {
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const AddIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
    return (
      <div>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div className={styles.InboundCustomerWp} style={{ display: this.state.webpartView }}>
          <div style={{ marginLeft: "", marginRight: "", width: "50rem" }}>
            <div style={{ fontWeight: "bold", fontSize: "15px", textAlign: "center" }}> Inbound Transmittal from {this.state.outlookCustomer}</div>
            <div hidden={this.state.transIdvisible} >
              <div style={{ display: "flex", margin: "7px" }} hidden={this.state.transIdvisible}>
                <Label hidden={this.state.transIdvisible}>Transmittal ID : {this.state.transmittalID}</Label>
              </div>
            </div>
            <div className={styles.row}>
              <div style={{ display: "flex", margin: "7px" }}>
                <Label>Transmittal Date : {this.state.todayDate}</Label>
                <Label style={{ padding: "0 0 0 185px" }}>Project : {this.state.projectNumber}-{this.state.projectName}</Label>
              </div>
              <div className={styles.border}>
                <Label>Project Documents</Label>
                <div style={{ display: "flex", margin: "7px", width: "100%" }}>
                  <div style={{ width: "100%" }} >
                    <Dropdown
                      style={{ width: "100%" }}
                      placeholder="Select Document Index"
                      label="Document Index"
                      options={this.state.documentIndexOption}
                      onChanged={this.DocIndex}
                      selectedKey={this.state.docId}
                    />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("docId", this.state.docId, "required")}{" "}
                    </div>
                  </div>
                </div>


                <div style={{ display: "flex", margin: "7px", width: "100%" }}>
                  <div style={{ width: '48%', marginRight: '4%' }}>
                    <DatePicker label="Received Date"
                      style={{ width: '100%', }}
                      value={this.state.receivedDate}
                      onSelectDate={this._onreceivedDateChange}
                      placeholder="Select a date..."
                      formatDate={this._onFormatDate}
                    />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("received date", this.state.receivedDate, "required")}{" "}</div>
                  </div>
                  <div style={{ width: '48%' }}>
                    <Dropdown
                      style={{ width: '100%' }}
                      placeholder="Select Acceptance Code"
                      label="Acceptance Code"
                      options={this.state.TransmittalCodeSettings}
                      onChanged={this.transcodechange}
                      selectedKey={this.state.transCodeKey}
                    />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Acceptance Code", this.state.transCodeKey, "required")}{" "}</div>
                  </div>
                </div>

                <div style={{ display: "flex", margin: "7px" }}>
                  <Label > Browse Comments    </Label>
                  <input type="file" style={{ margin: "7px", paddingLeft: "7px" }} className="custom-file-input" id="studyAttachment" onChange={this.handleFileUpload} />
                </div>
                <div style={{ display: "flex", width: "100%", margin: "7px" }}>
                  <div style={{ width: '50%', marginRight: '4%' }}>
                    <TextField label="Customer Contract Number" id="ponumber" name="ponumber" value={this.state.outlookPONumber} onChange={this._PoNumbChange}
                      style={{ width: '100%' }} readOnly={true}></TextField>
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("ponumber", this.state.outlookPONumber, "required")}{" "}</div>
                  </div>



                  <div style={{ width: '50%' }}>
                    <TextField id="CompanyDocNo" name="CompanyDocNo" label="Customer Doc No"
                      onChange={this._CompanyDocChange}
                      value={this.state.outlookCustomerDocNo} style={{ width: '100%' }} ></TextField>

                  </div>
                </div>


                <div style={{ display: "flex", margin: "7px", width: "100%" }}>
                  <div style={{ width: "90%" }}>
                    <TextField label="Comments" multiline autoAdjustHeight id="comments" onChange={this._titleChange} value={this.state.comments} />
                    <div style={{ color: "#dc3545" }}>
                      {this.validator.message("Comments", this.state.comments, "required")}{" "}</div>
                  </div>
                  <div>
                    <IconButton iconProps={AddIcon} title="Addindex" ariaLabel="Addindex" onClick={this.Addindex} style={{ padding: "58px 0px 0px 45px" }} />
                  </div>

                </div>
                <div hidden={this.state.docselected}> <div style={{ color: "#dc3545" }}><Label style={{ color: "#dc3545" }}>Document already selected.  </Label></div></div>
                <div hidden={this.state.AddIndex}>
                  <table style={{ border: '1px solid #ddd', width: '100%', overflowX: 'scroll', borderCollapse: 'collapse', marginLeft: "6px" }} >
                    <tr>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Document</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Owner</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Received Date</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Comments</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Transmittal Code</th>
                      <th hidden={this.state.queryParamYes} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Attachment</th>
                      <th hidden={this.state.queryParamNo} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Attachment URL</th>
                      <th hidden={this.state.projectdivVisible} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Delete</th>
                    </tr>
                    <tbody>
                      {this.state.ReactTableResult.map((item, key) => {
                        return (
                          <tr key={key}>

                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', height: 'auto', width: '70px', wordWrap: 'break-word' }} >{item.documentIndex}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}>{item.OwnerTitle}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}> {item.receivedDate}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', height: 'auto', width: '70px', wordWrap: 'break-word' }} >{item.comments}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}>{item.transmittalCode}</td>
                            <td hidden={this.state.queryParamYes} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}>
                              {item.Attachments == null ? null : item.Attachments.name}
                            </td>
                            <td hidden={this.state.queryParamNo} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}>
                              <a href={item.url} target="_blank">  {item.ss == null ? null : item.ss}</a>
                            </td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse' }}>
                              <IconButton iconProps={DeleteIcon} title="Delete" onClick={() => this._openDeleteConfirmation(item, key, "ProjectDocuments")} ariaLabel="Delete" /></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <div >
                  <hr style={{ marginTop: "20px" }} />
                  <Label>Additional Documents</Label>
                  <div hidden={this.state.docaddselected}> <div style={{ color: "#dc3545" }}><Label style={{ color: "#dc3545" }}>Document already selected.  </Label></div></div>
                  <div style={{ width: "100%", display: "flex", margin: "7px" }}>
                    <div style={{ width: "50%" }}>
                      <Label >Upload Document:</Label>
                      <input type="file" style={{ width: '98%' }} id="additionalfile" ref={this.fileInput} onChange={this.handleFileUpload2}

                      ></input>
                    </div>
                    <div style={{ width: "50%" }}>
                      <DatePicker label="Received Date"
                        style={{ width: '98%' }}
                        value={this.state.receivedDate2}
                        onSelectDate={this._onreceivedDateChange2}
                        placeholder="Select a date..."
                        formatDate={this._onFormatDate}
                      />
                      <div style={{ color: "#dc3545" }}>
                        {this.validator.message("receivedDate2", this.state.receivedDate2, "required")}{" "}</div>
                    </div>               </div>




                  <div className='addcomments' style={{ display: "flex", margin: "7px", width: "100%" }}>
                    <div style={{ width: "90%" }}>
                      <TextField label="Comments" style={{ width: "90%" }} multiline autoAdjustHeight id="comments2" onChange={this._titleChange2} value={this.state.comments2} />
                    </div><div>
                      <IconButton iconProps={AddIcon} title="AddDoc2" ariaLabel="AddDoc2" onClick={this.Addindex2} style={{ padding: "58px 0px 0px 45px" }} />
                    </div>
                  </div>
                </div>
                <div hidden={this.state.AddIndex2}>
                  {<table style={{ border: '1px solid #ddd', width: '100%', borderCollapse: 'collapse', marginLeft: "6px" }} hidden={this.state.AddIndex2} >
                    <tr>
                      <th hidden={this.state.queryParamYes} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Document </th>
                      <th hidden={this.state.queryParamNo} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Document </th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Received Date</th>
                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Comments</th>

                      <th style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', backgroundColor: '#eee' }}>Delete</th>
                    </tr>
                    <tbody>
                      {this.state.ReactTableResult2.map((item, key) => {
                        return (
                          <tr key={key}>
                            <td hidden={this.state.queryParamYes} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', height: 'auto', width: '170px', wordWrap: 'break-word' }}>
                              {item.documentName2}</td>
                            <td hidden={this.state.queryParamNo} style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', height: 'auto', width: '170px', wordWrap: 'break-word' }}>
                              <a target="_blank" href={item.adddocurl}>{item.documentName2}</a>
                            </td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', width: '70px' }}>{item.receivedDate2}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', height: 'auto', width: '100px', wordWrap: 'break-word' }}>{item.comments2}</td>
                            <td style={{ border: '1px solid #ddd', padding: '8px', borderCollapse: 'collapse', width: '10px' }}>
                              <IconButton iconProps={DeleteIcon} title="Delete" onClick={() => this._openDeleteConfirmation(item, key, "AdditionalDocuments")} /></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>}
                </div>
              </div>
              <div>
                {this.state.statusMessage.isShowMessage ?
                  <MessageBar
                    messageBarType={this.state.statusMessage.messageType}
                    isMultiline={false}
                    dismissButtonAriaLabel="Close"
                  >{this.state.statusMessage.message}</MessageBar>
                  : ''
                }
              </div>
              <div style={{ display: "flex", padding: "33px 26px 12px 2px", float: "right" }} hidden={this.state.btnsvisible}>
                <div hidden={this.state.btnsvisible}>
                  <PrimaryButton text="Save as draft" style={{ marginLeft: "300px" }} disabled={this.state.submitDisable} onClick={this.saveAsDraft} hidden={this.state.btnsvisible} />
                  <PrimaryButton text="Submit" style={{ marginLeft: "10px" }} onClick={this.submit} disabled={this.state.submitDisable} hidden={this.state.btnsvisible} />
                </div>
                <PrimaryButton text="Cancel" style={{ marginLeft: "10px" }} onClick={this._hideGrid} />
              </div>
            </div>
          </div >
          <div style={{ display: this.state.deleteConfirmMsg }}>
            <div>
              <Dialog
                hidden={this.state.confirmDialog}
                dialogContentProps={this.dialogDeleteProps}
                onDismiss={this._dialogCloseButton}
                styles={this.dialogStyles}
                modalProps={this.modalProps}>
                <DialogFooter>
                  <PrimaryButton onClick={() => this._confirmDeleteItem(this.state.tempDocIndexIDForDelete, "item", this.keyfordelete)} text="Yes" />
                  <DefaultButton onClick={this._confirmNoCancel} text="No" />
                </DialogFooter>
              </Dialog>
            </div>
          </div>
        </div>
        <div style={{ display: this.state.cancelConfirmMsg }}>
          <div>
            <Dialog
              hidden={this.state.confirmCancelDialog}
              dialogContentProps={this.dialogCancelContentProps}
              onDismiss={this._dialogCloseButton}
              styles={this.dialogStyles}
              modalProps={this.modalProps}>
              <DialogFooter>
                <PrimaryButton onClick={() => this._cancelConfirmYes()} text="Yes" />
                <DefaultButton onClick={this._confirmNoCancel} text="No" />
              </DialogFooter>
            </Dialog>
          </div>
        </div>
        {/* <div>
          {this.state.statusMessage.isShowMessage ?
            <MessageBar
              messageBarType={this.state.statusMessage.messageType}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >{this.state.statusMessage.message}</MessageBar>
            : ''
          }
        </div> */}
      </div>
    );
  }
}


