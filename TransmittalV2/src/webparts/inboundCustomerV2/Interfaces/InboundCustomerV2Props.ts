import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface InboundCustomerV2Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  siteUrl:string;
  hubSiteUrl:string;
  hubsite:string;
  projectInformationListName:string;
   TransmittalIDSettings:string;
   InboundTransmittalHeader:string;     
   InboundTransmittalDetails:string;
   OutboundTransmittalHeader:string;
   OutboundTransmittalDetails:string;
   InboundAdditionalDocuments :string;  
   documentIndexList:string;
   TransmittalHistory:string;
   TransmittalOutlookLibrary:string;
   TransmittalCodeSettings:string;
   EmailNotificationSettings:string;
   requestList:string;
   NotificationPreferenceSettings:string;
   redirectUrl:string;    
   PermissionMatrixSettings:string;
   accessGroupDetailsList:string;
}
export interface InboundCustomerV2State {
  recallConfirmMsgDiv: string;
  recallConfirmMsg: boolean;
  confirmDeleteDialog: boolean;
  confirmCancelDialog: boolean;
  docaddselected: boolean;
  access: any;
  docselected: boolean;
  ownerEmail: any;
  queryParamYes: boolean;
  queryParamNo: boolean;
  OwnerId: any;
  OwnerTitle: any;
  statusMessage: IMessage;
  currentinBoundDetailItem: projectData[];
  currentInboundAdditionalItem: ITableData2[];
  projectdivVisible: boolean;
  transIdvisible: boolean;
  inboundTransmittalHeaderId: any;
  Attachments: File;
  Attachments2: File;
  incrementSequenceNumber: any;
  outlookCustomerID: any;
  outlookContractNumber: any;
  outlookCustomer: any;
  outlookPONumber: any;
  outlookCustomerDocNo: any;
  AddIndex: boolean;
  AddIndex2: boolean;
  documentIndexOption: any[];
  TransmittalCodeSettings: any[];
  ReactTableResult: projectData[];
  ReactTableResult2: ITableData2[];
  receivedDate: any;
  receivedDate2: any;
  transmittalCode: string;
  comments: string;
  comments2: string;
  projectName: string;
  projectNumber: string;
  todayDate: any;
  docId: any;
  docKey: any;
  transCodeKey: any;
  transcodeText: any;
  documentIndex: any;
  poNumber: any;
  transmittalID: any;
  transmittalStatus: any;
  btnsvisible: boolean;
  addDocsVisible: boolean;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  tempDocIndexIDForDelete: any;
  deleteConfirmMsg: string;
  statusKey: any;
  SourceDocumentID: any;
  DocumentID: any;
  TransmittalHeaderId: any;
  accessDeniedMsgBar: any;
  loaderDisplay: string;
  webpartView: string;
  submitDisable: any;
  deleteConfirmation: string;
}
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
export interface projectData {
  //slNo: any;
  OwnerTitle: any;
  OwnerId: any;
  documentIndex: any;
  companyDocNumber: any;
  receiveDate: any;
  receivedDate: any;
  transmittalCode: any;
  poNumber: any;
  comments: any;
  Attachments: File;
  transmittalID: any;
  docKey: any;
  transCodeKey: any;
  docId: any;
  ss: any;
  url: any;
  DetailId: any;
}
export interface ITableData2 {
  documentName2: string;
  receivedDate2: any;
  comments2: string;
  additionalId: any;
  adddocurl: any;
  receiveDate2: any;
  Attachments2: File;
}