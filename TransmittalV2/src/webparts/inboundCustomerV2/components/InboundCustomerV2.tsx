import * as React from 'react';
import styles from './InboundCustomerV2.module.scss';
import {InboundCustomerV2Props, InboundCustomerV2State } from '../Interfaces/InboundCustomerV2Props';
import { IBCService } from '../Services/IBCService';
import { Label } from '@fluentui/react';

export default class InboundCustomerV2 extends React.Component<InboundCustomerV2Props,InboundCustomerV2State,{}> {
  fileInput: React.RefObject<unknown>;
  private _Service: IBCService;    
  constructor(props: InboundCustomerV2Props) {
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
    this._Service = new IBCService(this.props.context,window.location.protocol + "//" + window.location.hostname +this.props.hubSiteUrl);
    this._projectInformation = this._projectInformation.bind(this);
  }
  public render(): React.ReactElement<InboundCustomerV2Props> {
    const {      
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.inboundCustomerV2} ${hasTeamsContext ? styles.teams : ''}`}>
       <div style={{ fontWeight: "bold", fontSize: "15px", textAlign: "center" }}> Inbound Transmittal from {this.state.outlookCustomer}</div>
            <div hidden={this.state.transIdvisible} >
              <div style={{ display: "flex", margin: "7px" }} hidden={this.state.transIdvisible}>
                <Label hidden={this.state.transIdvisible}>Transmittal ID : {this.state.transmittalID}</Label>
              </div>
            </div>
            <div className={styles.header}>
            <div className={styles.divMetadataCol1}>
                <h3 >Project Details</h3>
              </div>
             </div>
            <div className={styles.row}>
              <div></div>
              <div>
                <Label>Transmittal Date : {this.state.todayDate}</Label>              
              </div>
              <div>
              <Label >Project : {this.state.projectNumber}-{this.state.projectName}</Label>
              </div>
              </div>
      </section>
    );
  }
  public async componentDidMount() {
   this._projectInformation();
  }
  public _projectInformation = async () => {
  this._Service.getListItems(this.props.siteUrl,this.props.projectInformationListName)
  .then(projectInformation =>{
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
      }
    }
  });
  }
}
