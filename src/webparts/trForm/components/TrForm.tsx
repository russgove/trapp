
/** Fabric */
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {  IStyle,  ITheme,  getTheme,  mergeStyles} from '@uifabric/styling';

/** SPFX Stuff */
import * as React from 'react';
import styles from './TrForm.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

/** Other utilities */
import * as moment from 'moment';
import * as _ from "lodash";
import * as tabs from "react-tabs";

/**  Custom Stuff */
import { DocumentIframe } from "./DocumentIframe";
import { TRDocument, TR, modes, Pigment, Test, PropertyTest, DisplayPropertyTest } from "../dataModel";
import { ITrFormProps } from './ITrFormProps';
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
import TRPicker from "./TRPicker";
import { ITRFormState } from "./ITRFormState";


export default class TrForm extends React.Component<ITrFormProps, ITRFormState> {
  private ckeditor: any;
  private originalAssignees: Array<number> = [];
  private originalStatus: string = "";
  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  constructor(props: ITrFormProps) {
    super(props);
    this.state = props.initialState;
    this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);// sasve original so we can email new assignees
    this.originalStatus = this.state.tr.TRStatus;// sasve original so we can email if it gets closed
    this.SaveButton = this.SaveButton.bind(this);
    this.ModeDisplay = this.ModeDisplay.bind(this);
    this.StatusDisplay = this.StatusDisplay.bind(this);
    this.save = this.save.bind(this);
    this.cancel = this.cancel.bind(this);
    this.rendeChildTRAsLink = this.rendeChildTRAsLink.bind(this);
    this.rendeDocumentAsLink = this.rendeDocumentAsLink.bind(this);
    this.selectChildTR = this.selectChildTR.bind(this);
    this.cancelTrSearch = this.cancelTrSearch.bind(this);
    this.parentTRSelected = this.parentTRSelected.bind(this);
    this.editParentTR = this.editParentTR.bind(this);
    this.uploadFile = this.uploadFile.bind(this);

  }


  public componentDidMount() {

    var ckEditorCdn: string = '//cdn.ckeditor.com/4.6.2/full/ckeditor.js';
    SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: 'CKEDITOR' }).then((CKEDITOR: any): void => {
      this.ckeditor = CKEDITOR;
      this.ckeditor.replace("tronoxtrtextarea-title"); // replaces the title with a ckeditor. the other textareas are not visible yet. They will be replaced when the tab becomes active

    });

  }
  public tabChanged(newTabID, oldTabID) {


    switch (newTabID) {
      case 0:
        if (this.ckeditor.instances["tronoxtrtextarea-title"] === undefined) {
          new Promise(resolve => setTimeout(resolve, 200)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-title");
            console.log("created tronoxtrtextarea-title");
          });
        }
        break;
      case 1:
        if (this.ckeditor.instances["tronoxtrtextarea-description"] === undefined) {
          new Promise(resolve => setTimeout(resolve, 200)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-description");
            console.log("created tronoxtrtextarea-description");
          });
        }
        break;
      case 2:
        if (this.ckeditor.instances["tronoxtrtextarea-summary"] === undefined) {
          new Promise(resolve => setTimeout(resolve, 200)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-summary");
            console.log("created tronoxtrtextarea-summary");
          });
        }
        break;
      case 3:
        if (this.ckeditor.instances["tronoxtrtextarea-testparams"] === undefined) {
          new Promise(resolve => setTimeout(resolve, 200)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-testparams");
            console.log("created tronoxtrtextarea-testparams");
          });
        }
        break;
      case 8:
        if (this.ckeditor.instances["tronoxtrtextarea-formulae"] === undefined) {
          new Promise(resolve => setTimeout(resolve, 200)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-formulae");
            console.log("created tronoxtrtextarea-formulae");
          });
        }
        break;
      default:

    }



  }
  public isValid(): boolean {

    this.state.errorMessages = [];
    let errorsFound = false;
    if (!this.state.tr.Title) {
      this.state.errorMessages.push(new md.Message("Request #  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.WorkTypeId) {
      this.state.errorMessages.push(new md.Message("Work Type is required"));
      errorsFound = true;
    }
    if (!this.state.tr.ApplicationTypeId) {
      this.state.errorMessages.push(new md.Message("Application Type is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequestDate) {
      this.state.errorMessages.push(new md.Message("Initiation Date   is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequiredDate) {
      this.state.errorMessages.push(new md.Message("Due Date  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.Site) {
      this.state.errorMessages.push(new md.Message("Site is required"));
      errorsFound = true;
    }
    if (!this.state.tr.TRPriority) {
      this.state.errorMessages.push(new md.Message("Proiority  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.TRStatus) {
      this.state.errorMessages.push(new md.Message("Status  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequestorId) {
      this.state.errorMessages.push(new md.Message("Requestor is required"));
      errorsFound = true;
    }


    return !errorsFound;
  }
  public save() {

    for (let instanceName in this.ckeditor.instances) {

      let instance = this.ckeditor.instances[instanceName];
      let data = instance.getData();
      switch (instanceName) {
        case "tronoxtrtextarea-title":
          this.state.tr.RequestTitle = data;
          break;
        case "tronoxtrtextarea-description":
          this.state.tr.Description = data;
          break;
        case "tronoxtrtextarea-summary":
          this.state.tr.Summary = data;
          break;
        case "tronoxtrtextarea-testparams":
          this.state.tr.TestingParameters = data;
          break;
        case "tronoxtrtextarea-formulae":
          this.state.tr.Formulae = data;
          break;
        default:
          alert("Text area missing in save");

      }
    }
    if (this.isValid()) {
      this.props.save(this.state.tr, this.originalAssignees, this.originalStatus)
        .then((result: TR) => {
          this.state.tr.Id = result.Id;
          this.state.isDirty = false;
          this.setState(this.state);
        })
        .catch((response) => {
          this.state.errorMessages.push(new md.Message(response.data.responseBody['odata.error'].message.value));
          this.setState(this.state);
        });
    } else {
      this.setState(this.state); // show errors
    }
    return false; // stop postback
  }
  public cancel() {
    this.props.cancel();
    return false; // stop postback
  }
  public updateCKEditorText(tr: TR) { // updates the text in all the existingck editors after we loaded a new TR (parent or child)
    for (let instanceName in this.ckeditor.instances) {
      let instance = this.ckeditor.instances[instanceName];
      switch (instanceName) {
        case "tronoxtrtextarea-title":
          instance.setData(tr.RequestTitle);
          break;
        case "tronoxtrtextarea-description":
          instance.setData(tr.Description);
          break;
        case "tronoxtrtextarea-summary":
          instance.setData(tr.Summary);
          break;
        case "tronoxtrtextarea-formulae":
          instance.setData(tr.Formulae);
          break;
        default:

      }
    }
  }
  public removeMessage(messageList: Array<md.Message>, messageId: string) {
    _.remove(messageList, {
      Id: messageId
    });
    this.setState(this.state);
  }

  public SaveButton(): JSX.Element {
    if (this.props.mode === modes.DISPLAY) {
      return <div />;
    } else return (
      <PrimaryButton  theme={getTheme()} buttonType={ButtonType.primary} onClick={this.save} icon="ms-Icon--Save">
        <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
        Save
      </PrimaryButton>
      /*<span style={{ margin: 20 }}>

        <a href="#" onClick={this.save} style={{ border: 5, backgroundColor: 'lightBlue', fontSize: 'large' }}>
          <i className="ms-Icon ms-Icon--Save"></i>Save
        </a>

      </span>*/
    );
  }
  public ModeDisplay(): JSX.Element {
    return (
      <Label>MODE : {modes[this.props.mode]}</Label>
    );

  }
  public StatusDisplay(): JSX.Element {
    return (
      <Label>Status : {(this.state.isDirty) ? "Unsaved" : "Saved"}</Label>
    );

  }

  /**
   * 
   * 
   * @returns 
   * 
   * @memberof TrForm
   */
  public getTests() {
    var tests = _.map(this.props.tests, (test) => {
      return {
        title: test.title,

        selected: ((this.state.tr.TestsId) ? this.state.tr.TestsId.indexOf(test.id) != -1 : null),
        id: test.id
      };
    });
    return _.orderBy(tests, ["selected", "title"], ["desc", "asc"]);
  }
  public trContainsPigment(tr: TR, PigmentId: number): boolean {
    if (tr.PigmentsId) {
      return (tr.PigmentsId.indexOf(PigmentId) != -1);
    }
    else return false;
  }
  public trContainsTest(tr: TR, TestId: number): boolean {
    if (tr.TestsId) {
      return (tr.TestsId.indexOf(TestId) != -1);
    }
    else {
      return false;
    }
  }
  /**
   * return Pigments on the tr
   */
  public getSelectedPigments(): Array<Pigment> {
    var tempPigments = _.filter(this.props.pigments, (pigment: Pigment) => {
      return this.trContainsPigment(this.state.tr, pigment.id);
    });
    var selectedPigments = _.map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        type: (pigment.type) ? pigment.type : "(none)",
        manufacturer: pigment.manufacturer,
        id: pigment.id
      };
    });
    return _.orderBy(selectedPigments, ["title"], ["asc"]);
  }


  /**
   * return tests available from the PropetyTest table and not on the tr
   */
  public getAvailableTests(): Array<DisplayPropertyTest> {
    // select the propertyTest available based on the applicationType and EndUse of the tr
    var temppropertyTest: Array<PropertyTest> = _.filter(this.props.propertyTests, (pt: PropertyTest) => {
      return (pt.applicationTypeid === this.state.tr.ApplicationTypeId
        && pt.endUseIds.indexOf(this.state.tr.EndUseId) != -1
      );
    });
    // get all the the tests in those propertyTests and output them as DisplayPropertyTest
    var tempDisplayTests: Array<DisplayPropertyTest> = [];
    for (const pt of temppropertyTest) {
      console.log(`checking PropertyTest ${pt.id} with property ${pt.property} which has ${pt.testIds.length} tests`);
      for (const testid of pt.testIds) {
        console.log(`Looking for  testid ${testid}`);
        const test: Test = _.find(this.props.tests, (t) => { return t.id === testid; });
        if (test) {
          console.log(`adding test ${test.title}`);
          tempDisplayTests.push({
            property: pt.property,
            testid: testid,
            test: test.title
          });
        }
        else {
          console.log(` test ${testid} NOT FOUND`);
        }

      }
    }
    // now remove those that are already on the tr
    const displayTests = _.filter(tempDisplayTests, (dt) => { return !this.trContainsTest(this.state.tr, dt.testid); });
    return _.orderBy(displayTests, ["type"], ["asc"]);


  }
  public getAvailableTestGroups(): Array<IGroup> {
    var displayPropertyTests: Array<DisplayPropertyTest> = this.getAvailableTests(); // all the avalable tests with their Property
    var properties = _.countBy(displayPropertyTests, (dpt: DisplayPropertyTest) => { return dpt.property; });// an object with an element for each propert, the value of the elemnt is the count of tsts with that property
    var groups: Array<IGroup> = [];
    for (const property in properties) {
      groups.push({
        name: property,
        key: property,
        startIndex: _.findIndex(displayPropertyTests, (dpt) => { return dpt.property === property; }),
        count: properties[property],
        isCollapsed: true
      });
    }
    return groups;
  }
  /**
   * return Tests on the tr
   */
  public getSelectedTests(): Array<Test> {
    var tempTests = _.filter(this.props.tests, (test: Test) => {
      return this.trContainsTest(this.state.tr, test.id);
    });
    var selectedTests = _.map(tempTests, (test) => {
      return {
        title: test.title,
        id: test.id
      };
    });
    return _.orderBy(selectedTests, ["title"], ["asc"]);
  }




  /**
   * return pigments Not on the Tr
   */
  public getAvailablePigments(): Array<Pigment> {
    var tempPigments = _.filter(this.props.pigments, (pigment: Pigment) => {
      return !this.trContainsPigment(this.state.tr, pigment.id);
    });
    var pigments = _.map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        type: (pigment.type) ? pigment.type : "(none)",
        manufacturer: pigment.manufacturer,
        id: pigment.id
      };
    });
    return _.orderBy(pigments, ["type"], ["asc"]);
  }
  public getAvailablePigmentGroups(): Array<IGroup> {
    var pigs: Array<Pigment> = this.getAvailablePigments();
    //var pigmentTypes=_.uniqWith(pigs,(p1:Pigment,p2:Pigment)=>{return p1.type === p2.type});
    var pigmentTypes = _.countBy(pigs, (p1: Pigment) => { return p1.type; });
    var groups: Array<IGroup> = [];
    for (const pt in pigmentTypes) {
      groups.push({
        name: pt,
        key: pt,
        startIndex: _.findIndex(pigs, (pig) => { return pig.type === pt; }),
        count: pigmentTypes[pt],
        isCollapsed: true
      });
    }
    return groups;
  }
  public getTechSpecs() {
    var techSpecs = _.map(this.props.techSpecs, (techSpec) => {
      return {
        title: techSpec.title,
        selected: ((this.state.tr.TRAssignedToId) ? this.state.tr.TRAssignedToId.indexOf(techSpec.id) != -1 : null),
        id: techSpec.id
      };
    });
    return _.orderBy(techSpecs, ["selected", "title"], ["desc", "asc"]);
  }
  public getStaffCC() {
    var staffCC = _.map(this.props.techSpecs, (techSpec) => {
      return {
        title: techSpec.title,
        selected: ((this.state.tr.StaffCCId) ? this.state.tr.StaffCCId.indexOf(techSpec.id) != -1 : null),
        id: techSpec.id
      };
    });
    return _.orderBy(staffCC, ["selected", "title"], ["desc", "asc"]);
  }

  public toggleTechSpec(isSelected: boolean, id: number) {

    this.state.isDirty = true;
    if (isSelected) {
      if (this.state.tr.TRAssignedToId) {
        this.state.tr.TRAssignedToId.push(id);//addit
      }
      else {
        this.state.tr.TRAssignedToId = [id];
      }
    }
    else {
      this.state.tr.TRAssignedToId = _.filter(this.state.tr.TRAssignedToId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }
  public renderTechSpecToggle(item?: any, index?: number, column?: IColumn): any {

    return (
      <Toggle theme={getTheme()} 
        checked={item.selected}
        onText="Selected"
        offText=""
        onChanged={e => { this.toggleTechSpec(e, item.id); }}
      />

    );
  }

  public toggleStaffCC(isSelected: boolean, id: number) {
    this.state.isDirty = true;
    if (isSelected) {
      if (this.state.tr.StaffCCId) {
        this.state.tr.StaffCCId.push(id);//addit
      }
      else {
        this.state.tr.StaffCCId = [id];
      }
    }
    else {
      this.state.tr.StaffCCId = _.filter(this.state.tr.StaffCCId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }
  public renderStaffCCToggle(item?: any, index?: number, column?: IColumn): any {

    return (
      <Toggle theme={getTheme()} 
        checked={item.selected}
        onText="Selected"
        offText=""
        onChanged={e => { this.toggleStaffCC(e, item.id); }}
      />

    );
  }
  /******** TEST Toggles , this is two lists, toggling adds from one , removes from the other*/
  public addTest(id: number) {
    this.state.isDirty = true;
    if (this.state.tr.TestsId) {
      this.state.tr.TestsId.push(id);//addit
    }
    else {
      this.state.tr.TestsId = [id];
    }
    this.setState(this.state);
  }
  public removeTest(id: number) {
    this.state.isDirty = true;
    if (this.state.tr.TestsId) {
      this.state.tr.TestsId = _.filter(this.state.tr.TestsId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }
  public renderAvailableTestsToggle(item?: any, index?: number, column?: IColumn): any {
    return (
      <Toggle theme={getTheme()} 
        checked={false}
        onText=""
        offText=""
        onChanged={e => { this.addTest(item.testid); }}
      />
    );
  }
  public renderSelectedTestsToggle(item?: any, index?: number, column?: IColumn): any {
    return (
      <Toggle theme={getTheme()} 
        checked={true}
        onText=""
        offText=""
        onChanged={e => { this.removeTest(item.id); }}
      />
    );
  }

  /******** Pigmemt Toggles , this is two lists, toggling adds from one , removes from the other*/
  public addPigment(id: number) {
    this.state.isDirty = true;
    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId.push(id);//addit
    }
    else {
      this.state.tr.PigmentsId = [id];
    }
    this.setState(this.state);
  }
  public removePigment(id: number) {

    this.state.isDirty = true;

    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId = _.filter(this.state.tr.PigmentsId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }
  public renderAvailablePigmentsToggle(item?: any, index?: number, column?: IColumn): any {

    return (
      <Toggle theme={getTheme()} 
        checked={false}
        onText=""
        offText=""
        onChanged={e => { this.addPigment(item.id); }}
      />
    );
  }

  public renderSelectedPigmentsToggle(item?: any, index?: number, column?: IColumn): any {
    return (
      <Toggle theme={getTheme()} 
        checked={true}
        onText=""
        offText=""
        onChanged={e => { this.removePigment(item.id); }}
      />
    );
  }

  public renderDate(item?: any, index?: number, column?: IColumn): any {

    return moment(item[column.fieldName]).format("MMM Do YYYY");
  }
  public cancelTrSearch(): void {
    this.state.showTRSearch = false;
    this.setState(this.state);
  }
  //make the child tr the currently selected tr

  public selectChildTR(trId: number): any {
    const childTr = _.find(this.state.childTRs, (tr) => { return tr.Id === trId; });

    if (childTr) {
      console.log("switching to tr " + trId);
      delete this.state.tr;
      this.state.tr = childTr;
      this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);
      this.originalStatus = this.state.tr.TRStatus;
      this.updateCKEditorText(this.state.tr);
      this.state.childTRs = [];

      this.setState(this.state);
      // now get its children, need to move children to state
      this.props.fetchChildTr(this.state.tr.Id).then((trs) => {
        this.state.childTRs = trs;
        this.setState(this.state);
      });
    }

    return false;
  }
  public editDocument(trdocument: TRDocument): void {
    debugger;
    //mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 1).then(url => {
      debugger;
      if (!url || url === "") {
        window.open(trdocument.serverRalativeUrl, '_blank');
      }
      else {
        window.open(url, '_blank');
      }
      //    this.state.wopiFrameUrl=url;
      //  this.setState(this.state);
      // window.location.href=url;

    });

  }
  public rendeDocumentAsLink(item?: any, index?: number, column?: IColumn): JSX.Element {
    return (
      <div>
        <i onClick={(e) => { debugger; this.editDocument(item); }}
          className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
      </div>
    );
  }
  public rendeChildTRAsLink(item?: any, index?: number, column?: IColumn): JSX.Element {
    return (
      <div>
        <i onClick={(e) => {
          debugger;
          this.selectChildTR(item.Id);
        }}
          className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
      </div>
    );
  }
  public parentTRSelected(id: number, title: string) {
    this.state.tr.ParentTR = title;
    this.state.tr.ParentTRId = id;
    this.state.isDirty = true;
    this.cancelTrSearch();
  }
  public uploadFile(e: any) {
    debugger;
    let target: any = e.target as any;
    let file = e.target["files"][0];
    this.props.uploadFile(file, this.state.tr.Id).then((response) => {
      this.props.getDocuments(this.state.tr.Id).then((dox) => {
        this.state.documents = dox;
        this.setState(this.state);
      });
      debugger;
    }).catch((error) => {
      debugger;
    });
  }
  public editParentTR() {

    if (this.state.tr.ParentTRId) {
      const parentId = this.state.tr.ParentTRId;
      this.props.fetchTR(parentId).then((parentTR) => {

        this.state.tr = parentTR;
        this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);
        this.originalStatus = this.state.tr.TRStatus;
        this.state.childTRs = [];
        this.setState(this.state);
        this.updateCKEditorText(this.state.tr);
        this.props.fetchChildTr(parentId).then((subTRs) => {
          this.state.childTRs = subTRs;
          this.setState(this.state);
        });
      });
    }


  }
  public documentRowMouseEnter(trdocument: TRDocument, e: any) {
    //mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 3).then(url => {
      if (!url || url === "") {
        url = trdocument.serverRalativeUrl;
      }
      this.state.documentCalloutIframeUrl = url;
      this.state.documentCalloutTarget = e.target;
      this.state.documentCalloutVisible = true;
      this.setState(this.state);

    });
  }
  public documentRowMouseOut(item: TRDocument, e: any) {

    this.state.documentCalloutTarget = null;
    this.state.documentCalloutVisible = false;
    this.setState(this.state);
    console.log("mouse exit for " + item.title);
  }
  public renderDocumentRow(props, defaultRender): JSX.Element {

    return (
      <div
        onMouseEnter={(event) => this.documentRowMouseEnter(props.item, event)}
        onMouseOut={(evemt) => this.documentRowMouseOut(props.item, event)}
      >
        {defaultRender(props)}
      </div>);
  }
  public render(): React.ReactElement<ITrFormProps> {

    let worktypeDropDoownoptions = _.map(this.props.workTypes, (wt) => {
      return {
        key: wt.id,
        text: wt.workType
      };
    });

    let applicationtypeDropDoownoptions =
      _.filter(this.props.applicationTypes, (at) => {
        // show if its valid for the selected Worktype, OR if its already on the tr
        return (at.workTypeIds.indexOf(this.state.tr.WorkTypeId) !== -1
          || at.id === this.state.tr.ApplicationTypeId);

      })
        .map((at) => {
          return {
            key: at.id,
            text: at.applicationType
          };
        });
    let enduseDropDoownoptions =
      _.filter(this.props.endUses, (eu) => {
        // show if its valid for the selected ApplicationType, OR if its already on the tr
        return (eu.applicationTypeId === this.state.tr.ApplicationTypeId
          || eu.id === this.state.tr.EndUseId);
      })
        .map((eu) => {
          return {
            key: eu.id,
            text: eu.endUse
          };
        });
    console.log("# of app types is " + applicationtypeDropDoownoptions.length);
    return (
      <div >

        <MessageDisplay messages={this.state.errorMessages}
          hideMessage={this.removeMessage.bind(this)} />
        <div style={{ float: "left" }}> <this.ModeDisplay /></div>
        <div style={{ float: "right" }}><this.StatusDisplay /></div>
        <div style={{ clear: "both" }}></div>
        <table>

          <tr>
            <td>
              <Label >Request #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Title} onChanged={e => {
                this.state.isDirty = true;
                this.state.tr.Title = e; this.setState(this.state);
              }} />
            </td>
            <td>
              <Label >Work Type</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.WorkTypeId}
                options={worktypeDropDoownoptions}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.WorkTypeId = e.key as number;
                  this.setState(this.state);
                }} />
            </td>

            <td>
              <Label >Site</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Site} onChanged={e => {
                this.state.isDirty = true;
                this.state.tr.Site = e;
                this.setState(this.state);
              }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label  >Parent TR</Label>
            </td>
            <td>
              <div>
                <Label style={{ "display": "inline" }}>
                  {this.state.tr.ParentTR}
                </Label>
                <i onClick={this.editParentTR}
                  className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                <i onClick={(e) => { this.state.showTRSearch = true; this.setState(this.state); }}
                  className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>

              </div>

            </td>
            <td>
              <Label >Application Type</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.ApplicationTypeId}
                options={applicationtypeDropDoownoptions}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.ApplicationTypeId = e.key as number;
                  this.setState(this.state);
                }}
              />
            </td>
            <td>
              <Label >Priority</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={[
                  { key: 'High', text: 'High' },
                  { key: 'Normal', text: 'Normal' },
                  { key: 'Routine', text: 'Routine' },

                ]}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.TRPriority = e.text;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.TRPriority} />
            </td>

          </tr>
          <tr>
            <td>
              <Label>CER #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.CER} onChanged={e => { this.state.isDirty = true; this.state.tr.CER = e; }} />
            </td>
            <td>
              <Label>Requestor</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={this.props.requestors.map((r) => { return { key: r.id, text: r.title }; })}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.RequestorId = e.key as number;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.RequestorId}
              />
            </td>
            <td>
              <Label>Customer</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={this.props.customers.map((r) => { return { key: r.id, text: r.title }; })}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.CustomerId = e.key as number;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.CustomerId}
              />

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Initiation Date</Label>
            </td>
            <td>

              <DatePicker  
                value={(this.state.tr.RequestDate) ? moment(this.state.tr.RequestDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.RequestDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >End Use</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.EndUseId}
                options={enduseDropDoownoptions}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.EndUseId = e.key as number;
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Status</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={[
                  { key: 'Pending', text: 'Pending' },
                  { key: 'In Progress', text: 'In Progress' },
                  { key: 'Completed', text: 'Completed' },
                  { key: 'Canceled', text: 'Canceled' },
                ]}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.TRStatus = e.text;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.TRStatus} />

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Due Date</Label>
            </td>
            <td>

              <DatePicker
                value={(this.state.tr.RequiredDate) ? moment(this.state.tr.RequiredDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.RequiredDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Actual Start Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualStartDate) ? moment(this.state.tr.ActualStartDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.ActualStartDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Actual Completion Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualCompletionDate) ? moment(this.state.tr.ActualCompletionDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.ActualCompletionDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label  >Estimated Hours</Label>
            </td>
            <td>

              <TextField datatype="number"
                value={(this.state.tr.EstManHours) ? this.state.tr.EstManHours.toString() : null}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.EstManHours = parseInt(e);
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label ></Label>
            </td>


            <td>
              <Label ></Label>
            </td>
            <td>

            </td>

          </tr>


        </table>
        <tabs.Tabs onSelect={this.tabChanged.bind(this)}>
          <tabs.TabList>
            <tabs.Tab>
              Title
             </tabs.Tab>
            <tabs.Tab>
              Description
             </tabs.Tab>
            <tabs.Tab>
              Summary
             </tabs.Tab>
            <tabs.Tab>
              Test Params
             </tabs.Tab>
            <tabs.Tab>
              Assigned To({(this.state.tr.TRAssignedToId) ? this.state.tr.TRAssignedToId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Staff cc({(this.state.tr.StaffCCId) ? this.state.tr.StaffCCId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Pigments({(this.state.tr.PigmentsId) ? this.state.tr.PigmentsId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Tests({(this.state.tr.TestsId) ? this.state.tr.TestsId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Formulae
             </tabs.Tab>
            <tabs.Tab>
              Child TRs({(this.state.childTRs) ? this.state.childTRs.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Documents
             </tabs.Tab>
          </tabs.TabList>
          <tabs.TabPanel >

            <textarea name="tronoxtrtextarea-title" id="tronoxtrtextarea-title" style={{ display: "none" }}>
              {this.state.tr.RequestTitle}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-description" id="tronoxtrtextarea-description" style={{ display: "none" }}>
              {this.state.tr.Description}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-summary" id="tronoxtrtextarea-summary" style={{ display: "none" }}>
              {this.state.tr.Summary}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-testparams" id="tronoxtrtextarea-testparams" style={{ display: "none" }}>
              {this.state.tr.TestingParameters}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              items={this.getTechSpecs()}
              setKey="id"
              columns={[
                { key: "title", name: "Technical Specialist", fieldName: "title", minWidth: 20, maxWidth: 200 },
                { key: "selected", name: "Assigned?", fieldName: "selected", minWidth: 200, onRender: this.renderTechSpecToggle.bind(this) }
              ]}
            />
          </tabs.TabPanel>
          <tabs.TabPanel>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              items={this.getStaffCC()}
              setKey="id"
              columns={[
                { key: "title", name: "Staff", fieldName: "title", minWidth: 20, maxWidth: 200 },
                { key: "selected", name: "cc'd?", fieldName: "selected", minWidth: 80, onRender: this.renderStaffCCToggle.bind(this) }
              ]}
            />
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div style={{ float: "left" }}>
              <Label> Available Pigments</Label>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                groups={this.getAvailablePigmentGroups()}
                items={this.getAvailablePigments()}
                setKey="id"
                columns={[
                  { key: "title", name: "Pigment Name", fieldName: "title", minWidth: 20, maxWidth: 100 },
                  { key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: this.renderAvailablePigmentsToggle.bind(this) }
                ]}
              />

            </div>
            <div style={{ float: "right" }}>
              <Label> Selected Pigments</Label>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={this.getSelectedPigments()}
                setKey="id"
                columns={[
                  { key: "title", name: "Pigment Name", fieldName: "title", minWidth: 20, maxWidth: 100 },
                  { key: "type", name: "Type", fieldName: "type", minWidth: 20, maxWidth: 100 },
                  { key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: this.renderSelectedPigmentsToggle.bind(this) }
                ]}
              />
            </div>
            <div style={{ clear: "both" }}></div>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div style={{ float: "left" }}>
              <Label> Available Tests</Label>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                groups={this.getAvailableTestGroups()}
                items={this.getAvailableTests()}
                setKey="id"
                columns={[
                  { key: "title", name: "test", fieldName: "test", minWidth: 20, maxWidth: 100 },
                  { key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: this.renderAvailableTestsToggle.bind(this) }
                ]}
              />

            </div>
            <div style={{ float: "right" }}>
              <Label> Selected Tests</Label>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={this.getSelectedTests()}
                setKey="id"
                columns={[
                  { key: "title", name: "Test Name", fieldName: "title", minWidth: 20, maxWidth: 200 },
                  { key: "selected", name: "Selected?", fieldName: "selected", minWidth: 200, onRender: this.renderSelectedTestsToggle.bind(this) }
                ]}
              />

            </div>
            <div style={{ clear: "both" }}></div>



          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-formulae" id="tronoxtrtextarea-formulae" style={{ display: "none" }}>
              {this.state.tr.Formulae}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>

            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.state.childTRs}
              setKey="id"
              selectionMode={SelectionMode.none}
              columns={[
                { key: "Edit", onRender: this.rendeChildTRAsLink, name: "", fieldName: "Title", minWidth: 20, },
                { key: "Title", name: "Request #", fieldName: "Title", minWidth: 80, },
                { key: "Status", name: "Status", fieldName: "Status", minWidth: 90 },
                { key: "InitiationDate", onRender: this.renderDate, name: "Initiation Date", fieldName: "InitiationDate", minWidth: 80 },
                { key: "TRDueDate", onRender: this.renderDate, name: "Due Date", fieldName: "TRDueDate", minWidth: 80 },
                { key: "ActualStartDate", onRender: this.renderDate, name: "Actual Start Date", fieldName: "ActualStartDate", minWidth: 90 },
                { key: "ActualCompetionDate", onRender: this.renderDate, name: "Actual Competion<br />Date", fieldName: "ActualCompetionDate", minWidth: 80 },
              ]}
            />
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div style={{ float: "left" }}>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                items={this.state.documents}
                onRenderRow={(props, defaultRender) => this.renderDocumentRow(props, defaultRender)}
                setKey="id"
                selectionMode={SelectionMode.none}
                columns={[
                  { key: "Edit", onRender: this.rendeDocumentAsLink, name: "", fieldName: "Title", minWidth: 20, },
                  { key: "title", name: "Request #", fieldName: "title", minWidth: 80, },

                ]}
              />
              <input type='file' id='uploadfile' onChange={e => { debugger; this.uploadFile(e); }} />
            </div>
            <div style={{ float: "right" }}>
              <DocumentIframe src={this.state.documentCalloutIframeUrl} />
            </div>
            <div style={{ clear: "both" }}></div>


          </tabs.TabPanel>
        </tabs.Tabs>

        <this.SaveButton />
        <span style={{ margin: 20 }}>
          <PrimaryButton theme={getTheme()} href="#" onClick={this.cancel} icon="ms-Icon--Cancel">
            <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
            Cancel
        </PrimaryButton>
        </span>
        <TRPicker
          isOpen={this.state.showTRSearch}
          callSearch={this.props.TRsearch}
          cancel={this.cancelTrSearch}
          select={this.parentTRSelected}
        />
      </div >
    );
  }
}
