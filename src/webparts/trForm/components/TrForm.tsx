import * as React from 'react';
import styles from './TrForm.module.scss';
import { ITrFormProps } from './ITrFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TR, modes } from "../dataModel";
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as _ from "lodash";
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
import * as tabs from "react-tabs";
export interface inITrFormState {
  tr: TR;
  childTRs:Array<TR>,
  errorMessages: Array<md.Message>;
  resultsPersonas: Array<IPersonaProps>;
  isDirty: boolean;
}

export default class TrForm extends React.Component<ITrFormProps, inITrFormState> {
  private ckeditor: any;

  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  constructor(props: ITrFormProps) {
    super(props);
    this.state = {
      tr: props.tr,
      childTRs:props.subTRs,
      errorMessages: [],
      resultsPersonas: [],
      isDirty: false,
    };
    this.SaveButton = this.SaveButton.bind(this);
    this.ModeDisplay = this.ModeDisplay.bind(this);
    this.StatusDisplay = this.StatusDisplay.bind(this);
    this.save = this.save.bind(this);
    this.cancel = this.cancel.bind(this);
    this.rendeChildTRAsLink = this.rendeChildTRAsLink.bind(this);
    this.selectChildTR = this.selectChildTR.bind(this);

  }

  public componentWillReceiveProps(nextProps: ITrFormProps, nextContext: any) {

  }
  public componentDidMount() {

    var ckEditorCdn: string = '//cdn.ckeditor.com/4.6.2/full/ckeditor.js';
    SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: 'CKEDITOR' }).then((CKEDITOR: any): void => {
      this.ckeditor = CKEDITOR;
      this.ckeditor.replace("tronoxtrtextarea-title");

    });


  }
  public tabChanged(newTabID, oldTabID) {

    switch (oldTabID) {
      case 0:
        let data = this.ckeditor.instances["tronoxtrtextarea-title"].getData();
        this.ckeditor.remove("tronoxtrtextarea-title");
        console.log("removed tronoxtrtextarea-title");
        break;
      case 1:
        let data1 = this.ckeditor.instances["tronoxtrtextarea-description"].getData();
        this.ckeditor.remove("tronoxtrtextarea-description");
        console.log("removed tronoxtrtextarea-description");
        break;
      case 2:
        let data2 = this.ckeditor.instances["tronoxtrtextarea-summary"].getData();
        this.ckeditor.remove("tronoxtrtextarea-summary");
        console.log("removed tronoxtrtextarea-summary");
        break;
      case 3:
        let data3 = this.ckeditor.instances["tronoxtrtextarea-testparams"].getData();
        this.ckeditor.remove("tronoxtrtextarea-testparams");
        console.log("removed tronoxtrtextarea-testparams");
        break;
      case 8:
        let data4 = this.ckeditor.instances["tronoxtrtextarea-formulae"].getData();
        this.ckeditor.remove("tronoxtrtextarea-formulae");
        console.log("removed tronoxtrtextarea-formulae");
        break;
      default:

    }
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
    if (!this.state.tr.InitiationDate) {
      this.state.errorMessages.push(new md.Message("Initiation Date   is required"));
      errorsFound = true;
    }
    if (!this.state.tr.TRDueDate) {
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
    if (!this.state.tr.Status) {
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
          this.state.tr.TitleArea = data;
          break;
        case "tronoxtrtextarea-description":
          this.state.tr.DescriptionArea = data;
          break;
        case "tronoxtrtextarea-summary":
          this.state.tr.SummaryArea = data;
          break;
        case "tronoxtrtextarea-formulae":
          this.state.tr.FormulaeArea = data;
          break;
        default:

      }
    }

    if (this.isValid()) {
      this.props.save(this.state.tr)
        .then((result) => {
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

  public resolveSuggestions(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps> | IPersonaProps[] {

    return this.props.peoplesearch(searchText, currentSelected);
  }
  public resolveSuggestionsTR(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps> | IPersonaProps[] {

    return this.props.TRsearch(searchText, currentSelected);
  }
  public removeMessage(messageList: Array<md.Message>, messageId: string) {
    _.remove(messageList, {
      Id: messageId
    });
    this.setState(this.state);
  }
  public getTextFromItem(persona: IPersonaProps): string {

    return persona.primaryText;
  }
  public requestorChanged(req: Array<IPersonaProps>) { // need to call ensure user
    this.state.isDirty = true;
    if (req.length > 0) {
      console.log("requestor changedd " + req[0].optionalText);// I am only adding a single user. req[0] , others are ignored
      const email = req[0].optionalText;
      this.props.ensureUser(email).then((user) => {
        this.state.tr.RequestorId = user.data.Id;
        this.setState(this.state);
      }).catch((error) => {
        this.state.errorMessages.push(new md.Message(error.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      });
    }
    else {
      console.log("requestor removed ");// I am only adding a single user. req[0] , others are ignored
      this.state.tr.RequestorId = null;
    }
  }
  public renderPeople(person: IPersonaProps): JSX.Element {

    return (<Persona
      size={PersonaSize.large}
      primaryText={person.primaryText}
      secondaryText={person.secondaryText}

      tertiaryText={person.tertiaryText}
      imageUrl={person.imageUrl}
      imageShouldFadeIn={true}

    />);

  }

  public renderTR(person: IPersonaProps): JSX.Element {

    return (<Persona
      size={PersonaSize.extraLarge}
      primaryText={person.primaryText}
      secondaryText={person.secondaryText}

      tertiaryText={person.tertiaryText}
      imageUrl={person.imageUrl}
      imageInitials=""

    />);

  }

  public SaveButton(): JSX.Element {
    if (this.props.mode === modes.DISPLAY) {
      return <div />;
    } else return (
      <span style={{ margin: 20 }}>
        <a href="#" onClick={this.save} style={{ border: 5, backgroundColor: 'lightBlue', fontSize: 'large' }}>
          Save
        </a>
      </span>
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
  public getTechSpecs() {

    var x = _.map(this.props.techSpecs, (techSpec) => {
      return {
        title: techSpec.title,
        selected: ((this.state.tr.TechSpecId) ? this.state.tr.TechSpecId.indexOf(techSpec.id) != -1 : null),
        id: techSpec.id
      };
    });
    return _.sortBy(x, "selected").reverse();

  }
  public toggleTechSpec(isSelected: boolean, id: number) {

    this.state.isDirty = true;
    if (isSelected) {
      if (this.state.tr.TechSpecId) {
        this.state.tr.TechSpecId.push(id);//addit
      }
      else {
        this.state.tr.TechSpecId = [id];
      }
    }
    else {
      this.state.tr.TechSpecId = _.filter(this.state.tr.TechSpecId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }
  public renderToggle(item?: any, index?: number, column?: IColumn): any {

    return (
      <Toggle
        checked={item.selected}
        onText="On Team"
        offText=""
        onChanged={e => { this.toggleTechSpec(e, item.id); }}
      />

    );
  }
  //make the child tr the currently selected tr
  public selectChildTR(trId: number): any {
    const childTr = _.find(this.state.childTRs, (tr) => { return tr.Id === trId; });
    debugger;
    if (childTr) {
      console.log("switching to tr " + trId);
      delete this.state.tr;
      this.state.tr = childTr
      this.state.childTRs=[];
      this.setState(this.state);
      // now get its childerm, need to move children to state
      this.props.getChildTr(this.state.tr.Id).then((trs)=>{
        this.state.childTRs=trs;
        this.setState(this.state);
      });
    }

    return false;
  }
  public rendeChildTRAsLink(item?: any, index?: number, column?: IColumn): JSX.Element {
    debugger;
    return (<a href="#" onClick={(e) => { debugger; this.selectChildTR(item.Id) }}>
      {item[column.fieldName]}
    </a>);
  }
  public renderDate(item?: any, index?: number, column?: IColumn): any {

    return moment(item[column.fieldName]).format("MMM Do YYYY");
  }
  public render(): React.ReactElement<ITrFormProps> {
    const suggestionProps: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: 'Suggested People',
      noResultsFoundText: 'No results found',
      loadingText: 'Loading',

    };
    let worktypeDropDoownoptions = _.map(this.props.workTypes, (wt) => {
      return {
        key: wt.id,
        text: wt.workType
      };
    });
    let applicationtypeDropDoownoptions =
      _.filter(this.props.applicationTypes, (at) => {

        return at.workTypeIds.indexOf(this.props.tr.WorkTypeId) !== -1;
      })
        .map((at) => {
          return {
            key: at.id,
            text: at.applicationType
          };
        });
    let enduseDropDoownoptions =
      _.filter(this.props.endUses, (eu) => {
        return (eu.applicationTypeId === this.props.tr.ApplicationTypeId);
      })
        .map((eu) => {
          return {
            key: eu.id,
            text: eu.endUse
          };
        });
    console.log("# of app types is " + applicationtypeDropDoownoptions.length);
    return (
      <div>

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
              <TextField value={this.state.tr.ParentTR} onChanged={e => {
                this.state.isDirty = true;
                this.state.tr.ParentTR = e;
                this.setState(this.state);
              }} />
              {/*<NormalPeoplePicker
                defaultSelectedItems={this.state.tr.ParentTRId ? [{ id: this.state.tr.ParentTRId.toString(), primaryText: this.state.tr.ParentTR }] : []}
                onResolveSuggestions={this.resolveSuggestionsTR.bind(this)}
                pickerSuggestionsProps={suggestionProps}
                getTextFromItem={this.getTextFromItem}
                onRenderSuggestionsItem={this.renderTR}
                onChange={e => {
                  this.state.isDirty = true;
                  this.state.tr.ParentTRId = (e.length > 0) ? parseInt(e[0].id) : null;
                  this.setState(this.state);
                }}
              />*/}
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
                  { key: 'Medium', text: 'Medium' },
                  { key: 'Low', text: 'Low' },

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
              {/*<NormalPeoplePicker
                defaultSelectedItems={this.state.tr.RequestorId ? [{ id: this.state.tr.RequestorId.toString(), primaryText: this.state.tr.RequestorName }] : []}
                onResolveSuggestions={this.resolveSuggestions.bind(this)}
                pickerSuggestionsProps={suggestionProps}
                getTextFromItem={this.getTextFromItem}
                onRenderSuggestionsItem={this.renderPeople}
                onChange={this.requestorChanged.bind(this)}
              />*/}
            </td>
            <td>
              <Label>Customer</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Customer}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.Customer = e;
                  this.setState(this.state);
                }} />

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Initiation Date</Label>
            </td>
            <td>

              <DatePicker
                value={(this.state.tr.InitiationDate) ? moment(this.state.tr.InitiationDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.InitiationDate = moment(e).toISOString();
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
                  { key: 'Complete', text: 'Complete' },
                  { key: 'Canceled', text: 'Canceled' },
                ]}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.Status = e.text;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.Status} />

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Due Date</Label>
            </td>
            <td>

              <DatePicker
                value={(this.state.tr.TRDueDate) ? moment(this.state.tr.InitiationDate).toDate() : null}
                onSelectDate={e => {
                  this.state.isDirty = true;
                  this.state.tr.TRDueDate = moment(e).toISOString();
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
                value={(this.state.tr.EstimatedHours) ? this.state.tr.EstimatedHours.toString() : null}
                onChanged={e => {
                  this.state.isDirty = true;
                  this.state.tr.EstimatedHours = parseInt(e);
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
              Tech Spec({(this.state.tr.TechSpecId) ? this.state.tr.TechSpecId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              staff cc
             </tabs.Tab>
            <tabs.Tab>
              pigments
             </tabs.Tab>
            <tabs.Tab>
              Tests
             </tabs.Tab>
            <tabs.Tab>
              Formulae
             </tabs.Tab>
            <tabs.Tab>
              Child TRs({(this.state.childTRs) ? this.state.childTRs.length : 0})
             </tabs.Tab>
          </tabs.TabList>
          <tabs.TabPanel >

            <textarea name="tronoxtrtextarea-title" id="tronoxtrtextarea-title" style={{ display: "none" }}>
              {this.state.tr.TitleArea}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-description" id="tronoxtrtextarea-description" style={{ display: "none" }}>
              {this.state.tr.DescriptionArea}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-summary" id="tronoxtrtextarea-summary" style={{ display: "none" }}>
              {this.state.tr.SummaryArea}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-testparams" id="tronoxtrtextarea-testparams" style={{ display: "none" }}>
              {this.state.tr.TestParamsArea}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.getTechSpecs()}
              setKey="id"
              columns={[
                { key: "title", name: "Tecnical Specialis Name", fieldName: "title", minWidth: 20, maxWidth: 200 },
                { key: "selected", name: "On Team?", fieldName: "selected", minWidth: 200, onRender: this.renderToggle.bind(this) }
              ]}
            />
          </tabs.TabPanel>
          <tabs.TabPanel>
            <h2>staff cc? just sen emails. or set notifications></h2>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <h2>pigments incolve</h2>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <h2>these are teh tests</h2>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-formulae" id="tronoxtrtextarea-formulae" style={{ display: "none" }}>
              {this.state.tr.FormulaeArea}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>

            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.state.childTRs}
              setKey="id"
              columns={[
                { key: "Title", onRender: this.rendeChildTRAsLink, name: "Request #", fieldName: "Title", minWidth: 80, },
                { key: "Status", name: "Status", fieldName: "Status", minWidth: 90 },
                { key: "InitiationDate", onRender: this.renderDate, name: "Initiation Date", fieldName: "InitiationDate", minWidth: 80 },
                { key: "TRDueDate", onRender: this.renderDate, name: "Due Date", fieldName: "TRDueDate", minWidth: 80 },
                { key: "ActualStartDate", onRender: this.renderDate, name: "Actual Start Date", fieldName: "ActualStartDate", minWidth: 90 },
                { key: "ActualCompetionDate", onRender: this.renderDate, name: "Actual Competion<br />Date", fieldName: "ActualCompetionDate", minWidth: 80 },

              ]}
            />
          </tabs.TabPanel>
        </tabs.Tabs>

        <this.SaveButton />
        <span style={{ margin: 20 }}>
          <a href="#" onClick={this.cancel} style={{ border: 5, backgroundColor: 'lightBlue', fontSize: 'large' }}>
            Cancel
        </a>
        </span>

      </div>
    );
  }
}
