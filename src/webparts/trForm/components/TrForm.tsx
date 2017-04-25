import * as React from 'react';
import styles from './TrForm.module.scss';
import { ITrFormProps } from './ITrFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TR } from "../dataModel";
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
export interface inITrFormState {
  tr: TR;
  errorMessages: Array<md.Message>;
  resultsPersonas: Array<IPersonaProps>;
}
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as _ from "lodash";
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
import * as tabs from "react-tabs";


export default class TrForm extends React.Component<ITrFormProps, inITrFormState> {
  private ckeditor: any;

  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  constructor(props: ITrFormProps) {
    super(props);
    this.state = {
      tr: props.tr,
      errorMessages: [],
      resultsPersonas: []
    };


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
      default:

    };
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
      default:

    }



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
        default:

      }
    }
    this.props.save(this.state.tr)
      .then((result) => { })
      .catch((response) => {
        this.state.errorMessages.push(new md.Message(response.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      });

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
    debugger;
    if (req.length > 0) {
      console.log("requestor changedd " + req[0].optionalText);// I am only adding a single user. req[0] , others are ignored
      const email = req[0].optionalText;
      this.props.ensureUser(email).then((user) => {
        this.state.tr.RequestorId = user.data.Id;
        this.setState(this.state);
      }).catch((error) => {
        this.state.errorMessages.push(new md.Message(error.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      })
    }
    else {
      console.log("requestor removed ");// I am only adding a single user. req[0] , others are ignored
      this.state.tr.RequestorId = null;
    }
  }
  public renderPeople(person: IPersonaProps): JSX.Element {

    return <Persona
      size={PersonaSize.large}
      primaryText={person.primaryText}
      secondaryText={person.secondaryText}

      tertiaryText={person.tertiaryText}
      imageUrl={person.imageUrl}
      imageShouldFadeIn={true}

    />

  }

  public renderTR(person: IPersonaProps): JSX.Element {

    return <Persona
      size={PersonaSize.extraLarge}
      primaryText={person.primaryText}
      secondaryText={person.secondaryText}

      tertiaryText={person.tertiaryText}
      imageUrl={person.imageUrl}
      imageInitials=""

    />

  }
  public render(): React.ReactElement<ITrFormProps> {
    const suggestionProps: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: 'Suggested People',
      noResultsFoundText: 'No results found',
      loadingText: 'Loading',

    };
    console.log("WorkTypeID is" + this.props.tr.WorkTypeId);
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
        <table>

          <tr>
            <td>
              <Label >Request #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Title} onChanged={e => {
                this.state.tr.Title = e; this.setState(this.state);
              }} />
            </td>
            <td>
              <Label >Work Type</Label>
            </td>
            <td>
              <Dropdown label='' selectedKey={this.state.tr.WorkTypeId} options={worktypeDropDoownoptions} onChanged={e => {

                this.state.tr.WorkTypeId = e.key as number;
                console.log("WorkType changing to " + this.state.tr.WorkTypeId);
                this.setState(this.state);
                console.log("WorkType changed to " + this.state.tr.WorkTypeId);
              }} />
            </td>

            <td>
              <Label >Site</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Site} onChanged={e => { this.state.tr.Site = e; }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label  >Parent TR</Label>
            </td>
            <td>

              <NormalPeoplePicker
                defaultSelectedItems={this.state.tr.ParentTRId ? [{ id: this.state.tr.ParentTRId.toString(), primaryText: this.state.tr.ParentTR }] : []}
                onResolveSuggestions={this.resolveSuggestionsTR.bind(this)}
                pickerSuggestionsProps={suggestionProps}
                getTextFromItem={this.getTextFromItem}
                onRenderSuggestionsItem={this.renderTR}
                onChange={e => { console.log("TR changedd" + e); debugger; this.state.tr.ParentTRId = parseInt(e[0].id); }}
              />
            </td>
            <td>
              <Label >Application Type</Label>
            </td>
            <td>
              <Dropdown label='' selectedKey={this.state.tr.ApplicationTypeId} options={applicationtypeDropDoownoptions} onChanged={e => { debugger; this.state.tr.ApplicationTypeId = e.key as number; this.setState(this.state); }} />
            </td>
            <td>
              <Label >MailBox</Label>
            </td>
            <td>
              <TextField value={this.state.tr.MailBox} onChanged={e => { this.state.tr.MailBox = e; }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label>CER #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.CER} onChanged={e => { this.state.tr.CER = e; }} />
            </td>
            <td>
              Requestor
            </td>
            <td>
              <NormalPeoplePicker
                defaultSelectedItems={this.state.tr.RequestorId ? [{ id: this.state.tr.RequestorId.toString(), primaryText: this.state.tr.RequestorName }] : []}
                onResolveSuggestions={this.resolveSuggestions.bind(this)}
                pickerSuggestionsProps={suggestionProps}
                getTextFromItem={this.getTextFromItem}
                onRenderSuggestionsItem={this.renderPeople}
                onChange={this.requestorChanged.bind(this)}
              />
            </td>
            <td>

            </td>
            <td>

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Initiation Date</Label>
            </td>
            <td>

              <DatePicker value={moment(this.state.tr.InitiationDate).toDate()} onSelectDate={e => { this.state.tr.InitiationDate = moment(e).toISOString(); }} />
            </td>
            <td>
              <Label >End Use</Label>
            </td>
            <td>
              <Dropdown label='' selectedKey={this.state.tr.EndUseId} options={enduseDropDoownoptions} onChanged={e => { debugger; this.state.tr.EndUseId = e.key as number; this.setState(this.state); }} />
            </td>
            <td>
              <Label >Customer</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Customer} onChanged={e => { this.state.tr.Customer = e; }} />
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
              tech Spec
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
              formulae
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
            <h2>these are the test pareameters</h2>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <h2>Specification</h2>
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
            <h2>formulae?></h2>
          </tabs.TabPanel>
        </tabs.Tabs>
        <Button buttonType={ButtonType.normal} onClick={this.save.bind(this)}>Save</Button>
      </div>
    );
  }
}
