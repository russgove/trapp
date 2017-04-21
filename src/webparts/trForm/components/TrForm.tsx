import * as React from 'react';
import styles from './TrForm.module.scss';
import { ITrFormProps } from './ITrFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TR } from "../dataModel";
import { TextField, Label, Button, ButtonType, MessageBar, MessageBarType, DatePicker, Dropdown, IDropdownProps } from 'office-ui-fabric-react';
export interface inITrFormState {
  tr: TR;
  errorMessages: Array<md.Message>;
}
import * as moment from 'moment';
import * as _ from "lodash";
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
export default class TrForm extends React.Component<ITrFormProps, inITrFormState> {

  constructor(props: ITrFormProps) {
    super(props);

    this.state = {
      tr: props.tr,
      errorMessages: []
    };
  }

  public componentWillReceiveProps(nextProps: ITrFormProps, nextContext: any) {

  }
  public save() {

    this.props.save(this.state.tr)
      .then((result) => { })
      .catch((response) => {
        this.state.errorMessages.push(new md.Message(response.data.responseBody['odata.error'].message.value));
        this.setState(this.state);
      });

  }
  public removeMessage(messageList: Array<md.Message>, messageId: string) {
    _.remove(messageList, {
      Id: messageId
    });
    this.setState(this.state);
  }
  public render(): React.ReactElement<ITrFormProps> {

    console.log("WorkTypeID is" + this.props.tr.WorkTypeId);
    let worktypeDropDoownoptions = _.map(this.props.workTypes, (wt) => {
      return {
        key: wt.id,
        text: wt.workType
      }
    });
    let applicationtypeDropDoownoptions =
      _.filter(this.props.applicationTypes, (at) => {

        return at.workTypeIds.indexOf(this.props.tr.WorkTypeId) !== -1
      })
        .map((at) => {
          return {
            key: at.id,
            text: at.applicationType
          }
        });
    let enduseDropDoownoptions =
      _.filter(this.props.endUses, (eu) => {

        return (eu.applicationTypeId === this.props.tr.ApplicationTypeId)
      })
        .map((eu) => {
          return {
            key: eu.id,
            text: eu.endUse
          }
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
                debugger;
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
              <TextField value={this.state.tr.Site} onChanged={e => { this.state.tr.Site = e }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label value='Request #' >Parent TR</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Id.toString()} readOnly={true} />
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
              <TextField value={this.state.tr.MailBox} onChanged={e => { this.state.tr.MailBox = e }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label value='Request #' >CER #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.CER} readOnly={true} onChanged={e => { this.state.tr.CER = e }} />
            </td>
            <td>

            </td>
            <td>

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
              <TextField value={this.state.tr.Customer} onChanged={e => { this.state.tr.Customer = e }} />
            </td>

          </tr>


        </table>
        <Button buttonType={ButtonType.normal} onClick={this.save.bind(this)}>Save</Button>
      </div>
    );
  }
}
