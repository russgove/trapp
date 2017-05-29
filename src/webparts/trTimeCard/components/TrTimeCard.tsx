import * as React from 'react';
import { ITrTimeCardProps } from './ITrTimeCardProps';
import { ITrTimeCardState } from './ITrTimeCardState';
import { escape } from '@microsoft/sp-lodash-subset';
///// Add tab for testinParameters text block


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
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as _ from "lodash";
import { TimeSpent } from "../dataModel";
export default class TrTimeCard extends React.Component<ITrTimeCardProps, ITrTimeCardState> {
  constructor(props: ITrTimeCardProps) {
    debugger;
    super(props);
    this.state = props.initialState;
    this.setState(this.state);
    this.getDisplayTRs = this.getDisplayTRs.bind(this);
  }
  public getDisplayTRs(): Array<TimeSpent> {
    return this.state.timeSpents;
  }

  public render(): React.ReactElement<ITrTimeCardProps> {
    debugger;
    return (
      <div>
        <Label>Time spent by Technical Specialist {this.props.userName} for the week ending {this.state.weekEndingDate.toDateString()}   </Label>
        <Label>(if you do not see a TR you are working on displayed here, please ask you adminstrator to assign it to you, or to reopen it) </Label>
        
        <DetailsList
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.none}
          items={this.getDisplayTRs()}
          setKey="id"
          columns={[
            { key: "Id", name: "Id", fieldName: "Id", minWidth: 20, maxWidth: 200 },
            { key: "TR", name: "TR", fieldName: "TR.title", minWidth: 20, maxWidth: 200 },
            { key: "HoursSpent", name: "HoursSpent", fieldName: "HoursSpent", minWidth: 200 }
          ]}
        />
      </div>

    );
  }
}
