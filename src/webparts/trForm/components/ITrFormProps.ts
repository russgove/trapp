import { TR, ApplicationType, WorkType, EndUse, modes,User } from "../dataModel";
import {
  IPersonaProps
} from 'office-ui-fabric-react';
export interface ITrFormProps {
  mode: modes;
  tr?: TR;

  workTypes: Array<WorkType>;
  applicationTypes: Array<ApplicationType>;
  endUses: Array<EndUse>;
  save: (tr) => any; //make this return promise
  cancel: () => any;
  ensureUser: (email) => Promise<any>;
  peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
  TRsearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
  requestors:Array<User>;
  techSpecs:Array<User>;
  childTRs:Array<TR>;
}
