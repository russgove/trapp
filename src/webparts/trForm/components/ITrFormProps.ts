import { TR, ApplicationType, WorkType, EndUse,modes } from "../dataModel";
import {
  IPersonaProps
} from 'office-ui-fabric-react';
export interface ITrFormProps {
  mode:modes,
  tr?: TR;

  workTypes: Array<WorkType>;
  applicationTypes: Array<ApplicationType>;
  endUses: Array<EndUse>;
  save: (tr) => any;
  ensureUser: (email) => Promise<any>;
  peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
   TRsearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
}
