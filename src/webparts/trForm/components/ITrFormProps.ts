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
  peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
}
