import { TR, ApplicationType, WorkType, EndUse, modes, User } from "../dataModel";
import {
  IPersonaProps
} from 'office-ui-fabric-react';
export interface ITrFormProps {
  mode: modes; // display , edit, new
  tr?: TR; // 
  workTypes: Array<WorkType>; // lookup column values
  applicationTypes: Array<ApplicationType>;// lookup column values
  endUses: Array<EndUse>;// lookup column values
  save: (tr) => any; //make this return promise // method to call to save tr
  cancel: () => any; // method to call to save tr
  ensureUser: (email) => Promise<any>; // method to call to ensure use is in the site
  peoplesearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>;
  TRsearch: (searchText: string, currentSelected: IPersonaProps[]) => Promise<IPersonaProps[]>; // method tyo call to searcgh gpr parenttr
  requestors: Array<User>; //lookup values for valid requestors on current site
  techSpecs: Array<User>; //lookup values for valid technical specialists on current site
  subTRs: Array<TR>; // child trs (trs nested under the main tr)
  getChildTr: (id: number) => Promise<Array<TR>>; // methid to call to cget child TRs if a user swicthes to a new TR

}
