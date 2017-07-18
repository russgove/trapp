import { TRDocument, Test, PropertyTest, Pigment, Customer, TR, ApplicationType, WorkType, EndUse, modes, User } from "../dataModel";
import { ITRFormState } from "./ITRFormState";
  import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';

export interface ITrFormProps {

  //callbacks
  fetchDocumentWopiFrameURL: (id: number, mode: number) => Promise<string>;
  save: (tr: TR, originalAssignees: Array<number>, originalStatus: string) => Promise<any>; //make this return promise // method to call to save tr
  cancel: () => any; // method to call to save tr
  TRsearch: (searchText: string) => Promise<TR[]>; // method tyo call to searcgh gpr parenttr
  peopleSearch: (filter: string, selectedItems?: IPersonaProps[]) => IPersonaProps[] | PromiseLike<IPersonaProps[]>;
  fetchChildTr: (id: number) => Promise<Array<TR>>; // methid to call to cget child TRs if a user swicthes to a new TR
  fetchTR: (id: number) => Promise<TR>; // methid to call to cget child TRs if a user swicthes to a new TR
  uploadFile: (file: any, trId: number) => Promise<any>;
  getDocuments: (trId: number) => Promise<Array<TRDocument>>;
   ensureUsersInPersonas: (items?: Array<IPersonaProps>) => void;
  //data
  initialState: ITRFormState;
  mode: modes; // display , edit, new
  workTypes: Array<WorkType>; // lookup column values
  applicationTypes: Array<ApplicationType>;// lookup column values
  endUses: Array<EndUse>;// lookup column values
  requestors: Array<User>; //lookup values for valid requestors on current site
  techSpecs: Array<User>; //lookup values for valid technical specialists on current site
  customers: Array<Customer>;
  pigments: Array<Pigment>;
  tests: Array<Test>;
  propertyTests: Array<PropertyTest>;
  delayPriorToSettingCKEditor: number;
  ckeditorUrl:string;


}
