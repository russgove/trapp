import { TRDocument, Test, PropertyTest, Pigment, Customer, TR, ApplicationType, WorkType, EndUse, modes, User } from "../dataModel";
import * as md from "./MessageDisplay";
export  interface ITRFormState {
  tr: TR;
  childTRs: Array<TR>;
  errorMessages: Array<md.Message>;
  isDirty: boolean;
  showTRSearch: boolean;
 
}