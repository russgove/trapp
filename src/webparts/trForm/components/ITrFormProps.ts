import {TR,ApplicationType,WorkType,EndUse} from "../dataModel";
export interface ITrFormProps {
  tr?: TR;
  workTypes: Array<WorkType>;
  applicationTypes: Array<ApplicationType>;
  endUses: Array<EndUse>;
  save:(tr)=>any;
}
