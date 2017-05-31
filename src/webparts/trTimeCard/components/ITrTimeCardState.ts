import {TimeSpent} from "../dataModel";
export interface ITrTimeCardState {
   weekEndingDate:Date;
   timeSpents:Array<TimeSpent>;
   message:string;
  
}
